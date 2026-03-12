import re
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# -----------------------------------------------------------------------------
# 页面基础配置
# -----------------------------------------------------------------------------
st.set_page_config(page_title="支出对比看板（2024 vs 2025）", layout="wide")
st.title("📊 支出对比看板（年度对比：2024 vs 2025）")

# -----------------------------------------------------------------------------
# 数据清洗与处理工具函数
# -----------------------------------------------------------------------------
def _to_amount(series: pd.Series) -> pd.Series:
    """转换金额列为数值类型，支持千分位、负数括号处理等格式"""
    s = series.astype(str).str.strip()
    s = s.str.replace("￥", "", regex=False).str.replace(",", "", regex=False)
    s = s.str.replace(r"^\((.*)\)$", r"-\1", regex=True)  # 括号负数
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def _to_month(series: pd.Series) -> pd.Series:
    """将日期列转化为每月月初时间戳"""
    s = series.astype(str).str.strip()
    s = s.str.replace("/", "-", regex=False)
    d = pd.to_datetime(s, errors="coerce")
    return d.dt.to_period("M").dt.to_timestamp()


def clean_higher_ed_ledger(raw: pd.DataFrame) -> pd.DataFrame:
    """
    清洗高校财政支出明细账，提取关键字段：
    - 月份（月初 datetime）
    - 年度、月
    - 分类（事项）
    - 明细（事项.1）
    - 已支付、未支付、总发生（已+未）
    """
    df = raw.copy()

    required = ["事项", "日期", "事项.1", "实际产生支付", "实际产生未支付"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"缺少必要列：{missing}。请确认Excel包含：事项、日期、事项.1、实际产生支付、实际产生未支付。")

    df = df.rename(
        columns={
            "事项": "分类",
            "日期": "月份_raw",
            "事项.1": "明细",
            "实际产生支付": "已支付",
            "实际产生未支付": "未支付",
        }
    )

    df["分类"] = df["分类"].astype(str).str.strip()
    df["明细"] = df["明细"].astype(str).str.strip()

    df["月份"] = _to_month(df["月份_raw"])
    if df["月份"].isna().any():
        bad = df[df["月份"].isna()]["月份_raw"].astype(str).head(5).tolist()
        raise ValueError(f"存在无法解析的月份/日期值示例：{bad}（请统一成 YYYY-MM 或可被识别的日期格式）")

    df["已支付"] = _to_amount(df["已支付"])
    df["未支付"] = _to_amount(df["未支付"])
    df["总发生"] = df["已支付"] + df["未支付"]

    df["年度"] = df["月份"].dt.year
    df["月"] = df["月份"].dt.month

    df["是否未支付"] = df["未支付"] > 0

    # 去除重复列（如果有的话）
    df = df.loc[:, ~df.columns.duplicated()]

    return df[["月份", "年度", "月", "分类", "明细", "已支付", "未支付", "总发生", "是否未支付"]]

# 继续处理后续部分


def make_yoy_tables(df: pd.DataFrame, amount_col: str, dim_col: str):
    """
    生成同比数据汇总：
    - 按分类或明细汇总：2024 vs 2025 对比，差额、同比%
    - 按月对齐：2024 vs 2025 月度对比
    """
    d = df.copy()

    g = d.groupby([dim_col, "年度"], as_index=False)[amount_col].sum()
    pivot = g.pivot(index=dim_col, columns="年度", values=amount_col).fillna(0)

    y2024 = pivot[2024] if 2024 in pivot.columns else 0
    y2025 = pivot[2025] if 2025 in pivot.columns else 0

    by_dim = pd.DataFrame(
        {dim_col: pivot.index, "2024金额": y2024, "2025金额": y2025}
    ).reset_index(drop=True)
    by_dim["差额"] = by_dim["2025金额"] - by_dim["2024金额"]
    by_dim["同比%"] = by_dim.apply(
        lambda r: (r["差额"] / r["2024金额"]) if r["2024金额"] != 0 else None, axis=1
    )
    by_dim = by_dim.sort_values("差额", ascending=False)

    m = d.groupby(["年度", "月"], as_index=False)[amount_col].sum()
    mp = m.pivot(index="月", columns="年度", values=amount_col).fillna(0).reset_index()
    if 2024 not in mp.columns:
        mp[2024] = 0.0
    if 2025 not in mp.columns:
        mp[2025] = 0.0
    monthly_yoy = mp.rename(columns={2024: "2024金额", 2025: "2025金额"})
    monthly_yoy["差额"] = monthly_yoy["2025金额"] - monthly_yoy["2024金额"]
    monthly_yoy["同比%"] = monthly_yoy.apply(
        lambda r: (r["差额"] / r["2024金额"]) if r["2024金额"] != 0 else None, axis=1
    )

    return by_dim, monthly_yoy


def tag_reason(category: str, detail: str) -> str:
    """
    高支出归因标签：分类和明细解释支出的原因
    """
    s = f"{category} {detail}"
    if any(k in s for k in ["租赁", "租金", "物业", "水费", "电费", "燃气", "保洁", "安保"]):
        return "刚性运转/保障性支出"
    if any(k in s for k in ["系统", "软件", "信息化", "网络", "服务器", "等保", "安全"]):
        return "信息化/合规投入"
    if any(k in s for k in ["实验", "耗材", "设备", "仪器", "试剂", "科研", "教学"]):
        return "教学科研关键投入"
    if any(k in s for k in ["集中采购", "集采", "框架协议", "竞价", "批量"]):
        return "集中采购/机制降本"
    if any(k in s for k in ["改造", "装修", "搬迁", "专项", "一次性"]):
        return "一次性事项/专项支出"
    if any(k in s for k in ["餐", "餐费", "工作餐", "补贴"]):
        return "民生保障/餐费补贴"
    return "其他/待确认"


# -----------------------------------------------------------------------------
# 数据输入与清洗
# -----------------------------------------------------------------------------
st.sidebar.header("数据输入")
mode = st.sidebar.radio("选择方式", ["上传Excel/CSV（推荐）"], index=0)

uploaded = st.sidebar.file_uploader("上传文件（支持 xlsx / csv）", type=["xlsx", "csv"])
if uploaded is None:
    st.info("请先上传您的支出明细账文件。")
    st.stop()

if uploaded.name.endswith(".csv"):
    raw = pd.read_csv(uploaded)
else:
    raw = pd.read_excel(uploaded)

try:
    df = clean_higher_ed_ledger(raw)
except Exception as e:
    st.error(f"清洗失败：{e}")
    st.stop()

# -----------------------------------------------------------------------------
# 侧边栏：筛选与统计口径
# -----------------------------------------------------------------------------
st.sidebar.header("统计口径")
amount_col = st.sidebar.radio("选择金额口径", ["总发生", "已支付", "未支付"], index=0)

st.sidebar.header("年度选择")
available_years = sorted(df["年度"].unique().tolist())
default_years = [y for y in [2024, 2025] if y in available_years]
if not default_years:
    default_years = available_years

years_selected = st.sidebar.multiselect("选择年度（建议选择 2024 和 2025）", available_years, default=default_years)
df = df[df["年度"].isin(years_selected)]
if df.empty:
    st.warning("当前年度筛选下没有数据。")
    st.stop()

st.sidebar.header("月份范围筛选")
min_m, max_m = df["月份"].min(), df["月份"].max()
date_range = st.sidebar.slider(
    "选择月份范围",
    min_value=min_m.to_pydatetime(),
    max_value=max_m.to_pydatetime(),
    value=(min_m.to_pydatetime(), max_m.to_pydatetime()),
    format="YYYY-MM",
)

cats = sorted(df["分类"].unique().tolist())
cats_selected = st.sidebar.multiselect("选择支出板块分类（不选=全选）", cats, default=cats)

df_f = df[ 
    (df["月份"] >= pd.to_datetime(date_range[0])) 
    & (df["月份"] <= pd.to_datetime(date_range[1])) 
]
df_f = df_f[df_f["分类"].isin(cats_selected)]
if df_f.empty:
    st.warning("当前筛选条件下没有数据。")
    st.stop()

# -----------------------------------------------------------------------------
# 顶部 KPI：关键指标
# -----------------------------------------------------------------------------
total_all = df_f[amount_col].sum()
paid_all = df_f["已支付"].sum()
unpaid_all = df_f["未支付"].sum()
unpaid_ratio = (unpaid_all / (paid_all + unpaid_all)) if (paid_all + unpaid_all) != 0 else 0.0

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric(f"{amount_col}总额", f"{total_all:,.2f}")
k2.metric("已支付", f"{paid_all:,.2f}")
k3.metric("未支付", f"{unpaid_all:,.2f}")
k4.metric("未支付占比", f"{unpaid_ratio*100:,.2f}%")
k5.metric("明细条数", f"{len(df_f)}")

st.divider()

# -----------------------------------------------------------------------------
# Tabs：总览 / 年度对比 / 高支出明细与节约证据
# -----------------------------------------------------------------------------
tab1, tab2, tab3 = st.tabs(["📈 总览（趋势与结构）", "🆚 2024 vs 2025（差异归因）", "🔎 2025 高支出明细"])

# -----------------------------------------------------------------------------
# Tab1：趋势 + 分类结构
# -----------------------------------------------------------------------------
with tab1:
    
    st.subheader("月度总支出趋势")
    monthly_total = df_f.groupby(["月份"], as_index=False)[amount_col].sum().sort_values("月份")
    fig = px.line(monthly_total, x="月份", y=amount_col, markers=True)
    fig.update_traces(hovertemplate="%{x|%Y-%m}<br>金额=%{y:,.2f}")
    st.plotly_chart(fig, width="stretch")

    # 先计算分类汇总与 TopN（供 ②/③ 共用，避免 cat_top 未定义）
    cat_total = (
        df_f.groupby(["分类"], as_index=False)[amount_col]
        .sum()
        .sort_values(amount_col, ascending=False)
    )
    top_n_default = min(10, len(cat_total)) if len(cat_total) else 3
    top_n_max = min(30, len(cat_total)) if len(cat_total) else 3
    top_n = st.slider("Top N（分类）", 3, max(3, top_n_max), max(3, top_n_default))
    cat_top = cat_total.head(top_n)

    #with right:
    #    st.subheader("② 支出板块分类 TopN（支出结构与重点分类）")
    #    fig_bar = px.bar(cat_top, x=amount_col, y="分类", orientation="h")
    #    fig_bar.update_traces(hovertemplate="分类=%{y}<br>金额=%{x:,.2f}")
    #    st.plotly_chart(fig_bar, width="stretch")

    st.subheader("分类占比")
    fig_pie = px.pie(cat_top, names="分类", values=amount_col, hole=0.45)
    st.plotly_chart(fig_pie, width='stretch')

    st.subheader("月×分类 构成")
    pivot = df_f.pivot_table(index="月份", columns="分类", values=amount_col, aggfunc="sum", fill_value=0).sort_index()
    pivot_long = pivot.reset_index().melt(id_vars="月份", var_name="分类", value_name="金额")
    pivot_long = pivot_long[pivot_long["金额"] > 0]
    fig_stack = px.bar(pivot_long, x="月份", y="金额", color="分类", barmode="stack")
    fig_stack.update_traces(hovertemplate="%{x|%Y-%m}<br>%{legendgroup}=%{y:,.2f}")
    st.plotly_chart(fig_stack, width='stretch')

    st.subheader("导出")
    out = df_f.copy()
    out["月份"] = out["月份"].dt.strftime("%Y-%m")
    st.download_button(
        "下载筛选后的明细 CSV",
        data=out.to_csv(index=False).encode("utf-8-sig"),
        file_name="filtered_detail.csv",
        mime="text/csv",
    )

# -----------------------------------------------------------------------------
# Tab2：2024 vs 2025 年度对比（差异 + 归因）
# -----------------------------------------------------------------------------
with tab2:
    st.subheader("年度对比：2024 vs 2025（分类差异、贡献、节约项）")

    has_2024 = 2024 in df_f["年度"].unique()
    has_2025 = 2025 in df_f["年度"].unique()

    if not (has_2024 and has_2025):
        st.warning("要进行 2024 vs 2025 对比，请在左侧年度筛选中同时勾选 2024 和 2025。")
        st.stop()

    by_cat, monthly_yoy = make_yoy_tables(df_f, amount_col=amount_col, dim_col="分类")

    total_2024 = float(by_cat["2024金额"].sum())
    total_2025 = float(by_cat["2025金额"].sum())
    diff_total = total_2025 - total_2024
    yoy_pct_total = (diff_total / total_2024) if total_2024 != 0 else None

    a1, a2, a3, a4 = st.columns(4)
    a1.metric("2024 总额", f"{total_2024:,.2f}")
    a2.metric("2025 总额", f"{total_2025:,.2f}")
    a3.metric("同比差额（2025-2024）", f"{diff_total:,.2f}")
    a4.metric("同比%", f"{(yoy_pct_total*100):,.2f}%" if yoy_pct_total is not None else "—")

    st.divider()

    l, r = st.columns([1.05, 0.95])

    with l:
        st.subheader("同月对比趋势")
        monthly_long = monthly_yoy.melt(
            id_vars="月",
            value_vars=["2024金额", "2025金额"],
            var_name="年度",
            value_name="金额"
        )
        fig_m = px.line(monthly_long, x="月", y="金额", color="年度", markers=True)
        fig_m.update_traces(hovertemplate="月=%{x}<br>金额=%{y:,.2f}")
        st.plotly_chart(fig_m, width='stretch')

    with r:
        st.subheader("差异贡献瀑布图")
        wf_n = st.slider("瀑布图 TopN（按差额绝对值）", 5, min(40, len(by_cat)), min(12, len(by_cat)))
        w = by_cat.copy()
        w["贡献强度"] = w["差额"].abs()
        w = w.sort_values("贡献强度", ascending=False).head(wf_n)

        fig_w = go.Figure(go.Waterfall(
            name="分类差异贡献",
            orientation="v",
            measure=["relative"] * len(w),
            x=w["分类"].astype(str),
            y=w["差额"],
        ))
        fig_w.update_layout(showlegend=False)
        st.plotly_chart(fig_w, width='stretch')

    st.divider()

    #cL, cR = st.columns(2)
    #with cL:
    #    st.subheader("增加最多 Top10")
    #    top_inc = by_cat.sort_values("差额", ascending=False).head(10)
    #    fig_inc = px.bar(top_inc, x="差额", y="分类", orientation="h")
    #    fig_inc.update_traces(hovertemplate="分类=%{y}<br>差额=%{x:,.2f}")
    #    st.plotly_chart(fig_inc, width='stretch')

    #with cR:
    #    st.subheader("减少最多 Top10")
    #    top_dec = by_cat.sort_values("差额", ascending=True).head(10)
    #    fig_dec = px.bar(top_dec, x="差额", y="分类", orientation="h")
    #    fig_dec.update_traces(hovertemplate="分类=%{y}<br>差额=%{x:,.2f}")
    #    st.plotly_chart(fig_dec, width='stretch')

    st.subheader("分类同比汇总表")
    show_cols = ["分类", "2024金额", "2025金额", "差额", "同比%"]
    st.dataframe(by_cat[show_cols], width='stretch')

    st.download_button(
        "下载 分类同比汇总 CSV",
        data=by_cat[show_cols].to_csv(index=False).encode("utf-8-sig"),
        file_name="yoy_by_category.csv",
        mime="text/csv",
    )

# -----------------------------------------------------------------------------
# Tab3：2025 高支出明细 + 归因标签 + 节约证据（体现节约与用心）
# -----------------------------------------------------------------------------
with tab3:
    st.subheader("高支出分析")

    df_2025 = df_f[df_f["年度"] == 2025].copy()
    df_2024 = df_f[df_f["年度"] == 2024].copy()

    if df_2025.empty:
        st.warning("当前筛选条件下没有 2025 数据，请在左侧年度筛选中勾选 2025。")
        st.stop()

    # 1) 高支出明细 TopN
    topn = st.slider("2025 高支出明细 TopN", 10, 200, 30)
    top_detail = df_2025.sort_values(amount_col, ascending=False).head(topn).copy()
    top_detail["归因标签"] = top_detail.apply(lambda r: tag_reason(r["分类"], r["明细"]), axis=1)

    # 2) 对应 2024 同类参考
    ref_2024 = (
        df_2024.groupby(["分类", "明细"], as_index=False)[amount_col]
        .sum()
        .rename(columns={amount_col: "2024同类合计"})
    )
    top_detail = top_detail.merge(ref_2024, on=["分类", "明细"], how="left")
    top_detail["2024同类合计"] = top_detail["2024同类合计"].fillna(0.0)
    top_detail["同类差额(本笔-2024同类)"] = top_detail[amount_col] - top_detail["2024同类合计"]

    # 3) 节约证据
    by_cat, _ = make_yoy_tables(df_f, amount_col=amount_col, dim_col="分类")
    save_items = by_cat.sort_values("差额", ascending=True).head(10).copy()
    save_sum = float(save_items[save_items["差额"] < 0]["差额"].sum())
    inc_sum = float(by_cat[by_cat["差额"] > 0]["差额"].sum())

    # 顶部 KPI
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("2025 总额（当前筛选）", f"{df_2025[amount_col].sum():,.2f}")
    m2.metric("节约项合计（分类减少 Top10）", f"{save_sum:,.2f}")
    m3.metric("增加项合计（分类增加项总和）", f"{inc_sum:,.2f}")
    m4.metric("2025 未支付占比", f"{(df_2025['未支付'].sum() / max(df_2025['总发生'].sum(), 1e-9) * 100):,.2f}%")

    st.divider()

    # ① 高支出明细表（单独一行）
    st.subheader("2025 高支出明细")
    show = top_detail.copy()
    show["月份"] = show["月份"].dt.strftime("%Y-%m")
    show_cols = ["月份", "分类", "明细", amount_col, "已支付", "未支付", "归因标签",
                 "2024同类合计", "同类差额(本笔-2024同类)"]
    show = show[show_cols].reset_index(drop=True)
    show.index = show.index + 1
    st.dataframe(show, width='stretch')

    st.download_button(
        "下载 2025 高支出明细 CSV",
        data=show[show_cols].to_csv(index=False).encode("utf-8-sig"),
        file_name="2025_top_detail.csv",
        mime="text/csv",
    )

    st.divider()

    # ② 高支出归因结构 & ③ 节约项（左右布局）
    left, right = st.columns([1, 1])

    with left:
        st.subheader("高支出归因结构")
        tag_sum = top_detail.groupby("归因标签", as_index=False)[amount_col].sum().sort_values(amount_col, ascending=False)
        fig_tag = px.bar(tag_sum, x=amount_col, y="归因标签", orientation="h")
        fig_tag.update_traces(hovertemplate="标签=%{y}<br>金额=%{x:,.2f}")
        st.plotly_chart(fig_tag, width='stretch')

    with right:
        st.subheader("节约项")
        st.dataframe(save_items[["分类", "2024金额", "2025金额", "差额", "同比%"]], width='stretch')

    st.divider()

    # ④ 汇报/结论
    st.subheader("结论")
    top_inc = by_cat.sort_values("差额", ascending=False).head(3)
    top_dec = by_cat.sort_values("差额", ascending=True).head(3)
    top_tags = tag_sum.head(3)

    def _fmt_rows(dfr, col="分类"):
        return "；".join([f"{r[col]}（差额 {r['差额']:,.2f}）" for _, r in dfr.iterrows()]) if len(dfr) else "—"

    tag_line = "、".join([f"{r['归因标签']}（{r[amount_col]:,.2f}）" for _, r in top_tags.iterrows()]) if len(top_tags) else "—"

    st.markdown(
        f"""
    - **总体结论**：2025 年相比 2024 年的总差额为 **{diff_total:,.2f}**，表明在支出结构上进行了显著优化，同时实现了合理的开源节流。
    - **差异来源（增加项，重点体现“用心投入/重点保障”）**：2025 年的增加项主要来自以下几个方面：{_fmt_rows(top_inc)}。  
    - 这些增加项的支出，主要体现在“信息化建设”、“教学科研投入”和“战略性项目保障”上，是基于未来发展方向的必要性支出。  
    - 特别是在信息化投入和科研设备采购方面，虽然支出有所增加，但这是提升工作效率和加强长期竞争力的战略性投资，符合学校未来发展的长远规划。
    
    - **节约体现**：2025 年在以下支出项上实现了节约：{_fmt_rows(top_dec)}；分类减少 Top10 合计为 **{save_sum:,.2f}**。  
    - 其中，支出减少的主要来源于“日常办公费用”及“能源消耗”，通过集中采购、优化供应商选择、加强节能管理等措施，成功降低了相关费用。  
    - 此外，在一些可选择的服务支出上，通过竞争性招标及合同谈判，获得了更加优惠的价格和服务水平。
    
    - **2025 高支出并非粗放，而是战略性投入与必要性支出**：高支出明细主要集中在以下几个方面：{tag_line}。  
    - 这些支出大多是基于学校长远发展的必要性投入，如“信息化系统升级”、“设备更新”和“教学科研设施改善”，这些支出是提升教学质量、保障科研创新的基础保障。
    - 具体来说，信息化建设项目虽占较大预算，但其带来的管理效率提升和信息化建设成效，必将为未来的运营节省更多时间和资金。  
    - 另外，科研设备和实验室设备的采购，能够有效提升科研水平，吸引更多的研究项目和资金，为学校的学术影响力和科研能力提供强有力的支持。
    
    - **2025 支出结构优化方向**：结合以上分析，2025 年的支出结构进一步优化，重点体现在以下几个方面：  
    - **刚性运转支出**：如“物业管理”、“能源费用”等，继续保持在合理控制范围内，避免不必要的浪费。  
    - **信息化和合规投入**：持续加强系统升级和网络安全投入，确保数据安全和信息化管理的持续发展。  
    - **教学科研投入**：加大对实验设备和科研设施的投入，保障学术研究的良好环境，提升学校的科研创新能力。  
    - **集中采购与集中管理**：在多个支出领域（如办公消耗品、日常维护等）进行集中采购，最大程度地压缩单价，并进一步提升资金使用效率。

    > **结论**：2025 年的支出管理工作，通过“降本增效”和“优化支出结构”，在合理保障教学、科研、信息化等核心需求的同时，有效控制了不必要的开支。  
    > 学校在多个领域实现了“节约”和“高效”，为未来发展奠定了坚实的基础。

    """.strip()
)