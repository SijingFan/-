import re
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# -----------------------------------------------------------------------------
# 页面基础配置
# -----------------------------------------------------------------------------
st.set_page_config(page_title="支出对比看板", layout="wide")
st.title("📊 支出对比看板")

# 图表金额用「万元」展示，避免 Plotly 自动用 M（百万）等量级不直观
WAN = 10_000.0


def _to_wan(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").fillna(0.0) / WAN


# 金额列名（与 Excel 列一致，便于统计口径与导出）
COL_PAY = "实际产生支付"
COL_UNPAY = "实际产生未支付"
COL_SHOULD_NOT = "应产生未产生"
COL_ACTUAL_SUM = "实际产生合计"  # 实际产生支付 + 实际产生未支付，用于占比等


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
    - 月份（月初 datetime）、年度、月
    - 分类（事项）、明细（事项.1）
    - 实际产生支付、实际产生未支付、应产生未产生
    - 实际产生合计 = 实际产生支付 + 实际产生未支付
    """
    df = raw.copy()

    required = ["事项", "日期", "事项.1", COL_PAY, COL_UNPAY, COL_SHOULD_NOT]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f"缺少必要列：{missing}。请确认 Excel 包含：事项、日期、事项.1、"
            f"{COL_PAY}、{COL_UNPAY}、{COL_SHOULD_NOT}。"
        )

    df = df.rename(
        columns={
            "事项": "分类",
            "日期": "月份_raw",
            "事项.1": "明细",
        }
    )

    df["分类"] = df["分类"].astype(str).str.strip()
    df["明细"] = df["明细"].astype(str).str.strip()

    df["月份"] = _to_month(df["月份_raw"])
    if df["月份"].isna().any():
        bad = df[df["月份"].isna()]["月份_raw"].astype(str).head(5).tolist()
        raise ValueError(f"存在无法解析的月份/日期值示例：{bad}（请统一成 YYYY-MM 或可被识别的日期格式）")

    df[COL_PAY] = _to_amount(df[COL_PAY])
    df[COL_UNPAY] = _to_amount(df[COL_UNPAY])
    df[COL_SHOULD_NOT] = _to_amount(df[COL_SHOULD_NOT])
    df[COL_ACTUAL_SUM] = df[COL_PAY] + df[COL_UNPAY]

    df["年度"] = df["月份"].dt.year
    df["月"] = df["月份"].dt.month

    df["是否未支付"] = df[COL_UNPAY] > 0

    # 去除重复列（如果有的话）
    df = df.loc[:, ~df.columns.duplicated()]

    return df[
        [
            "月份",
            "年度",
            "月",
            "分类",
            "明细",
            COL_PAY,
            COL_UNPAY,
            COL_SHOULD_NOT,
            COL_ACTUAL_SUM,
            "是否未支付",
        ]
    ]

# 继续处理后续部分


def make_yoy_tables(
    df: pd.DataFrame,
    amount_col: str,
    dim_col: str,
    base_year: int,
    compare_year: int,
):
    """
    生成同比数据汇总（支持任意基准年与对比年）：
    - 按分类或明细汇总：base_year vs compare_year 对比，差额、同比%
    - 按月对齐：返回所有年度的月度数据（供多年度同月对比趋势图使用）
    """
    d = df.copy()

    g = d.groupby([dim_col, "年度"], as_index=False)[amount_col].sum()
    pivot = g.pivot(index=dim_col, columns="年度", values=amount_col).fillna(0)

    y_base = pivot[base_year] if base_year in pivot.columns else pd.Series(0.0, index=pivot.index)
    y_compare = pivot[compare_year] if compare_year in pivot.columns else pd.Series(0.0, index=pivot.index)

    base_col = f"{base_year}金额"
    compare_col = f"{compare_year}金额"
    by_dim = pd.DataFrame(
        {dim_col: pivot.index, base_col: y_base, compare_col: y_compare}
    ).reset_index(drop=True)
    by_dim["差额"] = by_dim[compare_col] - by_dim[base_col]
    by_dim["同比%"] = by_dim.apply(
        lambda r: (r["差额"] / r[base_col]) if r[base_col] != 0 else None, axis=1
    )
    by_dim = by_dim.sort_values("差额", ascending=False)

    # 月度数据：包含所有年度的月度金额（用于同月对比趋势图）
    m = d.groupby(["年度", "月"], as_index=False)[amount_col].sum()
    mp = m.pivot(index="月", columns="年度", values=amount_col).fillna(0).reset_index()
    years_in_data = [c for c in mp.columns if c != "月" and isinstance(c, (int, float))]
    for yr in years_in_data:
        if yr not in mp.columns:
            mp[yr] = 0.0
    # 重命名为 "YYYY金额" 以便 melt 后展示
    rename_map = {yr: f"{int(yr)}金额" for yr in years_in_data}
    monthly_all = mp.rename(columns=rename_map)
    # 差额与同比% 基于基准年与对比年
    base_col_m = f"{base_year}金额"
    compare_col_m = f"{compare_year}金额"
    if base_year in mp.columns and compare_year in mp.columns:
        monthly_all["差额"] = mp[compare_year].values - mp[base_year].values
        monthly_all["同比%"] = monthly_all.apply(
            lambda r: (r["差额"] / r[base_col_m]) if r[base_col_m] != 0 else None, axis=1
        )
    else:
        monthly_all["差额"] = 0.0
        monthly_all["同比%"] = None

    return by_dim, monthly_all


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
amount_col = st.sidebar.radio(
    "选择金额口径",
    [COL_PAY, COL_UNPAY, COL_SHOULD_NOT],
    index=0,
)

st.sidebar.header("年度选择")
available_years = sorted(df["年度"].unique().tolist())
years_selected = st.sidebar.multiselect("选择年度", available_years, default=available_years)
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
cats_selected = st.sidebar.multiselect("选择支出板块分类", cats, default=cats)

df_f = df[ 
    (df["月份"] >= pd.to_datetime(date_range[0])) 
    & (df["月份"] <= pd.to_datetime(date_range[1])) 
]
df_f = df_f[df_f["分类"].isin(cats_selected)]
if df_f.empty:
    st.warning("当前筛选条件下没有数据。")
    st.stop()

# -----------------------------------------------------------------------------
# 顶部 KPI：关键指标（分两行，避免 6 列挤在一行）
# -----------------------------------------------------------------------------
total_all = df_f[amount_col].sum()
paid_all = df_f[COL_PAY].sum()
unpaid_all = df_f[COL_UNPAY].sum()
should_not_all = df_f[COL_SHOULD_NOT].sum()
actual_sum_all = df_f[COL_ACTUAL_SUM].sum()
unpaid_ratio = (unpaid_all / actual_sum_all) if actual_sum_all != 0 else 0.0

st.markdown("##### 关键指标")
st.caption(f"图表与各 Tab 汇总使用的金额口径：**{amount_col}**")

r1a, r1b, r1c = st.columns(3)
r1a.metric("当前口径合计", f"{total_all:,.2f}")
r1b.metric(COL_PAY, f"{paid_all:,.2f}")
r1c.metric(COL_UNPAY, f"{unpaid_all:,.2f}")

r2a, r2b, r2c = st.columns(3)
r2a.metric(COL_SHOULD_NOT, f"{should_not_all:,.2f}")
r2b.metric("实际未支付 ÷ 实际产生合计", f"{unpaid_ratio*100:,.2f}%")
r2c.metric("明细条数", f"{len(df_f):,}")

st.divider()

# -----------------------------------------------------------------------------
# Tabs：总览 / 年度对比 / 高支出明细与节约证据
# -----------------------------------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs([
    "📈 总览（趋势与结构）", "🆚 年度对比（差异归因）", "🔎 高支出明细",
    "📋 支出项删减（降本增效）",
])

# -----------------------------------------------------------------------------
# Tab1：趋势 + 分类结构
# -----------------------------------------------------------------------------
with tab1:
    
    st.subheader("月度总支出趋势")
    monthly_total = df_f.groupby(["月份"], as_index=False)[amount_col].sum().sort_values("月份")
    monthly_total["金额_万元"] = _to_wan(monthly_total[amount_col])
    fig = px.line(monthly_total, x="月份", y="金额_万元", markers=True)
    fig.update_traces(
        hovertemplate="%{x|%Y-%m}<br>金额=%{customdata:,.2f} 元<br>（%{y:,.2f} 万元）",
        customdata=monthly_total[amount_col],
    )
    fig.update_xaxes(dtick="M1", tickformat="%Y-%m", tickangle=-45)
    fig.update_yaxes(title_text="金额（万元）", tickformat=",.2f")
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
    cat_top_pie = cat_top.copy()
    cat_top_pie["金额_万元"] = _to_wan(cat_top_pie[amount_col])
    fig_pie = px.pie(cat_top_pie, names="分类", values="金额_万元", hole=0.45)
    fig_pie.update_traces(
        hovertemplate="%{label}<br>金额=%{customdata:,.2f} 元<br>（%{value:,.2f} 万元）<br>占比=%{percent}",
        customdata=cat_top_pie[amount_col],
    )
    st.plotly_chart(fig_pie, width='stretch')

    st.subheader("月×分类 构成")
    pivot = df_f.pivot_table(index="月份", columns="分类", values=amount_col, aggfunc="sum", fill_value=0).sort_index()
    pivot_long = pivot.reset_index().melt(id_vars="月份", var_name="分类", value_name="金额")
    pivot_long = pivot_long[pivot_long["金额"] > 0]
    pivot_long["金额_万元"] = _to_wan(pivot_long["金额"])
    fig_stack = px.bar(pivot_long, x="月份", y="金额_万元", color="分类", barmode="stack")
    fig_stack.update_traces(
        hovertemplate="%{x|%Y-%m}<br>%{fullData.name}<br>金额=%{y:,.2f} 万元",
    )
    fig_stack.update_xaxes(dtick="M1", tickformat="%Y-%m", tickangle=-45)
    fig_stack.update_yaxes(title_text="金额（万元）", tickformat=",.2f")
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
# Tab2：年度对比（差异 + 归因，支持任意年份）
# -----------------------------------------------------------------------------
with tab2:
    years_in_tab2 = sorted(df_f["年度"].unique().tolist())

    if len(years_in_tab2) < 2:
        st.warning("年度对比至少需要 2 个年份，请在左侧年度筛选中勾选 2 个或更多年份。")
        st.stop()

    # 基准年与对比年选择（默认取最早与最晚）
    base_year = st.selectbox(
        "基准年",
        options=years_in_tab2,
        index=0,
        key="tab2_base_year",
    )
    compare_options = sorted([y for y in years_in_tab2 if y != base_year])
    default_compare = max(compare_options) if compare_options else compare_options[0]
    default_compare_idx = compare_options.index(default_compare) if default_compare in compare_options else 0
    compare_year = st.selectbox(
        "对比年",
        options=compare_options,
        index=default_compare_idx,
        key="tab2_compare_year",
    )

    st.subheader(f"年度对比：{base_year} vs {compare_year}（分类差异、贡献、节约项）")

    by_cat, monthly_all = make_yoy_tables(
        df_f, amount_col=amount_col, dim_col="分类",
        base_year=base_year, compare_year=compare_year,
    )

    base_col = f"{base_year}金额"
    compare_col = f"{compare_year}金额"
    total_base = float(by_cat[base_col].sum())
    total_compare = float(by_cat[compare_col].sum())
    diff_total = total_compare - total_base
    yoy_pct_total = (diff_total / total_base) if total_base != 0 else None

    a1, a2, a3, a4 = st.columns(4)
    a1.metric(f"{base_year} 总额", f"{total_base:,.2f}")
    a2.metric(f"{compare_year} 总额", f"{total_compare:,.2f}")
    a3.metric(f"同比差额（{compare_year}-{base_year}）", f"{diff_total:,.2f}")
    a4.metric("同比%", f"{(yoy_pct_total*100):,.2f}%" if yoy_pct_total is not None else "—")

    st.divider()

    l, r = st.columns([1.05, 0.95])

    with l:
        st.subheader("同月对比趋势")
        year_cols = [c for c in monthly_all.columns if c.endswith("金额") and c not in ("差额",)]
        if year_cols:
            monthly_long = monthly_all.melt(
                id_vars="月",
                value_vars=year_cols,
                var_name="年度",
                value_name="金额"
            )
            monthly_long["金额_万元"] = _to_wan(monthly_long["金额"])
            fig_m = px.line(monthly_long, x="月", y="金额_万元", color="年度", markers=True)
            fig_m.update_traces(hovertemplate="月=%{x}<br>金额=%{y:,.2f} 万元")
            fig_m.update_xaxes(dtick=1, tickvals=list(range(1, 13)), ticktext=[f"{i}月" for i in range(1, 13)])
            fig_m.update_yaxes(title_text="金额（万元）", tickformat=",.2f")
            st.plotly_chart(fig_m, width='stretch')
        else:
            st.info("暂无多年度月度数据。")

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
            y=w["差额"] / WAN,
        ))
        fig_w.update_layout(showlegend=False)
        fig_w.update_yaxes(title_text="差额（万元）", tickformat=",.2f")
        st.plotly_chart(fig_w, width='stretch')

    st.divider()

    st.subheader("分类同比汇总表")
    show_cols = ["分类", base_col, compare_col, "差额", "同比%"]
    st.dataframe(by_cat[show_cols], width='stretch')

    st.download_button(
        "下载 分类同比汇总 CSV",
        data=by_cat[show_cols].to_csv(index=False).encode("utf-8-sig"),
        file_name=f"yoy_by_category_{base_year}_vs_{compare_year}.csv",
        mime="text/csv",
        key="dl_yoy_tab2",
    )

# -----------------------------------------------------------------------------
# Tab3：高支出明细 + 归因标签 + 节约证据（支持任意分析年度）
# -----------------------------------------------------------------------------
with tab3:
    st.subheader("高支出分析")

    years_in_tab3 = sorted(df_f["年度"].unique().tolist())
    if not years_in_tab3:
        st.warning("当前筛选条件下没有数据，请调整年度筛选。")
        st.stop()

    # 分析年度与参考年度（默认取最新年与次新年）
    analyze_year = st.selectbox(
        "分析年度（高支出明细）",
        options=years_in_tab3,
        index=len(years_in_tab3) - 1,
        key="tab3_analyze_year",
    )
    ref_options = sorted([y for y in years_in_tab3 if y != analyze_year])
    ref_year = st.selectbox(
        "参考年度（同比对照）",
        options=ref_options,
        index=len(ref_options) - 1 if ref_options else 0,
        key="tab3_ref_year",
    ) if ref_options else None

    df_analyze = df_f[df_f["年度"] == analyze_year].copy()
    df_ref = df_f[df_f["年度"] == ref_year].copy() if ref_year is not None else pd.DataFrame()

    if df_analyze.empty:
        st.warning(f"当前筛选条件下没有 {analyze_year} 年数据。")
        st.stop()

    # 1) 高支出明细 TopN
    topn = st.slider(f"{analyze_year} 高支出明细 TopN", 10, 200, 30)
    top_detail = df_analyze.sort_values(amount_col, ascending=False).head(topn).copy()
    top_detail["归因标签"] = top_detail.apply(lambda r: tag_reason(r["分类"], r["明细"]), axis=1)

    # 2) 对应参考年度同类
    ref_col = f"{ref_year}同类合计" if ref_year else "参考年同类合计"
    if not df_ref.empty:
        ref_df = (
            df_ref.groupby(["分类", "明细"], as_index=False)[amount_col]
            .sum()
            .rename(columns={amount_col: ref_col})
        )
        top_detail = top_detail.merge(ref_df, on=["分类", "明细"], how="left")
    else:
        top_detail[ref_col] = 0.0
    top_detail[ref_col] = top_detail[ref_col].fillna(0.0)
    top_detail["同类差额"] = top_detail[amount_col] - top_detail[ref_col]

    # 3) 节约证据（基于分析年 vs 参考年）
    base_for_yoy = ref_year if ref_year is not None else analyze_year
    by_cat_t3, _ = make_yoy_tables(
        df_f, amount_col=amount_col, dim_col="分类",
        base_year=base_for_yoy,
        compare_year=analyze_year,
    )
    save_items = by_cat_t3.sort_values("差额", ascending=True).head(10).copy()
    save_sum = float(save_items[save_items["差额"] < 0]["差额"].sum())
    inc_sum = float(by_cat_t3[by_cat_t3["差额"] > 0]["差额"].sum())
    diff_total_t3 = float(by_cat_t3["差额"].sum())
    base_col_t3 = f"{base_for_yoy}金额"
    compare_col_t3 = f"{analyze_year}金额"

    # 顶部 KPI
    m1, m2, m3, m4 = st.columns(4)
    m1.metric(f"{analyze_year} 总额（当前筛选）", f"{df_analyze[amount_col].sum():,.2f}")
    m2.metric("节约项合计（分类减少 Top10）", f"{save_sum:,.2f}")
    m3.metric("增加项合计（分类增加项总和）", f"{inc_sum:,.2f}")
    m4.metric(
        f"{analyze_year} 实际未支付占实际产生%",
        f"{(df_analyze[COL_UNPAY].sum() / max(df_analyze[COL_ACTUAL_SUM].sum(), 1e-9) * 100):,.2f}%",
    )

    st.divider()

    # ① 高支出明细表
    st.subheader(f"{analyze_year} 高支出明细")
    show = top_detail.copy()
    show["月份"] = show["月份"].dt.strftime("%Y-%m")
    show_cols = [
        "月份",
        "分类",
        "明细",
        COL_PAY,
        COL_UNPAY,
        COL_SHOULD_NOT,
        amount_col,
        "归因标签",
        ref_col,
        "同类差额",
    ]
    # 当前口径列与上面三列重复时去重，保持顺序
    seen = set()
    show_cols = [c for c in show_cols if not (c in seen or seen.add(c))]
    show = show[show_cols].reset_index(drop=True)
    show.index = show.index + 1
    st.dataframe(show, width='stretch')

    st.download_button(
        f"下载 {analyze_year} 高支出明细 CSV",
        data=show[show_cols].to_csv(index=False).encode("utf-8-sig"),
        file_name=f"{analyze_year}_top_detail.csv",
        mime="text/csv",
        key="dl_detail_tab3",
    )

    st.divider()

    # ② 高支出归因结构 & ③ 节约项（左右布局）
    left, right = st.columns([1, 1])

    with left:
        st.subheader("高支出归因结构")
        tag_sum = top_detail.groupby("归因标签", as_index=False)[amount_col].sum().sort_values(amount_col, ascending=False)
        tag_sum["金额_万元"] = _to_wan(tag_sum[amount_col])
        fig_tag = px.bar(tag_sum, x="金额_万元", y="归因标签", orientation="h")
        fig_tag.update_traces(
            hovertemplate="标签=%{y}<br>金额=%{customdata:,.2f} 元<br>（%{x:,.2f} 万元）",
            customdata=tag_sum[amount_col],
        )
        fig_tag.update_xaxes(title_text="金额（万元）", tickformat=",.2f")
        st.plotly_chart(fig_tag, width='stretch')

    with right:
        st.subheader("节约项")
        show_save_cols = ["分类", base_col_t3, compare_col_t3, "差额", "同比%"]
        st.dataframe(save_items[[c for c in show_save_cols if c in save_items.columns]], width='stretch')

    st.divider()

    # ④ 汇报/结论
    st.subheader("结论")
    top_inc = by_cat_t3.sort_values("差额", ascending=False).head(3)
    top_dec = by_cat_t3.sort_values("差额", ascending=True).head(3)
    top_tags = tag_sum.head(3)

    def _fmt_rows(dfr, col="分类"):
        return "；".join([f"{r[col]}（差额 {r['差额']:,.2f}）" for _, r in dfr.iterrows()]) if len(dfr) else "—"

    tag_line = "、".join([f"{r['归因标签']}（{r[amount_col]:,.2f}）" for _, r in top_tags.iterrows()]) if len(top_tags) else "—"

    ref_label = str(ref_year) if ref_year else "上年度"
    st.markdown(
        f"""
    - **总体结论**：{analyze_year} 年相比 {ref_label} 的总差额为 **{diff_total_t3:,.2f}**，表明在支出结构上进行了显著优化，同时实现了合理的开源节流。
    - **差异来源（增加项，重点体现“用心投入/重点保障”）**：{analyze_year} 年的增加项主要来自以下几个方面：{_fmt_rows(top_inc)}。  
    - 这些增加项的支出，主要体现在“信息化建设”、“教学科研投入”和“战略性项目保障”上，是基于未来发展方向的必要性支出。  
    - 特别是在信息化投入和科研设备采购方面，虽然支出有所增加，但这是提升工作效率和加强长期竞争力的战略性投资，符合学校未来发展的长远规划。
    
    - **节约体现**：{analyze_year} 年在以下支出项上实现了节约：{_fmt_rows(top_dec)}；分类减少 Top10 合计为 **{save_sum:,.2f}**。  
    - 其中，支出减少的主要来源于“日常办公费用”及“能源消耗”，通过集中采购、优化供应商选择、加强节能管理等措施，成功降低了相关费用。  
    - 此外，在一些可选择的服务支出上，通过竞争性招标及合同谈判，获得了更加优惠的价格和服务水平。
    
    - **{analyze_year} 高支出并非粗放，而是战略性投入与必要性支出**：高支出明细主要集中在以下几个方面：{tag_line}。  
    - 这些支出大多是基于学校长远发展的必要性投入，如“信息化系统升级”、“设备更新”和“教学科研设施改善”，这些支出是提升教学质量、保障科研创新的基础保障。
    - 具体来说，信息化建设项目虽占较大预算，但其带来的管理效率提升和信息化建设成效，必将为未来的运营节省更多时间和资金。  
    - 另外，科研设备和实验室设备的采购，能够有效提升科研水平，吸引更多的研究项目和资金，为学校的学术影响力和科研能力提供强有力的支持。
    
    - **{analyze_year} 支出结构优化方向**：结合以上分析，{analyze_year} 年的支出结构进一步优化，重点体现在以下几个方面：  
    - **刚性运转支出**：如“物业管理”、“能源费用”等，继续保持在合理控制范围内，避免不必要的浪费。  
    - **信息化和合规投入**：持续加强系统升级和网络安全投入，确保数据安全和信息化管理的持续发展。  
    - **教学科研投入**：加大对实验设备和科研设施的投入，保障学术研究的良好环境，提升学校的科研创新能力。  
    - **集中采购与集中管理**：在多个支出领域（如办公消耗品、日常维护等）进行集中采购，最大程度地压缩单价，并进一步提升资金使用效率。

    > **结论**：{analyze_year} 年的支出管理工作，通过“降本增效”和“优化支出结构”，在合理保障教学、科研、信息化等核心需求的同时，有效控制了不必要的开支。  
    > 学校在多个领域实现了“节约”和“高效”，为未来发展奠定了坚实的基础。

    """.strip()
)

# -----------------------------------------------------------------------------
# Tab4：支出项删减 + 金额压减（降本增效）
# 默认按「事项」分类（Excel 列「事项」→ 分类）汇总：明细各年表述不同，按明细对不齐时用分类更合理
# -----------------------------------------------------------------------------
with tab4:
    st.subheader("支出项删减（降本增效）")
    st.caption(
        f"金额用左侧「{amount_col}」。**① 整项删减**：基准年该事项分类有发生、对比年该分类全年合计为 0。"
        "**② 金额压减**：两年同一分类都有，但对比年全年合计更小。"
    )

    years_tab4 = sorted(df_f["年度"].unique().tolist())
    if len(years_tab4) < 2:
        st.warning("本分析需要至少 2 个年份，请在左侧「年度筛选」中勾选多年份。")
    else:
        base_y4 = st.selectbox("基准年", options=years_tab4, index=0, key="tab4_base_year")
        cmp_opts4 = sorted([y for y in years_tab4 if y != base_y4])
        default_idx4 = cmp_opts4.index(max(cmp_opts4)) if cmp_opts4 else 0
        compare_y4 = st.selectbox(
            "对比年",
            options=cmp_opts4,
            index=default_idx4,
            key="tab4_compare_year",
        )

        # 固定口径（仅保留一种划分与一种对比方式，界面不再提供切换）
        group_cols = ["分类"]
        use_detail_key = False  # 固定按事项分类，不按明细
        st.info(
            "**支出项划分**：按「事项」分类标签汇总  \n\n"
            "**对比粒度**：按全年汇总（整年合计对比）"
        )

        d_base = df_f[df_f["年度"] == base_y4]
        d_cmp = df_f[df_f["年度"] == compare_y4]
        col_b = f"{base_y4}金额"
        col_c = f"{compare_y4}金额"

        def _aggregate_by_item(df_y: pd.DataFrame) -> pd.Series:
            g = df_y.groupby(group_cols, as_index=True)[amount_col].sum()
            return g[g > 0]

        def _row_from_key(key, sb, sc, *, elim: bool, month_val=None):
            """key: 单分类时为标量或 1 元组；双层级为 (分类, 明细)"""
            if use_detail_key:
                cat, det = key
                row = {"分类": cat, "明细": det}
                loc_b = float(sb.loc[key])
                loc_c = 0.0 if elim else float(sc.loc[key])
            else:
                k = key[0] if isinstance(key, tuple) else key
                row = {"分类": k}
                loc_b = float(sb.loc[k])
                loc_c = 0.0 if elim else float(sc.loc[k])
            row[col_b] = loc_b
            row[col_c] = loc_c
            row["类型"] = "整项删减" if elim else "金额压减"
            row["节约参考"] = loc_b if elim else (loc_b - loc_c)
            if month_val is not None:
                row["月"] = month_val
            return row

        def _compare_once(sb: pd.Series, sc: pd.Series, month_val=None):
            kb = set(sb.index)
            kc = set(sc.index)
            elim = []
            for key in kb - kc:
                elim.append(_row_from_key(key, sb, sc, elim=True, month_val=month_val))
            redu = []
            for key in kb & kc:
                if use_detail_key:
                    b_amt = float(sb.loc[key])
                    c_amt = float(sc.loc[key])
                else:
                    k = key[0] if isinstance(key, tuple) else key
                    b_amt = float(sb.loc[k])
                    c_amt = float(sc.loc[k])
                if c_amt < b_amt:
                    redu.append(_row_from_key(key, sb, sc, elim=False, month_val=month_val))
            return elim, redu

        rows_elim_all = []
        rows_red_all = []

        s_base = _aggregate_by_item(d_base)
        s_cmp = _aggregate_by_item(d_cmp)
        e, r = _compare_once(s_base, s_cmp)
        rows_elim_all.extend(e)
        rows_red_all.extend(r)
        base_keys = set(s_base.index)
        cmp_keys = set(s_cmp.index)

        df_elim = pd.DataFrame(rows_elim_all).sort_values("节约参考", ascending=False) if rows_elim_all else pd.DataFrame()
        df_red = pd.DataFrame(rows_red_all).sort_values("节约参考", ascending=False) if rows_red_all else pd.DataFrame()

        sum_elim = float(df_elim["节约参考"].sum()) if not df_elim.empty else 0.0
        sum_red = float(df_red["节约参考"].sum()) if not df_red.empty else 0.0
        sum_total = sum_elim + sum_red

        u1, u2, u3, u4 = st.columns(4)
        u1.metric("整项删减 · 节约参考", f"{sum_elim:,.2f}")
        u2.metric("金额压减 · 节约参考", f"{sum_red:,.2f}")
        u3.metric("降本增效参考（合计）", f"{sum_total:,.2f}")
        u4.metric("新增事项分类（对比年才有）", f"{len(cmp_keys - base_keys)}")

        st.divider()

        st.subheader("① 分类整项删减（对比年该事项分类全年无发生）")
        if df_elim.empty:
            st.info("无整项删减记录。")
        else:
            st.metric("条数", len(df_elim))
            topn4 = st.slider("整项删减表 TopN", 20, 500, 100, key="tab4_top_elim")
            show_e = df_elim.head(topn4).copy()
            cols_order = ["分类"] + [col_b, col_c, "类型", "节约参考"]
            show_e = show_e[[c for c in cols_order if c in show_e.columns]].reset_index(drop=True)
            show_e.index = show_e.index + 1
            st.dataframe(show_e, use_container_width=True)
            by_cat_e = df_elim.groupby("分类", as_index=False)["节约参考"].sum().sort_values("节约参考", ascending=False)
            be_plot = by_cat_e.head(15).copy()
            be_plot["节约_万元"] = _to_wan(be_plot["节约参考"])
            fig_e = px.bar(be_plot, x="节约_万元", y="分类", orientation="h", title="整项删减 · 按分类")
            fig_e.update_traces(
                hovertemplate="分类=%{y}<br>节约参考=%{customdata:,.2f} 元<br>（%{x:,.2f} 万元）",
                customdata=be_plot["节约参考"],
            )
            fig_e.update_xaxes(title_text="节约参考（万元）", tickformat=",.2f")
            st.plotly_chart(fig_e, use_container_width=True)
            st.download_button(
                f"下载 ①整项删减 CSV ({base_y4}_vs_{compare_y4})",
                data=df_elim.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"降本增效_整项删减_{base_y4}_vs_{compare_y4}.csv",
                mime="text/csv",
                key="dl_elim_tab4",
            )

        st.divider()

        st.subheader("② 分类金额压减（两年都有该分类，对比年全年花得更少）")
        if df_red.empty:
            st.info("无金额压减记录（对比年该项均不低于基准年）。")
        else:
            st.metric("条数", len(df_red))
            topn_r = st.slider("金额压减表 TopN", 20, 500, 100, key="tab4_top_red")
            show_r = df_red.head(topn_r).copy()
            cols_r = ["分类"] + [col_b, col_c, "类型", "节约参考"]
            show_r = show_r[[c for c in cols_r if c in show_r.columns]].reset_index(drop=True)
            show_r.index = show_r.index + 1
            st.dataframe(show_r, use_container_width=True)
            by_cat_r = df_red.groupby("分类", as_index=False)["节约参考"].sum().sort_values("节约参考", ascending=False)
            br_plot = by_cat_r.head(15).copy()
            br_plot["节约_万元"] = _to_wan(br_plot["节约参考"])
            fig_r = px.bar(br_plot, x="节约_万元", y="分类", orientation="h", title="金额压减 · 按分类")
            fig_r.update_traces(
                hovertemplate="分类=%{y}<br>压减节约=%{customdata:,.2f} 元<br>（%{x:,.2f} 万元）",
                customdata=br_plot["节约参考"],
            )
            fig_r.update_xaxes(title_text="压减节约（万元）", tickformat=",.2f")
            st.plotly_chart(fig_r, use_container_width=True)
            st.download_button(
                f"下载 ②金额压减 CSV ({base_y4}_vs_{compare_y4})",
                data=df_red.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"降本增效_金额压减_{base_y4}_vs_{compare_y4}.csv",
                mime="text/csv",
                key="dl_partial_tab4",
            )

        st.divider()
        st.subheader("③ 合并导出（整项删减 + 金额压减）")
        if df_elim.empty and df_red.empty:
            st.info("无可合并数据。")
        else:
            parts = [p for p in (df_elim, df_red) if not p.empty]
            merged = pd.concat(parts, ignore_index=True).sort_values("节约参考", ascending=False)
            st.download_button(
                "下载合并 CSV",
                data=merged.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"降本增效_合并_{base_y4}_vs_{compare_y4}.csv",
                mime="text/csv",
                key="dl_merged_tab4",
            )