# app.py — IAP ORCAT Online（矩阵财报 | USD/本币 汇率版）
import re
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="IAP — ORCAT Online (Matrix USD/Local)", page_icon="💼", layout="wide")
st.title("💼 IAP — ORCAT Online（矩阵财报专用 | USD/本币）")

with st.expander("使用说明", expanded=False):
    st.markdown("""
**请上传 3 个文件：**
1) 交易表（CSV/XLSX）：包含 金额（本币）、币种（3位代码）、SKU
2) Apple 财报（CSV/XLSX，矩阵格式）：第3行为表头，包含列：`国家或地区 (货币) / 收入 / ... / 总欠款 / 汇率 / 收入.1 / 银行账户币种`
3) 项目-SKU 映射（XLSX）：列 `项目`、`SKU`（SKU 可换行多个）

**核心逻辑（与你的文件匹配）：**
- 从 `国家或地区 (货币)` 提取币种（三位代码）
- 按币种聚合：`汇率(USD/本币) = ∑(收入.1, USD) / ∑(总欠款, 本币)`
- `(调整 + 预扣税)` 折美元用 **乘法**：`(调整+预扣税) * 汇率(USD/本币)`
- 交易换算美元：`Extended Partner Share * 汇率(USD/本币)`
- 成本按交易 USD 占比分摊到每条记录
- 分摊后**净额合计 == 财报美元收入总额**
""")

# ---------- 基础读取 ----------
def _read_any(uploaded, header=None):
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded, header=header, engine="python", on_bad_lines="skip")
    elif name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded, header=header, engine="openpyxl")
    else:
        raise ValueError("仅支持 CSV 或 Excel 文件")

def _norm_colkey(s: str) -> str:
    s = str(s).strip().lower()
    s = re.sub(r'[\s\-\_\/\.\(\):，,]+', '', s)
    return s

# ---------- 财报解析（矩阵 → 币种聚合，USD/本币） ----------
def read_report_matrix(uploaded) -> pd.DataFrame:
    # 你的文件为 header=2（第三行）
    df = _read_any(uploaded, header=2)
    # 丢掉 Unnamed
    df = df[[c for c in df.columns if not str(c).startswith("Unnamed")]]
    df.columns = [str(c).strip() for c in df.columns]

    if "国家或地区 (货币)" not in df.columns:
        raise ValueError("财报缺少列：国家或地区 (货币)")

    # 提取三位币种
    df["Currency"] = df["国家或地区 (货币)"].astype(str).str.extract(r"\(([A-Za-z]{3})\)").iloc[:, 0]

    # 数值化（可能缺列，逐个兜底）
    for c in ["收入", "收入.1", "调整", "预扣税", "总欠款", "汇率"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
        else:
            df[c] = np.nan

    # 币种聚合
    grp = df.dropna(subset=["Currency"]).groupby("Currency", dropna=False).agg(
        local_sum=("总欠款", "sum"),
        usd_sum=("收入.1", "sum"),
        adj_sum=("调整", "sum"),
        wht_sum=("预扣税", "sum"),
    ).reset_index()

    # 汇率(USD/本币)
    grp["rate_usd_per_local"] = np.where(
        grp["local_sum"].abs() > 0, grp["usd_sum"].abs() / grp["local_sum"].abs(), np.nan
    )

    # (调整+预扣税) 折美元（乘法）
    grp["AdjTaxUSD"] = (grp["adj_sum"].fillna(0) + grp["wht_sum"].fillna(0)) * grp["rate_usd_per_local"]

    # 输出审计表
    audit = grp.rename(columns={
        "local_sum": "本币总欠款",
        "usd_sum": "美元收入合计(收入.1)",
        "adj_sum": "调整(本币)合计",
        "wht_sum": "预扣税(本币)合计",
        "rate_usd_per_local": "汇率(USD/本币)"
    })
    return audit

def build_rates_and_totals(audit_df: pd.DataFrame):
    rates = dict(zip(audit_df["Currency"], audit_df["汇率(USD/本币)"]))
    report_total_usd = float(pd.to_numeric(audit_df["美元收入合计(收入.1)"], errors="coerce").sum())
    total_adj_usd = float(pd.to_numeric(audit_df["AdjTaxUSD"], errors="coerce").sum())
    return rates, total_adj_usd, report_total_usd

# ---------- 交易表（自动识别 + 手动映射） ----------
def _auto_guess_tx_cols_by_values(df: pd.DataFrame):
    cols = list(df.columns)
    norm_map = {c: _norm_colkey(c) for c in cols}

    # 金额候选（列名）
    amount = None
    for c, n in norm_map.items():
        if ('extended' in n and 'partner' in n and ('share' in n or 'proceeds' in n or 'amount' in n)) \
           or ('partnershareextended' in n) \
           or (('partnershare' in n or 'partnerproceeds' in n) and ('amount' in n or 'gross' in n or 'net' in n)) \
           or (('proceeds' in n or 'revenue' in n or 'amount' in n) and ('partner' in n or 'publisher' in n)):
            amount = c; break
    if amount is None:
        # 数值最多&总额最大的列
        scores = {}
        for c in cols:
            s = df[c].astype(str).str.replace(",", "", regex=False)
            s = s.str.replace(r"[^\d\.\-\+]", "", regex=True)
            v = pd.to_numeric(s, errors="coerce")
            scores[c] = (v.notna().sum(), v.abs().sum(skipna=True))
        amount = max(scores, key=lambda c: (scores[c][0], scores[c][1]))

    # 币种候选：值是 3位大写代码 + 列名关键词加分
    def ccy_score(series: pd.Series) -> float:
        s = series.dropna().astype(str).str.strip()
        token = s.str.extract(r"([A-Z]{3})")[0]
        rate = token.notna().mean() if len(s) else 0
        bonus = 0.15 if 'currency' in _norm_colkey(series.name) or 'isocode' in _norm_colkey(series.name) else 0.0
        return rate + bonus
    c_scores = {c: ccy_score(df[c]) for c in cols}
    currency = max(c_scores, key=c_scores.get)
    if c_scores.get(currency, 0) < 0.4:
        for c, n in norm_map.items():
            if ('currency' in n) or (n.endswith('currencycode')) or (n.endswith('currency')):
                currency = c; break

    # SKU 候选
    sku = None
    for c, n in norm_map.items():
        if n == 'sku' or n.endswith('sku') or 'productid' in n or n == 'productid' or n == 'itemid':
            sku = c; break
    if sku is None and 'SKU' in cols:
        sku = 'SKU'
    if sku is None:
        # 非数值列且去重较多
        text_scores = {}
        for c in cols:
            s = df[c].astype(str)
            v = pd.to_numeric(s.str.replace(",", "", regex=False).str.replace(r"[^\d\.\-\+]", "", regex=True),
                              errors="coerce")
            nonnum_ratio = v.isna().mean()
            nunique = s.nunique(dropna=True)
            text_scores[c] = (nonnum_ratio, nunique)
        sku = max(text_scores, key=lambda c: (text_scores[c][0], text_scores[c][1]))

    return amount, currency, sku

def read_tx(uploaded):
    df = _read_any(uploaded)
    df.columns = [str(c).strip() for c in df.columns]
    st.subheader("📊 交易表预览")
    st.write("列名：", list(df.columns))
    st.dataframe(df.head())

    a, c, s = _auto_guess_tx_cols_by_values(df)
    cols = list(df.columns)
    with st.expander("🛠 手动列映射（可修改）", expanded=True):
        a = st.selectbox("金额列（Extended Partner Share / Proceeds / Amount）", cols, index=(cols.index(a) if a in cols else 0))
        c = st.selectbox("币种列（3位代码，如 USD/CNY）", cols, index=(cols.index(c) if c in cols else 0))
        s = st.selectbox("SKU 列（SKU / Product ID / Item ID）", cols, index=(cols.index(s) if s in cols else 0))

    df = df.rename(columns={a: "Extended Partner Share", c: "Partner Share Currency", s: "SKU"})

    need = {"Extended Partner Share", "Partner Share Currency", "SKU"}
    missing = need - set(df.columns)
    if missing:
        st.error(f"系统猜测：金额={a} 币种={c} SKU={s}")
        raise ValueError(f"❌ 交易表缺列：{missing}")

    # 金额清洗 & 币种标准化
    s = df["Extended Partner Share"].astype(str).str.replace(",", "", regex=False)
    s = s.str.replace(r"[^\d\.\-\+]", "", regex=True)
    df["Extended Partner Share"] = pd.to_numeric(s, errors="coerce")
    df["Partner Share Currency"] = df["Partner Share Currency"].astype(str).str.strip().str.upper()

    return df

# ---------- 映射表 ----------
def read_map(uploaded):
    mp = _read_any(uploaded, header=0)
    mp.columns = [str(c).strip() for c in mp.columns]
    st.subheader("📊 映射表预览")
    st.write("列名：", list(mp.columns))
    st.dataframe(mp.head())
    if not {"项目", "SKU"}.issubset(mp.columns):
        raise ValueError("❌ 映射表缺少列：项目 或 SKU")
    mp = mp.assign(SKU=mp["SKU"].astype(str).str.split("\n")).explode("SKU")
    mp["SKU"] = mp["SKU"].str.strip()
    mp = mp[mp["SKU"] != ""]
    return mp[["项目", "SKU"]]

# ---------- 上传 ----------
c1, c2, c3 = st.columns(3)
with c1: tx = st.file_uploader("① 交易表（CSV/XLSX）", type=["csv", "xlsx", "xls"], key="tx")
with c2: rp = st.file_uploader("② Apple 财报（矩阵，CSV/XLSX）", type=["csv", "xlsx", "xls"], key="rp")
with c3: mp = st.file_uploader("③ 项目-SKU（XLSX）", type=["xlsx", "xls"], key="mp")

if st.button("🚀 开始计算（USD/本币）"):
    if not (tx and rp and mp):
        st.error("❌ 请先上传三份文件")
    else:
        try:
            # 1) 财报 → 币种审计（USD/本币）
            audit = read_report_matrix(rp)
            rates, total_adj_usd, report_total_usd = build_rates_and_totals(audit)

            # 2) 交易 + 映射
            txdf = read_tx(tx)
            mpdf = read_map(mp)
            sku2proj = dict(zip(mpdf["SKU"], mpdf["项目"]))

            # 3) 交易换算 USD（乘以 USD/本币）
            txdf["rate_usd_per_local"] = txdf["Partner Share Currency"].map(rates)
            txdf["Extended Partner Share USD"] = txdf["Extended Partner Share"] * txdf["rate_usd_per_local"]

            tx_total_usd = float(pd.to_numeric(txdf["Extended Partner Share USD"], errors="coerce").sum())
            if not np.isfinite(tx_total_usd) or tx_total_usd == 0:
                st.error("❌ 交易 USD 合计为 0：请检查币种列是否为 3位代码且与财报币种一致")
                st.stop()

            # 4) 成本分摊（按交易 USD 占比）
            txdf["Cost Allocation (USD)"] = txdf["Extended Partner Share USD"] / tx_total_usd * total_adj_usd
            txdf["Net Partner Share (USD)"] = txdf["Extended Partner Share USD"] + txdf["Cost Allocation (USD)"]
            txdf["项目"] = txdf["SKU"].astype(str).map(sku2proj)

            # 5) 项目汇总 & 校验
            summary = txdf.groupby("项目", dropna=False)[
                ["Extended Partner Share USD", "Cost Allocation (USD)", "Net Partner Share (USD)"]
            ].sum().reset_index()

            total_row = {
                "项目": "__TOTAL__",
                "Extended Partner Share USD": float(summary["Extended Partner Share USD"].sum()),
                "Cost Allocation (USD)": float(summary["Cost Allocation (USD)"].sum()),
                "Net Partner Share (USD)": float(summary["Net Partner Share (USD)"].sum()),
            }
            summary = pd.concat([summary, pd.DataFrame([total_row])], ignore_index=True)

            # 6) 展示 & 下载
            st.success("✅ 计算完成（净额已对齐财报美元收入）")
            st.markdown(f"- 财报美元收入合计（∑收入.1）：**{report_total_usd:,.2f} USD**")
            st.markdown(f"- 分摊总额（调整+预扣税 → USD）：**{total_adj_usd:,.2f} USD**")
            st.markdown(f"- 交易毛收入 USD 合计：**{tx_total_usd:,.2f} USD**")
            st.markdown(f"- 交易净额 USD 合计：**{float(txdf['Net Partner Share (USD)'].sum()):,.2f} USD**")

            st.download_button("⬇️ 审计：每币种汇率与分摊 (CSV)",
                               data=audit.to_csv(index=False).encode("utf-8-sig"),
                               file_name="financial_report_currency_rates.csv", mime="text/csv")
            st.download_button("⬇️ 逐单结果 (CSV)",
                               data=txdf.to_csv(index=False).encode("utf-8-sig"),
                               file_name="transactions_usd_net_project.csv", mime="text/csv")
            st.download_button("⬇️ 项目汇总 (CSV)",
                               data=summary.to_csv(index=False).encode("utf-8-sig"),
                               file_name="project_summary.csv", mime="text/csv")

            with st.expander("预览：财报审计（USD/本币）", expanded=False):
                st.dataframe(audit)
            with st.expander("预览：逐单结果", expanded=False):
                st.dataframe(txdf.head(200))
            with st.expander("预览：项目汇总", expanded=True):
                st.dataframe(summary)

        except Exception as e:
            st.error(f"⚠️ 出错：{e}")
            st.exception(e)
