# app.py — IAP ORCAT Online（严格模式｜矩阵财报＝第1行表头｜USD/本币）
import re
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="IAP — ORCAT Online (Strict)", page_icon="💼", layout="wide")
st.title("💼 IAP — ORCAT Online（严格模式｜矩阵财报｜USD/本币）")

with st.expander("使用说明", expanded=False):
    st.markdown("""
**请上传 3 个文件（列名必须与下方完全一致）：**

**① 财报（CSV/XLSX，表头=第1行，矩阵）**  
- `国家或地区 (货币)`（例如 `阿拉伯联合酋长国 (AED)`，括号内**必须是3位代码**）  
- `销量`（可有可无）  
- `收入`（本币）  
- `税前小计`（本币）  
- `进项税`（本币）  
- `调整`（本币）  
- `预扣税`（本币）  
- `总欠款`（本币） **[必需]**  
- `汇率`（可空；实际计算用 ∑USD/∑本币）  
- `收入.1`（美元收入 USD） **[必需]**  
- `银行账户币种`（可空）

**② 交易表（CSV/XLSX）**  
- `Extended Partner Share`（本币金额） **[必需]**  
- `Partner Share Currency`（3位币种代码） **[必需]**  
- `SKU` **[必需]**

**③ 项目–SKU 映射（XLSX）**  
- `项目` **[必需]**  
- `SKU`（支持一格多值，换行分隔） **[必需]**

**计算规则（固定）：**  
- 每币种汇率 = **USD/本币** = `∑(收入.1) / ∑(总欠款)`  
- `(调整+预扣税)` 折美元 = **乘法** `(调整+预扣税) * 汇率(USD/本币)`  
- 交易USD = `Extended Partner Share * 汇率(USD/本币)`（USD 自身=1）  
- 成本按交易USD占比分摊到每行；对账：`∑净额 ≈ ∑(收入.1)`  
""")

# ---------------- 工具 ----------------
def _read_any(uploaded, header=0):
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded, header=header, engine="python", on_bad_lines="skip")
    elif name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded, header=header, engine="openpyxl")
    else:
        raise ValueError("仅支持 CSV 或 Excel 文件")

def _num(s: pd.Series) -> pd.Series:
    t = s.astype(str).str.replace(",", "", regex=False).str.replace(r"[^\d\.\-\+]", "", regex=True)
    return pd.to_numeric(t, errors="coerce")

# ---------------- 财报（严格列名） ----------------
REQ_REPORT = ["国家或地区 (货币)", "总欠款", "收入.1"]  # 其余列可空
OPT_REPORT = ["调整", "预扣税", "收入", "税前小计", "进项税", "汇率", "银行账户币种", "销量"]

def read_report_strict(uploaded) -> pd.DataFrame:
    df = _read_any(uploaded, header=0)
    # 去掉 Unnamed
    df = df[[c for c in df.columns if not str(c).startswith("Unnamed")]]
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REQ_REPORT if c not in df.columns]
    if missing:
        raise ValueError(f"财报缺少必需列：{missing}。请确认表头在第1行且列名完全一致。")

    # 数值化（不存在则补NaN）
    for c in REQ_REPORT + OPT_REPORT:
        if c in df.columns:
            if c in ["总欠款","收入.1","调整","预扣税","收入","税前小计","进项税","汇率"]:
                df[c] = _num(df[c])
        else:
            df[c] = np.nan

    # 提取币种：括号中的3字母
    if "国家或地区 (货币)" not in df.columns:
        raise ValueError("财报缺少列：国家或地区 (货币)")
    df["Currency"] = df["国家或地区 (货币)"].astype(str).str.extract(r"\(([A-Za-z]{3})\)").iloc[:,0]
    if df["Currency"].isna().all():
        raise ValueError("无法从“国家或地区 (货币)”提取币种（应为如 `中国 (CNY)` 的格式）。")

    # 按币种聚合
    grp = df.dropna(subset=["Currency"]).groupby("Currency", dropna=False).agg(
        local_sum=("总欠款","sum"),
        usd_sum=("收入.1","sum"),
        adj_sum=("调整","sum"),
        wht_sum=("预扣税","sum")
    ).reset_index()

    # 汇率（USD/本币）
    grp["rate_usd_per_local"] = np.where(grp["local_sum"].abs()>0,
                                         grp["usd_sum"].abs()/grp["local_sum"].abs(), np.nan)
    # (调整+预扣税) 折美元（乘法）
    grp["AdjTaxUSD"] = (grp["adj_sum"].fillna(0) + grp["wht_sum"].fillna(0)) * grp["rate_usd_per_local"]

    audit = grp.rename(columns={
        "local_sum":"本币总欠款",
        "usd_sum":"美元收入合计(收入.1)",
        "adj_sum":"调整(本币)合计",
        "wht_sum":"预扣税(本币)合计",
        "rate_usd_per_local":"汇率(USD/本币)"
    })

    rates = dict(zip(audit["Currency"], audit["汇率(USD/本币)"]))
    report_total_usd = float(pd.to_numeric(audit["美元收入合计(收入.1)"], errors="coerce").sum())
    total_adj_usd = float(pd.to_numeric(audit["AdjTaxUSD"], errors="coerce").sum())

    return audit, rates, total_adj_usd, report_total_usd

# ---------------- 交易表（严格列名） ----------------
REQ_TX = ["Extended Partner Share", "Partner Share Currency", "SKU"]

def read_tx_strict(uploaded, amount_unit: str) -> pd.DataFrame:
    df = _read_any(uploaded, header=0)
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQ_TX if c not in df.columns]
    if missing:
        raise ValueError(f"交易表缺少必需列：{missing}")

    df["Extended Partner Share"] = _num(df["Extended Partner Share"])
    if amount_unit == "分(÷100)":
        df["Extended Partner Share"] = df["Extended Partner Share"] / 100.0
    elif amount_unit == "厘(÷1000)":
        df["Extended Partner Share"] = df["Extended Partner Share"] / 1000.0

    df["Partner Share Currency"] = df["Partner Share Currency"].astype(str).str.strip().str.upper()
    return df[REQ_TX].copy()

# ---------------- 映射表（严格列名） ----------------
REQ_MAP = ["项目","SKU"]

def read_map_strict(uploaded) -> pd.DataFrame:
    mp = _read_any(uploaded, header=0)
    mp.columns = [str(c).strip() for c in mp.columns]
    missing = [c for c in REQ_MAP if c not in mp.columns]
    if missing:
        raise ValueError(f"映射表缺少必需列：{missing}")
    # SKU 支持一格多值（换行）
    mp = mp.assign(SKU=mp["SKU"].astype(str).str.split("\n")).explode("SKU")
    mp["SKU"] = mp["SKU"].str.strip()
    mp = mp[mp["SKU"]!=""]
    return mp[["项目","SKU"]].copy()

# ---------------- UI ----------------
c1, c2, c3 = st.columns(3)
with c1: tx_file = st.file_uploader("① 交易表（CSV/XLSX）", type=["csv","xlsx","xls"])
with c2: rp_file = st.file_uploader("② 财报（CSV/XLSX｜表头=第1行）", type=["csv","xlsx","xls"])
with c3: mp_file = st.file_uploader("③ 项目–SKU（XLSX）", type=["xlsx","xls"])

amount_unit = st.radio("交易金额单位", ["元(不用换)", "分(÷100)", "厘(÷1000)"], index=0, horizontal=True)
strict_check = st.checkbox("严格对账：|∑净额 − ∑财报USD| ≤ 0.5", value=True)

if st.button("🚀 开始计算（严格模式）"):
    try:
        # 1) 财报
        if not rp_file: raise ValueError("未上传财报")
        audit, rates, total_adj_usd, report_total_usd = read_report_strict(rp_file)

        # 2) 交易表
        if not tx_file: raise ValueError("未上传交易表")
        tx = read_tx_strict(tx_file, amount_unit)

        # 校验：币种必须全部在财报中
        tx_ccy = set(tx["Partner Share Currency"].dropna().unique())
        missing_ccy = sorted(tx_ccy - set(k for k,v in rates.items() if np.isfinite(v)))
        if missing_ccy:
            raise ValueError(f"交易表出现财报未覆盖的币种：{missing_ccy}（请修正币种或财报）")

        # 3) 交易→USD，分摊
        tx["rate_usd_per_local"] = tx["Partner Share Currency"].map(rates).astype(float)
        tx["Extended Partner Share USD"] = tx["Extended Partner Share"] * tx["rate_usd_per_local"]

        tx_total_usd = float(pd.to_numeric(tx["Extended Partner Share USD"], errors="coerce").sum())
        if not np.isfinite(tx_total_usd) or tx_total_usd == 0:
            raise ValueError("交易 USD 合计为 0：请检查金额列与金额单位。")

        tx["Cost Allocation (USD)"] = tx["Extended Partner Share USD"] / tx_total_usd * total_adj_usd
        tx["Net Partner Share (USD)"] = tx["Extended Partner Share USD"] + tx["Cost Allocation (USD)"]

        # 4) 项目映射与汇总
        if not mp_file: raise ValueError("未上传项目–SKU 映射")
        mp = read_map_strict(mp_file)
        sku2proj = dict(zip(mp["SKU"], mp["项目"]))
        tx["项目"] = tx["SKU"].astype(str).map(sku2proj)

        summary = tx.groupby("项目", dropna=False)[
            ["Extended Partner Share USD","Cost Allocation (USD)","Net Partner Share (USD)"]
        ].sum().reset_index()

        net_total = float(pd.to_numeric(tx["Net Partner Share (USD)"], errors="coerce").sum())
        diff = net_total - report_total_usd
        if strict_check and (not np.isfinite(diff) or abs(diff) > 0.5):
            raise ValueError(f"对账失败：交易净额 {net_total:,.2f} USD 与财报 {report_total_usd:,.2f} USD 差异 {diff:,.2f}。")

        # 5) 展示与下载
        st.success("✅ 计算完成（严格模式）")
        st.markdown(f"- 财报美元收入合计（∑收入.1）：**{report_total_usd:,.2f} USD**")
        st.markdown(f"- 分摊总额（调整+预扣税→USD）：**{total_adj_usd:,.2f} USD**")
        st.markdown(f"- 交易毛收入 USD 合计：**{tx_total_usd:,.2f} USD**")
        st.markdown(f"- 交易净额 USD 合计：**{net_total:,.2f} USD**（差异：**{diff:,.2f} USD**）")

        st.download_button("⬇️ 审计：每币种汇率与分摊 (CSV)",
            data=audit.to_csv(index=False).encode("utf-8-sig"),
            file_name="financial_report_currency_rates.csv", mime="text/csv")
        st.download_button("⬇️ 逐单结果 (CSV)",
            data=tx.to_csv(index=False).encode("utf-8-sig"),
            file_name="transactions_usd_net_project.csv", mime="text/csv")
        st.download_button("⬇️ 项目汇总 (CSV)",
            data=summary.to_csv(index=False).encode("utf-8-sig"),
            file_name="project_summary.csv", mime="text/csv")

        with st.expander("预览：财报审计（USD/本币）", expanded=False):
            st.dataframe(audit)
        with st.expander("预览：逐单结果", expanded=False):
            st.dataframe(tx.head(200))
        with st.expander("预览：项目汇总", expanded=True):
            st.dataframe(summary)

    except Exception as e:
        st.error(f"⚠️ 出错：{e}")
        st.exception(e)
