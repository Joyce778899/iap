# app.py — IAP ORCAT Online（严格模式｜财报表头=第3行｜使用“汇率”列）

import re
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="IAP — ORCAT Online (Final)", page_icon="💼", layout="wide")
st.title("💼 IAP — ORCAT Online（严格｜财报表头=第3行｜使用财报汇率）")

with st.expander("使用说明", expanded=False):
    st.markdown("""
**① 财报（CSV/XLSX，表头=第3行）**  
- `国家或地区 (货币)`：如 `阿拉伯联合酋长国 (AED)`（括号内为 3 位代码）  
- `总欠款`（本币）  
- `收入.1`（美元收入＝总欠款折 USD）  
- `调整`（本币，可空）  
- `预扣税`（本币，可空）  
- `汇率`（**USD/本币**；直接使用该列）

**② 交易表（CSV/XLSX）**  
- `Extended Partner Share`（本币金额）  
- `Partner Share Currency`（3位币种代码）  
- `SKU`

**③ 项目–SKU 映射（XLSX）**  
- `项目`，`SKU`（SKU 支持换行多个）

**规则**  
- 财报按币种取 `汇率` 中位数为 **USD/本币**  
- `(调整+预扣税)`（本币）×汇率 → USD 后分摊到交易  
- 交易本币 ×汇率 → USD；按占比摊成本  
- 对账：∑净额 ≈ ∑财报 USD（容差 0.5）
""")

# ---------- 工具 ----------
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

# ---------- 财报 ----------
REQ_REPORT = ["国家或地区 (货币)", "总欠款", "收入.1", "汇率"]
OPT_REPORT = ["调整", "预扣税"]

def read_report_final(uploaded):
    # header=2 → 第3行表头
    df = _read_any(uploaded, header=2)
    df = df[[c for c in df.columns if not str(c).startswith("Unnamed")]]
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REQ_REPORT if c not in df.columns]
    if missing:
        raise ValueError(f"财报缺少必需列：{missing}")

    for c in REQ_REPORT + OPT_REPORT:
        if c in df.columns:
            df[c] = _num(df[c])
        else:
            df[c] = np.nan

    # 拆分国家和币种
    m = df["国家或地区 (货币)"].astype(str).str.extract(r"^\s*(.+?)\s*\(([A-Za-z]{3})\)\s*$")
    df["Country"] = m[0]
    df["Currency"] = m[1]
    if df["Currency"].isna().any():
        bad = df.loc[df["Currency"].isna(), "国家或地区 (货币)"].head(5).tolist()
        raise ValueError(f"无法提取币种，示例问题行：{bad}")

    grp = df.groupby("Currency", dropna=False).agg(
        local_sum=("总欠款", "sum"),
        usd_sum=("收入.1", "sum"),
        adj_sum=("调整", "sum"),
        wht_sum=("预扣税", "sum"),
        rate_median=("汇率", "median"),
        rate_min=("汇率", "min"),
        rate_max=("汇率", "max"),
        rows=("汇率", "size"),
    ).reset_index()

    grp["汇率(USD/本币)"] = grp["rate_median"]
    grp["AdjTaxUSD"] = (grp["adj_sum"].fillna(0) + grp["wht_sum"].fillna(0)) * grp["汇率(USD/本币)"]

    audit = grp.rename(columns={
        "local_sum": "本币总欠款",
        "usd_sum": "美元收入合计(收入.1)",
        "adj_sum": "调整(本币)合计",
        "wht_sum": "预扣税(本币)合计",
    })

    rates = dict(zip(audit["Currency"], audit["汇率(USD/本币)"]))
    report_total_usd = audit["美元收入合计(收入.1)"].sum()
    total_adj_usd = audit["AdjTaxUSD"].sum()

    inconsistent = audit.loc[audit["rate_min"].round(8) != audit["rate_max"].round(8), ["Currency","rate_min","rate_max","rows"]]
    if len(inconsistent):
        st.warning("以下币种的财报`汇率`存在差异，已使用**中位数**：")
        st.dataframe(inconsistent)

    return audit, rates, total_adj_usd, report_total_usd

# ---------- 交易 ----------
REQ_TX = ["Extended Partner Share", "Partner Share Currency", "SKU"]

def read_tx_final(uploaded, amount_unit: str):
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

# ---------- 映射 ----------
REQ_MAP = ["项目","SKU"]

def read_map_final(uploaded):
    mp = _read_any(uploaded, header=0)
    mp.columns = [str(c).strip() for c in mp.columns]
    missing = [c for c in REQ_MAP if c not in mp.columns]
    if missing:
        raise ValueError(f"映射表缺少必需列：{missing}")
    mp = mp.assign(SKU=mp["SKU"].astype(str).str.split("\n")).explode("SKU")
    mp["SKU"] = mp["SKU"].str.strip()
    mp = mp[mp["SKU"]!=""]
    return mp[["项目","SKU"]].copy()

# ---------- UI ----------
c1, c2, c3 = st.columns(3)
with c1:
    tx_file = st.file_uploader("① 交易表（CSV/XLSX）", type=["csv","xlsx","xls"])
with c2:
    rp_file = st.file_uploader("② 财报（CSV/XLSX｜表头=第3行）", type=["csv","xlsx","xls"])
with c3:
    mp_file = st.file_uploader("③ 项目–SKU（XLSX）", type=["xlsx","xls"])

amount_unit = st.radio("交易金额单位", ["元(不用换)", "分(÷100)", "厘(÷1000)"], index=0, horizontal=True)
strict_check = st.checkbox("严格对账：|∑净额 − ∑财报USD| ≤ 0.5", value=True)

if st.button("🚀 开始计算"):
    try:
        # 1) 财报
        if not rp_file:
            raise ValueError("未上传财报")
        audit, rates, total_adj_usd, report_total_usd = read_report_final(rp_file)

        # 2) 交易
        if not tx_file:
            raise ValueError("未上传交易表")
        tx = read_tx_final(tx_file, amount_unit)

        tx_ccy = set(tx["Partner Share Currency"].dropna().unique())
        missing_ccy = sorted(tx_ccy - set(rates.keys()))
        if missing_ccy:
            raise ValueError(f"交易表出现财报未覆盖的币种：{missing_ccy}")

        tx["rate_usd_per_local"] = tx["Partner Share Currency"].map(rates).astype(float)
        tx["Extended Partner Share USD"] = tx["Extended Partner Share"] * tx["rate_usd_per_local"]

        tx_total_usd = tx["Extended Partner Share USD"].sum()
        if not np.isfinite(tx_total_usd) or tx_total_usd == 0:
            raise ValueError("交易 USD 合计为 0，请检查。")

        tx["Cost Allocation (USD)"] = tx["Extended Partner Share USD"] / tx_total_usd * total_adj_usd
        tx["Net Partner Share (USD)"] = tx["Extended Partner Share USD"] + tx["Cost Allocation (USD)"]

        # 3) 映射与汇总
        if not mp_file:
            raise ValueError("未上传项目–SKU 映射")
        mp = read_map_final(mp_file)
        sku2proj = dict(zip(mp["SKU"], mp["项目"]))
        tx["项目"] = tx["SKU"].astype(str).map(sku2proj)

        summary = tx.groupby("项目", dropna=False)[
            ["Extended Partner Share USD","Cost Allocation (USD)","Net Partner Share (USD)"]
        ].sum().reset_index()

        net_total = tx["Net Partner Share (USD)"].sum()
        diff = net_total - report_total_usd
        if strict_check and abs(diff) > 0.5:
            raise ValueError(f"对账失败：交易净额 {net_total:,.2f} USD 与财报 {report_total_usd:,.2f} USD 差异 {diff:,.2f} USD")

        # 4) 输出
        st.success("✅ 计算完成")
        st.markdown(f"- 财报美元收入合计（∑收入.1）：**{report_total_usd:,.2f} USD**")
        st.markdown(f"- 分摊总额（调整+预扣税 → USD）：**{total_adj_usd:,.2f} USD**")
        st.markdown(f"- 交易毛收入 USD 合计：**{tx_total_usd:,.2f} USD**")
        st.markdown(f"- 交易净额 USD 合计：**{net_total:,.2f} USD**（差异 {diff:,.2f} USD）")

        st.download_button("⬇️ 审计表 (CSV)",
            data=audit.to_csv(index=False).encode("utf-8-sig"),
            file_name="financial_report_audit.csv", mime="text/csv")
        st.download_button("⬇️ 逐单结果 (CSV)",
            data=tx.to_csv(index=False).encode("utf-8-sig"),
            file_name="transactions_usd.csv", mime="text/csv")
        st.download_button("⬇️ 项目汇总 (CSV)",
            data=summary.to_csv(index=False).encode("utf-8-sig"),
            file_name="project_summary.csv", mime="text/csv")

        with st.expander("预览：财报审计", expanded=False):
            st.dataframe(audit)

        with st.expander("预览：逐单结果", expanded=False):
            st.dataframe(tx.head(200))

        with st.expander("预览：项目汇总", expanded=True):
            st.dataframe(summary)

    except Exception as e:
        st.error(f"⚠️ 出错：{e}")
        st.exception(e)
