# app.py — IAP ORCAT Online（严格模式｜财报表头=第3行｜使用“汇率”列｜稳健币种解析&有效行筛选）

import re
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="IAP — ORCAT Online (Strict+Robust)", page_icon="💼", layout="wide")
st.title("💼 IAP — ORCAT Online（严格｜财报表头=第3行｜使用财报汇率）")

with st.expander("使用说明", expanded=False):
    st.markdown("""
**① 财报（CSV/XLSX，表头=第3行）**  
- `国家或地区 (货币)`：例如 `阿拉伯联合酋长国 (AED)` / `阿联酋（AED）` / `阿联酋-AED` / `阿联酋 AED`  
- `总欠款`（本币）、`收入.1`（USD）、`调整`（本币，可空）、`预扣税`（本币，可空）、`汇率`（USD/本币，**直接使用**）

**② 交易表（CSV/XLSX）**：`Extended Partner Share`、`Partner Share Currency`、`SKU`  
**③ 映射表（XLSX）**：`项目`、`SKU`（SKU 可换行多值）

**规则**  
- 仅在**有效数据行**上解析币种与做统计（排除标题/小计/空行）  
- 同一币种多行时，`汇率`取**中位数**；(调整+预扣税)×汇率→USD 后按交易 USD 占比分摊  
- 交易本币 ×汇率 → USD；`∑净额 ≈ ∑财报(收入.1)`（容差 0.5 USD）
""")

# ---------- 通用工具 ----------
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

# ---------- 财报读取 ----------
REQ_REPORT = ["国家或地区 (货币)", "总欠款", "收入.1", "汇率"]
OPT_REPORT = ["调整", "预扣税"]

# 谨慎回退（绝不使用“银行账户币种”这种常为单值的列）
_CCY_FALLBACK_COLS = ["币种", "货币", "Currency"]

def _extract_currency_on_valid(df: pd.DataFrame) -> pd.Series:
    """只在有效数据行上解析币种，并对少量缺失行进行忽略处理"""
    s = df["国家或地区 (货币)"].astype(str)

    # 1) 括号中的 3 位代码：全/半角都支持
    pat_paren = re.compile(r"[（(]\s*([A-Za-z]{3})\s*[）)]")
    c_paren = s.str.extract(pat_paren, expand=False)

    # 2) 末尾连接符/空格/斜杠后的 3 位代码（如 '- AED'、'/AED'、' AED'）
    c_tail = s.where(c_paren.notna(), s).str.extract(r"(?:-|/|\s)([A-Za-z]{3})\s*$", expand=False)

    # 3) 全文任意 3 位大写代码（最后兜底）
    c_any = s.where(c_paren.notna() | c_tail.notna(), s).str.extract(r"\b([A-Z]{3})\b", expand=False)

    cur = c_paren.fillna(c_tail).fillna(c_any).astype(str).str.upper().replace("NAN", np.nan)

    # 只在“有效行”上评估缺失率
    valid_mask = (
        df["总欠款"].notna() |
        df["收入.1"].notna() |
        df["汇率"].notna()
    )
    valid_cnt = valid_mask.sum()
    if valid_cnt == 0:
        raise ValueError("财报没有有效数据行（请检查表头是否在第3行、数值列是否为空）")

    # 谨慎回退：仅当有效行缺失率仍高时，且回退列不是单一常量才使用
    miss_ratio = cur[valid_mask].isna().mean()
    if miss_ratio > 0.2:
        for col in _CCY_FALLBACK_COLS:
            if col in df.columns:
                alt_raw = df[col].astype(str)
                alt = alt_raw.str.extract(r"\b([A-Za-z]{3})\b", expand=False).str.upper()
                uniq = alt[valid_mask].dropna().unique()
                if len(uniq) <= 1:
                    continue
                cur = cur.fillna(alt)

    # 仍有极少数有效行未能识别：直接忽略这些行并提示
    still_nan = cur[valid_mask].isna()
    miss_rows = int(still_nan.sum())
    if miss_rows > 0:
        bad_samples = s[valid_mask & still_nan].head(6).tolist()
        st.warning(f"以下有效行无法提取币种，已自动忽略 {miss_rows} 行（示例：{bad_samples}）")
        # 对这些行置空即可，后续会在 groupby 前丢弃 Currency 为 NaN 的行
    return cur

def read_report_final(uploaded):
    # 表头=第3行
    df = _read_any(uploaded, header=2)
    df = df[[c for c in df.columns if not str(c).startswith("Unnamed")]]
    df.columns = [str(c).strip() for c in df.columns]

    # 必需列检查
    missing = [c for c in REQ_REPORT if c not in df.columns]
    if missing:
        raise ValueError(f"财报缺少必需列：{missing}")

    # 数值化
    for c in REQ_REPORT + OPT_REPORT:
        if c in df.columns:
            df[c] = _num(df[c])
        else:
            df[c] = np.nan

    # 币种解析（仅看有效行、对少数无法识别的行直接忽略）
    df["Currency"] = _extract_currency_on_valid(df)

    # 仅保留 Currency 非空的有效统计行
    stat = df.loc[df["Currency"].notna() & (
        df["总欠款"].notna() | df["收入.1"].notna() | df["汇率"].notna()
    )].copy()

    if stat.empty:
        raise ValueError("财报有效统计行为空（可能全部为标题/合计或币种列完全缺失）")

    grp = stat.groupby("Currency", dropna=False).agg(
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

    # 若只识别出 1 个币种，也允许继续（但会在交易覆盖校验时报错更明确）
    rates = dict(zip(audit["Currency"], audit["汇率(USD/本币)"]))
    report_total_usd = float(audit["美元收入合计(收入.1)"].sum())
    total_adj_usd = float(audit["AdjTaxUSD"].sum())

    inconsistent = audit.loc[
        audit["rate_min"].round(8) != audit["rate_max"].round(8),
        ["Currency","rate_min","rate_max","rows"]
    ]
    if len(inconsistent):
        st.warning("以下币种的财报`汇率`存在差异，**已使用中位数**：")
        st.dataframe(inconsistent)

    st.info(f"财报识别到的币种：{sorted(audit['Currency'].dropna().unique().tolist())}")

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

        # 交易币种覆盖校验
        tx_ccy = set(tx["Partner Share Currency"].dropna().unique())
        missing_ccy = sorted(tx_ccy - set(rates.keys()))
        if missing_ccy:
            raise ValueError(f"交易表出现财报未覆盖的币种：{missing_ccy}")

        # 3) 交易计算
        tx["rate_usd_per_local"] = tx["Partner Share Currency"].map(rates).astype(float)
        tx["Extended Partner Share USD"] = tx["Extended Partner Share"] * tx["rate_usd_per_local"]

        tx_total_usd = float(tx["Extended Partner Share USD"].sum())
        if not np.isfinite(tx_total_usd) or tx_total_usd == 0:
            raise ValueError("交易 USD 合计为 0，请检查金额列或单位。")

        tx["Cost Allocation (USD)"] = tx["Extended Partner Share USD"] / tx_total_usd * total_adj_usd
        tx["Net Partner Share (USD)"] = tx["Extended Partner Share USD"] + tx["Cost Allocation (USD)"]

        # 4) 映射与汇总
        if not mp_file:
            raise ValueError("未上传项目–SKU 映射")
        mp = read_map_final(mp_file)
        sku2proj = dict(zip(mp["SKU"], mp["项目"]))
        tx["项目"] = tx["SKU"].astype(str).map(sku2proj)

        summary = tx.groupby("项目", dropna=False)[
            ["Extended Partner Share USD","Cost Allocation (USD)","Net Partner Share (USD)"]
        ].sum().reset_index()

        net_total = float(tx["Net Partner Share (USD)"].sum())
        diff = net_total - report_total_usd
        if strict_check and abs(diff) > 0.5:
            raise ValueError(f"对账失败：交易净额 {net_total:,.2f} USD 与财报 {report_total_usd:,.2f} USD 差异 {diff:,.2f} USD")

        # 5) 输出
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
