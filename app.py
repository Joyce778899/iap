# app.py
import io
import re
import json
import time
import math
import requests
import numpy as np
import pandas as pd
import streamlit as st
from io import StringIO
from datetime import datetime
from typing import Dict, Tuple

st.set_page_config(page_title="IAP 对账助手（财报汇率优先）", layout="wide")

# ----------------------------- Utils -----------------------------
def numify(s: pd.Series) -> pd.Series:
    t = s.astype(str).str.replace(",", "", regex=False).str.replace(r"[^\d\.\-\+]", "", regex=True)
    return pd.to_numeric(t, errors="coerce")

FALLBACK_RATES_USD_PER_1_LOCAL = {
    # 兜底（若在线获取失败再用）
    "AED": 0.272294,  # 1 AED ≈ 0.272294 USD（长期盯住美元）
    "QAR": 0.274725,  # 1 QAR ≈ 0.274725 USD（长期盯住美元）
    "ILS": 0.302000,  # 1 ILS ≈ 0.3020 USD
    "CLP": 1/955.383, # 1 CLP ≈ 0.0010467 USD
    "HUF": 1/332.981, # 1 HUF ≈ 0.0030032 USD
    "PKR": 1/282.639, # 1 PKR ≈ 0.0035381 USD
}

def fetch_usd_per_local(ccy_list):
    """
    从 exchangerate.host 获取 USD/本币（返回：1 本币 = ? USD）
    若接口不可用，落入兜底；未知币种返回 NaN。
    """
    rates = {}
    # exchangerate.host: GET https://api.exchangerate.host/latest?base=USD&symbols=XXX
    # 返回 base=USD，所以 1 本币 = USD/本币 = 1 / (base->ccy)
    for ccy in ccy_list:
        try:
            if ccy in ("USD", "NONE", "", None):
                rates[ccy] = (1.0 if ccy == "USD" else np.nan)
                continue
            resp = requests.get(
                "https://api.exchangerate.host/latest",
                params={"base": "USD", "symbols": ccy},
                timeout=8,
            )
            if resp.status_code == 200:
                data = resp.json()
                usd_to_ccy = data.get("rates", {}).get(ccy)
                if usd_to_ccy and usd_to_ccy > 0:
                    # base=USD: 1 USD = usd_to_ccy CCY
                    # 1 CCY = 1/usd_to_ccy USD
                    rates[ccy] = 1.0 / float(usd_to_ccy)
                else:
                    rates[ccy] = np.nan
            else:
                rates[ccy] = np.nan
        except Exception:
            rates[ccy] = np.nan

    # 兜底补全
    for ccy in ccy_list:
        if (ccy not in rates) or (pd.isna(rates[ccy]) or not np.isfinite(rates[ccy])):
            rates[ccy] = FALLBACK_RATES_USD_PER_1_LOCAL.get(ccy, np.nan)
    return rates

def detect_report_header_idx(csv_bytes: bytes) -> int:
    """
    自动判断财报 CSV 表头所在行：
    - 优先找含“国家或地区 (货币)”或“收入.1”的那一行（0-based）
    - 常见情况：0 或 2
    """
    text = csv_bytes.decode("utf-8", errors="ignore")
    lines = text.splitlines()
    for idx in range(min(6, len(lines))):
        if ("国家或地区" in lines[idx] and "货币" in lines[idx]) or ("收入.1" in lines[idx]):
            return idx
    # 默认第 2（即第三行）
    return 2

# ----------------------------- Readers -----------------------------
def read_report(csv_file) -> Tuple[pd.DataFrame, Dict[str,float], float, float]:
    """
    读取财报 CSV：
    - 自动定位表头行（第1或第3行）
    - 以“汇率”列为准；若某币种缺失，在线补齐
    - 汇总时仅使用同时具备【总欠款 & 收入.1】的行（避免口径串行）
    返回：audit_df, rates_dict(USD/本币), total_adj_usd, report_total_usd
    """
    raw = csv_file.read()
    header_idx = detect_report_header_idx(raw)
    df = pd.read_csv(io.BytesIO(raw), header=header_idx, engine="python", on_bad_lines="skip")
    # 去掉 Unnamed
    df = df[[c for c in df.columns if not str(c).startswith("Unnamed")]]
    df.columns = [str(c).strip() for c in df.columns]

    # 标准列
    need = ["国家或地区 (货币)", "总欠款", "收入.1"]
    for c in need:
        if c not in df.columns:
            raise ValueError(f"财报缺少列：{c}")

    # 数值列
    for c in ["总欠款","收入.1","调整","预扣税","汇率"]:
        if c in df.columns:
            df[c] = numify(df[c])
        else:
            df[c] = np.nan

    # 有效行 + 向下填充国家/货币
    mask = df[["总欠款","收入.1","调整","预扣税","汇率"]].notna().any(axis=1)
    col = "国家或地区 (货币)"
    df[col] = df[col].replace({"": np.nan, "nan": np.nan, "NaN": np.nan})
    df.loc[mask, col] = df.loc[mask, col].ffill()

    # 排除合计行
    bad_row = df[col].astype(str).str.contains(r"合计|总计|小计|Subtotal|Total", case=False, na=False)
    stat = df.loc[mask & ~bad_row].copy()

    # 提取币种：括号里的三字母，或末尾/任意位置的三字母
    s = stat[col].astype(str)
    pat_paren = re.compile(r"[（(]\s*([A-Za-z]{3})\s*[）)]")
    c_paren = s.str.extract(pat_paren, expand=False)
    c_tail  = s.where(c_paren.notna(), s).str.extract(r"(?:-|/|\s)([A-Za-z]{3})\s*$", expand=False)
    c_any   = s.where(c_paren.notna() | c_tail.notna(), s).str.extract(r"\b([A-Z]{3})\b", expand=False)
    stat["Currency"] = c_paren.fillna(c_tail).fillna(c_any).astype(str).str.upper().replace("NAN", np.nan)
    stat = stat.loc[stat["Currency"].notna()].copy()

    # 仅使用同时具备【总欠款 & 收入.1】的行参与本币/美元对比与反推
    stat["has_both"] = stat["总欠款"].notna() & stat["收入.1"].notna()
    stat["local_pair"] = stat["总欠款"].where(stat["has_both"])
    stat["usd_pair"]   = stat["收入.1"].where(stat["has_both"])

    # 财报“汇率”中位数（若存在）
    rate_report = None
    if "汇率" in stat.columns and stat["汇率"].notna().any():
        rate_report = stat.groupby("Currency")["汇率"].median().rename("rate_report").reset_index()

    grp = stat.groupby("Currency", dropna=False).agg(
        本币总欠款=("local_pair","sum"),
        美元收入合计_收入1=("usd_pair","sum"),
        调整_本币合计=("调整","sum"),
        预扣税_本币合计=("预扣税","sum"),
    ).reset_index().rename(columns={"美元收入合计_收入1": "美元收入合计(收入.1)"})

    if rate_report is not None:
        grp = grp.merge(rate_report, on="Currency", how="left")

    # 需要补齐汇率的币种
    need_ccy = sorted(set(grp.loc[grp["rate_report"].isna() if "rate_report" in grp.columns else grp.index*0 == 1, "Currency"]))
    if need_ccy:
        online = fetch_usd_per_local(need_ccy)
    else:
        online = {}

    grp["rate_used"] = grp["rate_report"] if "rate_report" in grp.columns else np.nan
    grp["rate_used"] = grp.apply(lambda r: online.get(r["Currency"], r["rate_used"]), axis=1)

    # 兜底再补
    grp["rate_used"] = grp.apply(
        lambda r: (FALLBACK_RATES_USD_PER_1_LOCAL.get(r["Currency"], r["rate_used"])
                   if (pd.isna(r["rate_used"]) or not np.isfinite(r["rate_used"])) else r["rate_used"]),
        axis=1
    )

    # (调整+预扣税) * 汇率(USD/本币)
    grp["AdjTaxUSD"] = (grp["调整_本币合计"].fillna(0) + grp["预扣税_本币合计"].fillna(0)) * grp["rate_used"]

    rates = dict(zip(grp["Currency"], grp["rate_used"]))
    report_total_usd = float(pd.to_numeric(grp["美元收入合计(收入.1)"], errors="coerce").sum())
    total_adj_usd = float(pd.to_numeric(grp["AdjTaxUSD"], errors="coerce").sum())
    return grp, rates, total_adj_usd, report_total_usd

def read_fd_txt(fd_file) -> pd.DataFrame:
    """
    读取 Apple FD 明细 TXT：
    - 跳过前三行，第四行是表头
    - 行末多一列时自动丢弃
    - 取必需列：SKU / Extended Partner Share / Partner Share Currency
    """
    text = fd_file.read().decode("utf-8", errors="ignore")
    lines = text.splitlines()
    if len(lines) < 4:
        raise ValueError("FD 文件内容不足，无法解析表头")
    hdr = lines[3].split("\t")
    names = hdr + ["_extra"]
    data_str = "\n".join(lines[4:])
    tx = pd.read_csv(StringIO(data_str), sep="\t", header=None, names=names, engine="python")
    if "_extra" in tx.columns:
        tx = tx.drop(columns=["_extra"])

    tx.columns = [str(c).strip() for c in tx.columns]
    need = ["SKU","Extended Partner Share","Partner Share Currency"]
    for c in need:
        if c not in tx.columns:
            raise ValueError(f"FD 缺少列：{c}")
    tx = tx[need].copy()
    tx["Extended Partner Share"] = numify(tx["Extended Partner Share"])
    tx["Partner Share Currency"] = tx["Partner Share Currency"].astype(str).str.strip().str.upper()
    return tx

def read_mapping_xlsx(map_file) -> pd.DataFrame:
    """
    读取项目-SKU 映射（支持一行多个 SKU，用换行/逗号分隔）
    必需列：项目、SKU
    """
    mp = pd.read_excel(map_file, header=0, engine="openpyxl")
    mp.columns = [str(c).strip() for c in mp.columns]
    if "项目" not in mp.columns or "SKU" not in mp.columns:
        raise ValueError("项目-SKU 映射缺少必要列（项目 / SKU）")
    mp["SKU"] = mp["SKU"].astype(str).str.replace("\r", "\n")
    mp = mp.assign(SKU=mp["SKU"].str.split(r"[\n,]")).explode("SKU")
    mp["SKU"] = mp["SKU"].astype(str).str.strip()
    mp = mp[mp["SKU"] != ""]
    return mp[["项目","SKU"]].drop_duplicates()

# ----------------------------- App UI -----------------------------
st.title("IAP 对账助手 · 在线版（财报汇率优先｜缺失自动补齐）")

col1, col2, col3 = st.columns(3)
with col1:
    report_file = st.file_uploader("上传财报：文件名包含 financial_report 的 .csv", type=["csv"])
with col2:
    fd_file = st.file_uploader("上传 Apple 交易明细：FD_*.txt", type=["txt"])
with col3:
    map_file = st.file_uploader("上传 项目-SKU 映射：.xlsx", type=["xlsx"])

do_run = st.button("🚀 开始计算 / Reconcile")

if do_run:
    try:
        if not (report_file and fd_file and map_file):
            st.error("请同时上传三份文件（财报 CSV、交易 TXT、项目-SKU 映射 XLSX）")
            st.stop()

        # -- 读取三份表
        audit, rates, total_adj_usd, report_total_usd = read_report(csv_file=report_file)
        # 由于 read_report 读取时把文件指针读完了，这里要重置其他文件的指针以便读取
        fd_file.seek(0)
        map_file.seek(0)
        tx = read_fd_txt(fd_file)
        mp = read_mapping_xlsx(map_file)

        # -- 以“财报汇率优先；缺失用在线汇率；忽略 NONE”
        valid_rates = {k: v for k, v in rates.items() if isinstance(v, (float,int)) and np.isfinite(v)}
        if "NONE" in valid_rates: valid_rates.pop("NONE", None)

        # 过滤只保留可换算币种的交易
        tx_in = tx[tx["Partner Share Currency"].isin(valid_rates.keys())].copy()
        tx_in["rate_usd_per_local"] = tx_in["Partner Share Currency"].map(valid_rates).astype(float)
        tx_in["Extended Partner Share USD"] = tx_in["Extended Partner Share"] * tx_in["rate_usd_per_local"]

        # 分摊：(调整+预扣税)×汇率 按交易 USD 占比分摊
        tx_total_usd = float(tx_in["Extended Partner Share USD"].sum())
        alloc = (tx_in["Extended Partner Share USD"] / tx_total_usd) * total_adj_usd if tx_total_usd else 0.0
        tx_in["Cost Allocation (USD)"] = alloc
        tx_in["Net Partner Share (USD)"] = tx_in["Extended Partner Share USD"] + tx_in["Cost Allocation (USD)"]

        # 项目映射
        sku2proj = dict(zip(mp["SKU"], mp["项目"]))
        tx_in["项目"] = tx_in["SKU"].astype(str).map(sku2proj)

        # 汇总
        summary_project = tx_in.groupby("项目", dropna=False)[
            ["Extended Partner Share USD","Cost Allocation (USD)","Net Partner Share (USD)"]
        ].sum().reset_index()

        # 币种对账
        tx_ccy = tx_in.groupby("Partner Share Currency")["Extended Partner Share USD"].sum().rename("交易毛收入USD")
        rp_ccy = audit.set_index("Currency")[["美元收入合计(收入.1)","rate_used"]]
        ccy_recon = rp_ccy.join(tx_ccy, how="outer")
        ccy_recon["差异(交易-财报)"] = ccy_recon["交易毛收入USD"] - ccy_recon["美元收入合计(收入.1)"]

        # 总数
        net_total = float(tx_in["Net Partner Share (USD)"].sum())
        diff = net_total - report_total_usd

        # -- 展示关键数
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("财报美元收入合计（∑收入.1）", f"{report_total_usd:,.2f} USD")
        k2.metric("分摊总额（调整+预扣税 → USD）", f"{total_adj_usd:,.2f} USD")
        k3.metric("交易毛收入 USD 合计", f"{tx_total_usd:,.2f} USD")
        k4.metric("交易净额 USD 合计", f"{net_total:,.2f} USD")
        k5.metric("对账差额（交易净额 − 财报美元收入）", f"{diff:,.2f} USD")

        # -- 导出 Excel
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as xw:
            audit.to_excel(xw, index=False, sheet_name="财报审计(最终)")
            tx_in.to_excel(xw, index=False, sheet_name="逐单结果")
            summary_project.to_excel(xw, index=False, sheet_name="项目汇总")
            ccy_recon.reset_index().rename(columns={"index": "Currency"}).to_excel(
                xw, index=False, sheet_name="币种对账"
            )
        out.seek(0)

        st.download_button(
            label="⬇️ 下载 Excel（对账结果）",
            data=out,
            file_name=f"iap_reconciliation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        with st.expander("查看中间结果 / 调试"):
            st.write("财报审计(前 10 行)：")
            st.dataframe(audit.head(10))
            st.write("逐单结果(前 10 行)：")
            st.dataframe(tx_in.head(10))
            st.write("项目汇总：")
            st.dataframe(summary_project)
            st.write("币种对账：")
            st.dataframe(ccy_recon)

    except Exception as e:
        st.error(f"⚠️ 出错：{e}")
        st.stop()
