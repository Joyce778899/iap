# -*- coding: utf-8 -*-
"""
IAP ORCAT Pipeline — 全量替换版
- 支持 Apple 财报 CSV/XLSX，自动识别表头（前 0~5 行）
- 自动修复“收入.1”缺失（回退到“收入”或包含“收入”的列）
- 规范关键列：国家或地区 (货币) / 总欠款 / 收入.1 / 调整 / 预扣税
- 从财报推导各币种汇率：rate = 总欠款 / 收入.1
- 计算总分摊(调整+预扣税)的 USD，并按交易占比分摊
- 交易表支持 CSV/XLSX；映射表 XLSX（SKU 可多行）
- 输出：逐单、项目汇总、运行日志
"""
import os
import sys
import argparse
import pandas as pd

# -----------------------
# 基础工具
# -----------------------
def _read_any(path, header=None, dtype=None):
    ext = os.path.splitext(path)[1].lower()
    if ext in [".xlsx", ".xls"]:
        return pd.read_excel(path, header=header, dtype=dtype, engine="openpyxl")
    else:
        # CSV 默认 utf-8-sig；失败再尝试 latin1
        try:
            return pd.read_csv(path, header=header, dtype=dtype, low_memory=False, encoding="utf-8-sig")
        except Exception:
            return pd.read_csv(path, header=header, dtype=dtype, low_memory=False, encoding="latin1")

def _coerce_numeric(df, cols):
    for c in cols:
        if c not in df.columns:
            df[c] = 0
        df[c] = pd.to_numeric(df[c], errors="coerce")

def _ensure_cols(df, need, name_for_log):
    missing = [c for c in need if c not in df.columns]
    if missing:
        raise ValueError(f"{name_for_log} 缺少列：{missing}")

# -----------------------
# 财报读取与修复
# -----------------------
def read_financial_report(report_path, debug=False):
    """
    智能读取 Apple 财报：
    - 自动尝试 header=0..5，寻找包含“国家或地区”和“货币”的列（通常是“国家或地区 (货币)”）
    - 修复“收入.1”列，回退到“收入”或其它包含“收入”的列
    - 转数值并提取 Currency
    """
    # 自动识别表头（优先找到包含“货币”的列）
    df = None
    target_col = None
    for h in range(6):
        try:
            tmp = _read_any(report_path, header=h, dtype=str)
            tmp.columns = [str(c).strip() for c in tmp.columns]
            # 寻找币种列
            cand = [c for c in tmp.columns if ("国家或地区" in c and "货币" in c) or ("货币" in c)]
            if cand:
                df = tmp
                target_col = cand[0]
                break
        except Exception:
            continue
    if df is None:
        # 兜底再读一次（不指定 header），让错误更可解释
        df = _read_any(report_path, dtype=str)
        df.columns = [str(c).strip() for c in df.columns]
        cand = [c for c in df.columns if ("国家或地区" in c and "货币" in c) or ("货币" in c)]
        if cand:
            target_col = cand[0]
        else:
            raise ValueError("无法识别财报表头：未找到包含“货币”的列。请检查文件内容。")

    # 关键列修复：收入.1
    if "收入.1" not in df.columns:
        alt = [c for c in df.columns if ("收入" in c and c != "收入")]
        if alt:
            df["收入.1"] = df[alt[0]]
            if debug: print(f"[report] 使用列 {alt[0]} 作为 收入.1")
        elif "收入" in df.columns:
            df["收入.1"] = df["收入"]
            if debug: print(f"[report] 使用列 收入 作为 收入.1")
        else:
            raise ValueError("财报未找到 '收入.1' 或等价列（收入/包含“收入”的列）。")

    # 转换数值列
    num_cols = ["总欠款", "收入.1", "调整", "预扣税"]
    # 缺失的“调整/预扣税”填 0，其它缺失报错（下面统一按 0 处理再校验）
    for c in num_cols:
        if c not in df.columns and c in ["调整", "预扣税"]:
            df[c] = 0
    _coerce_numeric(df, num_cols)

    # 提取 Currency
    if target_col is None:
        raise ValueError("财报缺少“国家或地区 (货币)”列。")
    df["Currency"] = df[target_col].astype(str).str.extract(r"\((\w+)\)")
    df = df.dropna(subset=["Currency"])

    if debug:
        print("[report] 列名：", list(df.columns)[:20])
        print("[report] 预览：")
        print(df.head(3).to_string())

    return df[["Currency", "总欠款", "收入.1", "调整", "预扣税"]]

def build_rates_and_totals(df_report, debug=False):
    """
    从财报推导：
    - 各币种汇率 rate = 总欠款 / 收入.1
    - 分摊总额(USD) = (调整 + 预扣税) / rate（逐币种求和）
    - 财报美元收入总额（sum of 收入.1）
    """
    valid = df_report[(df_report["收入.1"].notna()) & (df_report["收入.1"] != 0)]
    if valid.empty:
        raise ValueError("财报中 '收入.1' 全为 0/空，无法推导汇率。")

    rates = (valid["总欠款"] / valid["收入.1"]).groupby(valid["Currency"]).first().to_dict()

    # 逐币种换算 USD 的调整+预扣税，然后求和
    df = df_report.copy()
    df["rate"] = df["Currency"].map(rates)
    df["AdjTaxUSD"] = (df["调整"].fillna(0) + df["预扣税"].fillna(0)) / df["rate"]
    df["AdjTaxUSD"] = pd.to_numeric(df["AdjTaxUSD"], errors="coerce").fillna(0)

    total_adj_usd = float(df["AdjTaxUSD"].sum())
    report_total_usd = float(pd.to_numeric(df["收入.1"], errors="coerce").sum())

    if debug:
        print(f"[report] 汇率个数: {len(rates)}")
        print(f"[report] 分摊总额(USD): {total_adj_usd:,.2f}")
        print(f"[report] 财报 USD 收入总额: {report_total_usd:,.2f}")

    return rates, total_adj_usd, report_total_usd

# -----------------------
# 交易表 + 映射
# -----------------------
def read_transactions(tx_path, debug=False):
    """
    交易表：CSV/XLSX
    必需列：Extended Partner Share / Partner Share Currency / SKU
    """
    # 容错：尝试 0..3 行作为 header
    df = None
    need = {"Extended Partner Share", "Partner Share Currency", "SKU"}
    for h in range(4):
        try:
            tmp = _read_any(tx_path, header=h)
            tmp.columns = [str(c).strip() for c in tmp.columns]
            if need.issubset(tmp.columns):
                df = tmp
                break
        except Exception:
            continue
    if df is None:
        df = _read_any(tx_path)
        df.columns = [str(c).strip() for c in df.columns]
        _ensure_cols(df, need, "交易表")

    df["Extended Partner Share"] = pd.to_numeric(df["Extended Partner Share"], errors="coerce")
    if debug:
        print("[tx] 列名：", list(df.columns)[:20])
        print("[tx] 预览：")
        print(df.head(3).to_string())
    return df

def read_mapping(mapping_path, debug=False):
    """
    映射表：XLSX（需列：项目 / SKU；SKU 可换行多值）
    """
    df = pd.read_excel(mapping_path, dtype=str, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    _ensure_cols(df, ["项目", "SKU"], "映射表")

    df = df.assign(SKU=df["SKU"].astype(str).str.split("\n")).explode("SKU")
    df["SKU"] = df["SKU"].str.strip()
    df = df[df["SKU"] != ""]
    if debug:
        print("[map] 预览：")
        print(df.head(3).to_string())
    return df[["项目", "SKU"]]

# -----------------------
# 主流程
# -----------------------
def process(tx_path, report_path, mapping_path, outdir="output", debug=False):
    os.makedirs(outdir, exist_ok=True)

    # 1) 财报
    rep = read_financial_report(report_path, debug=debug)
    rates, total_adj_usd, report_total_usd = build_rates_and_totals(rep, debug=debug)

    # 2) 交易
    tx = read_transactions(tx_path, debug=debug)
    # 汇率换算（毛收入 USD）
    tx["Extended Partner Share USD"] = tx.apply(
        lambda r: (r["Extended Partner Share"] / rates.get(str(r["Partner Share Currency"]), 1))
        if pd.notnull(r["Extended Partner Share"]) else None,
        axis=1,
    )
    total_usd = pd.to_numeric(tx["Extended Partner Share USD"], errors="coerce").sum(min_count=1)
    if not pd.notnull(total_usd) or total_usd == 0:
        raise ValueError("交易表 USD 汇总为 0，请检查币种与金额列是否正确。")

    # 3) 分摊：按交易 USD 占比分摊总成本
    tx["Cost Allocation (USD)"] = tx["Extended Partner Share USD"] / total_usd * total_adj_usd
    tx["Net Partner Share (USD)"] = tx["Extended Partner Share USD"] + tx["Cost Allocation (USD)"]

    # 4) 映射项目
    mp = read_mapping(mapping_path, debug=debug)
    sku2proj = dict(zip(mp["SKU"], mp["项目"]))
    tx["项目"] = tx["SKU"].map(sku2proj)

    # 5) 输出
    out_tx = os.path.join(outdir, "transactions_usd_net_project.csv")
    out_sum = os.path.join(outdir, "project_summary.csv")
    out_log = os.path.join(outdir, "run_log.txt")

    tx.to_csv(out_tx, index=False, encoding="utf-8-sig")

    summary = tx.groupby("项目", dropna=False)[
        ["Extended Partner Share USD", "Cost Allocation (USD)", "Net Partner Share (USD)"]
    ].sum().reset_index()

    # 总计行
    total_row = {
        "项目": "__TOTAL__",
        "Extended Partner Share USD": float(summary["Extended Partner Share USD"].sum()),
        "Cost Allocation (USD)": float(summary["Cost Allocation (USD)"].sum()),
        "Net Partner Share (USD)": float(summary["Net Partner Share (USD)"].sum()),
    }
    summary = pd.concat([summary, pd.DataFrame([total_row])], ignore_index=True)
    summary.to_csv(out_sum, index=False, encoding="utf-8-sig")

    with open(out_log, "w", encoding="utf-8") as f:
        f.write("=== IAP ORCAT PIPELINE LOG ===\n")
        f.write(f"Report Total USD (sum of 收入.1): {report_total_usd:,.2f}\n")
        f.write(f"Adj+Withholding Total USD: {total_adj_usd:,.2f}\n")
        f.write(f"TX Total USD (before allocation): {float(total_usd):,.2f}\n")
        f.write(f"TX Net USD (after allocation): {float(pd.to_numeric(tx['Net Partner Share (USD)'], errors='coerce').sum()):,.2f}\n")

    if debug:
        print("[done] 输出文件：")
        print(" -", out_tx)
        print(" -", out_sum)
        print(" -", out_log)

# -----------------------
# CLI
# -----------------------
def main():
    ap = argparse.ArgumentParser(description="IAP ORCAT Pipeline — 全量替换版")
    ap.add_argument("--tx", required=True, help="交易表 CSV/XLSX")
    ap.add_argument("--report", required=True, help="Apple 财报 CSV/XLSX（支持自动识别表头）")
    ap.add_argument("--mapping", required=True, help="项目-SKU 映射（XLSX）")
    ap.add_argument("--outdir", default="output", help="输出目录（默认 output）")
    ap.add_argument("--debug", action="store_true", help="打印调试信息")
    args = ap.parse_args()

    process(args.tx, args.report, args.mapping, args.outdir, debug=args.debug)

if __name__ == "__main__":
    main()
