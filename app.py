import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="ORCAT Online Pipeline", page_icon="💼", layout="wide")
st.title("💼 ORCAT Online Pipeline — IAP")
st.write("上传三份文件，在线计算：逐单USD、分摊成本、净收入、按项目汇总。")

with st.expander("输入要求（请仔细核对列名）", expanded=False):
    st.markdown("""
**1) 交易明细（CSV 或 XLSX）** 必含列：
- `Extended Partner Share`（本币）  
- `Partner Share Currency`（币种）  
- `SKU`  

**2) Apple 财务报表（CSV 或 XLSX）** 必含列：
- `国家或地区 (货币)`  
- `总欠款`（本币）  
- `收入.1`（美元收入；或任何替代列，见下）  
- `调整`、`预扣税`（若缺则按 0 处理）  

**3) 项目-SKU 映射（XLSX）** 必含列：
- `项目`、`SKU`（SKU 可用换行分隔多个）
""")

def read_any_table(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    elif name.endswith(".xlsx") or name.endswith(".xls"):
        return pd.read_excel(uploaded_file, engine="openpyxl")
    else:
        raise ValueError("仅支持 CSV/XLSX 文件")

def read_financial_report(file) -> pd.DataFrame:
    df_raw = read_any_table(file)
    # 若没有目标列名，尝试把某一行当表头（最多前4行）
    if "国家或地区 (货币)" not in df_raw.columns:
        for header in range(1, 4):
            try:
                file.seek(0)
                df_try = pd.read_csv(file, header=header) if file.name.lower().endswith(".csv") else pd.read_excel(file, header=header, engine="openpyxl")
                if "国家或地区 (货币)" in df_try.columns:
                    df_raw = df_try
                    break
            except Exception:
                pass
        file.seek(0)
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # 处理收入列
    if "收入.1" not in df.columns:
        alt = [c for c in df.columns if ("收入" in c and c != "收入")]
        if alt:
            df["收入.1"] = df[alt[0]]
        elif "收入" in df.columns:
            df["收入.1"] = df["收入"]
        else:
            raise ValueError("财报未找到美元收入列（收入.1）。")

    for col in ["总欠款", "收入.1", "调整", "预扣税"]:
        if col not in df.columns:
            if col in ["调整", "预扣税"]:
                df[col] = 0
            else:
                raise ValueError(f"财报缺少列：{col}")
        df[col] = pd.to_numeric(df[col], errors="coerce")

    if "国家或地区 (货币)" not in df.columns:
        raise ValueError("财报缺少列：国家或地区 (货币)")
    df["Currency"] = df["国家或地区 (货币)"].astype(str).str.extract(r"\((\w+)\)")
    df = df.dropna(subset=["Currency"])
    return df

def build_rates_and_total_adj_usd(df_report: pd.DataFrame):
    valid = df_report[(df_report["收入.1"].notna()) & (df_report["收入.1"] != 0)]
    rates = dict(zip(valid["Currency"], valid["总欠款"] / valid["收入.1"]))
    df_report = df_report.copy()
    df_report["rate"] = df_report["Currency"].map(rates)
    df_report["AdjTaxUSD"] = (df_report["调整"].fillna(0) + df_report["预扣税"].fillna(0)) / df_report["rate"]
    df_report["AdjTaxUSD"] = df_report["AdjTaxUSD"].fillna(0)
    total_adj_usd = float(df_report["AdjTaxUSD"].sum())
    report_total_usd = float(pd.to_numeric(df_report["收入.1"], errors="coerce").sum())
    return rates, total_adj_usd, report_total_usd

def read_tx(file) -> pd.DataFrame:
    df_raw = read_any_table(file)
    # 尝试不同的 header 行
    if not {"Extended Partner Share", "Partner Share Currency", "SKU"}.issubset(set(df_raw.columns)):
        for header in range(1, 4):
            try:
                file.seek(0)
                df_try = pd.read_csv(file, header=header) if file.name.lower().endswith(".csv") else pd.read_excel(file, header=header, engine="openpyxl")
                if {"Extended Partner Share", "Partner Share Currency", "SKU"}.issubset(set(df_try.columns)):
                    df_raw = df_try
                    break
            except Exception:
                pass
        file.seek(0)
    needed = {"Extended Partner Share", "Partner Share Currency", "SKU"}
    if not needed.issubset(set(df_raw.columns)):
        raise ValueError(f"交易表缺少列：{needed - set(df_raw.columns)}")
    df_raw["Extended Partner Share"] = pd.to_numeric(df_raw["Extended Partner Share"], errors="coerce")
    return df_raw

def read_mapping(file) -> pd.DataFrame:
    df_map = pd.read_excel(file, engine="openpyxl", dtype=str)
    if "项目" not in df_map.columns or "SKU" not in df_map.columns:
        raise ValueError("项目映射需包含列：'项目' 与 'SKU'")
    df_map = df_map.assign(SKU=df_map["SKU"].astype(str).str.split("\n")).explode("SKU")
    df_map["SKU"] = df_map["SKU"].str.strip()
    df_map = df_map[df_map["SKU"] != ""]
    return df_map[["项目", "SKU"]]

# Upload widgets
col1, col2, col3 = st.columns(3)
with col1:
    tx_file = st.file_uploader("① 交易明细（CSV/XLSX）", type=["csv", "xlsx", "xls"], key="tx")
with col2:
    report_file = st.file_uploader("② Apple 财报（CSV/XLSX）", type=["csv", "xlsx", "xls"], key="report")
with col3:
    mapping_file = st.file_uploader("③ 项目-SKU（XLSX）", type=["xlsx", "xls"], key="mapping")

run = st.button("🚀 开始计算")

if run:
    try:
        if not (tx_file and report_file and mapping_file):
            st.error("请先上传三份文件。")
        else:
            df_report = read_financial_report(report_file)
            rates, total_adj_usd, report_total_usd = build_rates_and_total_adj_usd(df_report)
            df_tx = read_tx(tx_file)
            df_map = read_mapping(mapping_file)

            df_tx["Extended Partner Share USD"] = df_tx.apply(
                lambda r: (r["Extended Partner Share"] / rates.get(str(r["Partner Share Currency"]), 1))
                          if pd.notnull(r["Extended Partner Share"]) else None,
                axis=1
            )
            total_usd = pd.to_numeric(df_tx["Extended Partner Share USD"], errors="coerce").sum(min_count=1)
            if not pd.notnull(total_usd) or total_usd == 0:
                st.error("交易表 USD 汇总为 0，可能币种不匹配或数据为空。")
            else:
                df_tx["Cost Allocation (USD)"] = df_tx["Extended Partner Share USD"] / total_usd * total_adj_usd
                df_tx["Net Partner Share (USD)"] = df_tx["Extended Partner Share USD"] + df_tx["Cost Allocation (USD)"]

                sku2proj = dict(zip(df_map["SKU"], df_map["项目"]))
                df_tx["项目"] = df_tx["SKU"].map(sku2proj)

                summary = df_tx.groupby("项目", dropna=False)[
                    ["Extended Partner Share USD", "Cost Allocation (USD)", "Net Partner Share (USD)"]
                ].sum().reset_index()
                total_row = {
                    "项目": "__TOTAL__",
                    "Extended Partner Share USD": float(summary["Extended Partner Share USD"].sum()),
                    "Cost Allocation (USD)": float(summary["Cost Allocation (USD)"].sum()),
                    "Net Partner Share (USD)": float(summary["Net Partner Share (USD)"].sum())
                }
                summary = pd.concat([summary, pd.DataFrame([total_row])], ignore_index=True)

                st.success("✅ 计算完成")
                st.markdown(f"- 财报美元收入合计（sum of 收入.1）≈ **{report_total_usd:,.2f} USD**")
                st.markdown(f"- 分摊总额（调整+预扣税）≈ **{total_adj_usd:,.2f} USD**")
                st.markdown(f"- 交易毛收入 USD 合计 ≈ **{float(total_usd):,.2f} USD**")

                out_tx = df_tx.to_csv(index=False).encode("utf-8-sig")
                out_summary = summary.to_csv(index=False).encode("utf-8-sig")
                st.download_button("⬇️ 下载 逐单结果 CSV", data=out_tx, file_name="transactions_usd_net_project.csv", mime="text/csv")
                st.download_button("⬇️ 下载 项目汇总 CSV", data=out_summary, file_name="project_summary.csv", mime="text/csv")

                with st.expander("查看预览：逐单结果", expanded=False):
                    st.dataframe(df_tx.head(100))
                with st.expander("查看预览：项目汇总", expanded=True):
                    st.dataframe(summary)

    except Exception as e:
        st.error(f"发生错误：{e}")
        st.exception(e)
