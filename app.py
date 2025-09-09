import pandas as pd
import streamlit as st

st.set_page_config(page_title="IAP — ORCAT Online", page_icon="💼", layout="wide")
st.title("💼 IAP — ORCAT Online")

with st.expander("输入要求", expanded=False):
    st.markdown("""
**交易明细（CSV/XLSX）**：列 `Extended Partner Share`、`Partner Share Currency`、`SKU`  
**Apple 财报（CSV/XLSX）**：列 `国家或地区 (货币)`、`总欠款`、`收入.1`（或等价）、`调整`、`预扣税`  
**项目-SKU（XLSX）**：列 `项目`、`SKU`（SKU 可换行多个）
""")

def read_any(file):
    name = file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(file)
    elif name.endswith((".xlsx", ".xls")):
        return pd.read_excel(file, engine="openpyxl")
    else:
        raise ValueError("仅支持 CSV/XLSX")

def read_report(f):
    df = read_any(f)
    df.columns = [str(c).strip() for c in df.columns]
    if "收入.1" not in df.columns:
        alt = [c for c in df.columns if ("收入" in c and c != "收入")]
        if alt: df["收入.1"] = df[alt[0]]
        elif "收入" in df.columns: df["收入.1"] = df["收入"]
        else: st.stop()
    for c in ["总欠款","收入.1","调整","预扣税"]:
        if c not in df.columns: df[c] = 0 if c in ["调整","预扣税"] else None
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["Currency"] = df["国家或地区 (货币)"].astype(str).str.extract(r"\((\w+)\)")
    df = df.dropna(subset=["Currency"])
    return df

def build_rates(df):
    valid = df[(df["收入.1"].notna()) & (df["收入.1"]!=0)]
    rates = dict(zip(valid["Currency"], valid["总欠款"]/valid["收入.1"]))
    df["rate"] = df["Currency"].map(rates)
    df["AdjTaxUSD"] = (df["调整"].fillna(0)+df["预扣税"].fillna(0))/df["rate"]
    df["AdjTaxUSD"] = df["AdjTaxUSD"].fillna(0)
    return rates, float(df["AdjTaxUSD"].sum()), float(pd.to_numeric(df["收入.1"], errors="coerce").sum())

def read_tx(f):
    df = read_any(f)
    need = {"Extended Partner Share","Partner Share Currency","SKU"}
    if not need.issubset(df.columns): st.stop()
    df["Extended Partner Share"] = pd.to_numeric(df["Extended Partner Share"], errors="coerce")
    return df

def read_map(f):
    df = pd.read_excel(f, engine="openpyxl", dtype=str)
    df = df.assign(SKU=df["SKU"].astype(str).str.split("\n")).explode("SKU")
    df["SKU"] = df["SKU"].str.strip()
    return df[df["SKU"]!=""][["项目","SKU"]]

c1,c2,c3 = st.columns(3)
with c1: tx = st.file_uploader("① 交易明细", type=["csv","xlsx","xls"], key="tx")
with c2: rp = st.file_uploader("② Apple 财报", type=["csv","xlsx","xls"], key="rp")
with c3: mp = st.file_uploader("③ 项目-SKU", type=["xlsx","xls"], key="mp")
if st.button("🚀 开始计算"):
    if not (tx and rp and mp):
        st.error("请先上传三份文件"); st.stop()
    rep = read_report(rp)
    rates, adj_usd, report_total = build_rates(rep)
    txdf = read_tx(tx)
    mpdf = read_map(mp)
    txdf["Extended Partner Share USD"] = txdf.apply(
        lambda r: (r["Extended Partner Share"]/rates.get(str(r["Partner Share Currency"]),1)) if pd.notnull(r["Extended Partner Share"]) else None, axis=1
    )
    total_usd = pd.to_numeric(txdf["Extended Partner Share USD"], errors="coerce").sum(min_count=1)
    if not pd.notnull(total_usd) or total_usd==0: st.error("交易 USD 汇总为 0"); st.stop()
    txdf["Cost Allocation (USD)"] = txdf["Extended Partner Share USD"]/total_usd*adj_usd
    txdf["Net Partner Share (USD)"] = txdf["Extended Partner Share USD"]+txdf["Cost Allocation (USD)"]
    sku2proj = dict(zip(mpdf["SKU"], mpdf["项目"]))
    txdf["项目"] = txdf["SKU"].map(sku2proj)
    summary = txdf.groupby("项目", dropna=False)[
        ["Extended Partner Share USD","Cost Allocation (USD)","Net Partner Share (USD)"]
    ].sum().reset_index()
    total_row = {"项目":"__TOTAL__",
                 "Extended Partner Share USD": float(summary["Extended Partner Share USD"].sum()),
                 "Cost Allocation (USD)": float(summary["Cost Allocation (USD)"].sum()),
                 "Net Partner Share (USD)": float(summary["Net Partner Share (USD)"].sum())}
    summary = pd.concat([summary, pd.DataFrame([total_row])], ignore_index=True)

    st.success("✅ 完成")
    st.markdown(f"- 财报 USD 合计 ≈ **{report_total:,.2f}**")
    st.markdown(f"- 分摊总额（调整+预扣税）≈ **{adj_usd:,.2f}**")
    st.markdown(f"- 交易毛收入 USD 合计 ≈ **{float(total_usd):,.2f}**")

    st.download_button("⬇️ 下载 逐单结果 CSV", data=txdf.to_csv(index=False).encode("utf-8-sig"),
                       file_name="transactions_usd_net_project.csv", mime="text/csv")
    st.download_button("⬇️ 下载 项目汇总 CSV", data=summary.to_csv(index=False).encode("utf-8-sig"),
                       file_name="project_summary.csv", mime="text/csv")

    with st.expander("预览：逐单结果", expanded=False): st.dataframe(txdf.head(100))
    with st.expander("预览：项目汇总", expanded=True): st.dataframe(summary)
