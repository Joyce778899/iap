import pandas as pd
import streamlit as st

st.set_page_config(page_title="IAP — ORCAT Online (DEBUG)", page_icon="🐞", layout="wide")
st.title("🐞 IAP — ORCAT Online Debug Version")

with st.expander("输入要求", expanded=False):
    st.markdown("""
**交易明细（CSV/XLSX）**：列 `Extended Partner Share`、`Partner Share Currency`、`SKU`  
**Apple 财报（CSV/XLSX）**：列 `国家或地区 (货币)`、`总欠款`、`收入.1`（或等价）、`调整`、`预扣税`  
**项目-SKU（XLSX）**：列 `项目`、`SKU`
""")

def read_any(file):
    name = file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(file)
    elif name.endswith((".xlsx", ".xls")):
        return pd.read_excel(file, engine="openpyxl")
    else:
        raise ValueError("仅支持 CSV/XLSX")

# --- 上传控件 ---
c1, c2, c3 = st.columns(3)
with c1: tx = st.file_uploader("① 交易明细", type=["csv","xlsx","xls"], key="tx")
with c2: rp = st.file_uploader("② Apple 财报", type=["csv","xlsx","xls"], key="rp")
with c3: mp = st.file_uploader("③ 项目-SKU", type=["xlsx","xls"], key="mp")

if st.button("🚀 开始计算 (Debug)"):

    # --- 检查文件 ---
    if not (tx and rp and mp):
        st.error("❌ 三份文件没有全部上传")
    else:
        try:
            # --- Apple 报表 ---
            rep = read_any(rp)
            st.subheader("📊 Apple 报表预览")
            st.write("列名：", list(rep.columns))
            st.dataframe(rep.head())

            if "收入.1" not in rep.columns:
                st.warning("⚠️ 未检测到 `收入.1` 列，尝试用其它含'收入'的列代替")
                alt = [c for c in rep.columns if ("收入" in c and c != "收入")]
                if alt:
                    rep["收入.1"] = rep[alt[0]]
                    st.info(f"已自动使用列：{alt[0]} 作为收入.1")
                elif "收入" in rep.columns:
                    rep["收入.1"] = rep["收入"]
                    st.info("已自动使用列：收入 作为收入.1")
                else:
                    st.error("❌ 财报没有找到 `收入.1` 或等价列")
                    st.stop()

            # 转数值
            for c in ["总欠款","收入.1","调整","预扣税"]:
                if c not in rep.columns:
                    st.warning(f"⚠️ 财报缺少列 {c}，已补0")
                    rep[c] = 0
                rep[c] = pd.to_numeric(rep[c], errors="coerce")

            rep["Currency"] = rep["国家或地区 (货币)"].astype(str).str.extract(r"\((\w+)\)")
            rep = rep.dropna(subset=["Currency"])
            rates = dict(zip(rep["Currency"], rep["总欠款"]/rep["收入.1"]))
            rep["rate"] = rep["Currency"].map(rates)
            rep["AdjTaxUSD"] = (rep["调整"].fillna(0)+rep["预扣税"].fillna(0))/rep["rate"]
            adj_usd = rep["AdjTaxUSD"].sum()
            report_total = pd.to_numeric(rep["收入.1"], errors="coerce").sum()

            # --- 交易表 ---
            txdf = read_any(tx)
            st.subheader("📊 交易表预览")
            st.write("列名：", list(txdf.columns))
            st.dataframe(txdf.head())

            need = {"Extended Partner Share","Partner Share Currency","SKU"}
            if not need.issubset(txdf.columns):
                st.error(f"❌ 交易表缺少列：{need - set(txdf.columns)}")
                st.stop()

            txdf["Extended Partner Share"] = pd.to_numeric(txdf["Extended Partner Share"], errors="coerce")
            txdf["Extended Partner Share USD"] = txdf.apply(
                lambda r: (r["Extended Partner Share"]/rates.get(str(r["Partner Share Currency"]),1)) if pd.notnull(r["Extended Partner Share"]) else None,
                axis=1
            )
            total_usd = pd.to_numeric(txdf["Extended Partner Share USD"], errors="coerce").sum(min_count=1)
            if total_usd == 0:
                st.error("❌ 交易表 USD 汇总为 0，可能币种不匹配")
                st.stop()

            # --- 映射表 ---
            mpdf = pd.read_excel(mp, engine="openpyxl", dtype=str)
            st.subheader("📊 映射表预览")
            st.write("列名：", list(mpdf.columns))
            st.dataframe(mpdf.head())

            if not {"项目","SKU"}.issubset(mpdf.columns):
                st.error("❌ 映射表缺少列 `项目` 或 `SKU`")
                st.stop()
            mpdf = mpdf.assign(SKU=mpdf["SKU"].astype(str).str.split("\n")).explode("SKU")
            mpdf["SKU"] = mpdf["SKU"].str.strip()
            sku2proj = dict(zip(mpdf["SKU"], mpdf["项目"]))
            txdf["项目"] = txdf["SKU"].map(sku2proj)

            # --- 汇总 ---
            txdf["Cost Allocation (USD)"] = txdf["Extended Partner Share USD"]/total_usd*adj_usd
            txdf["Net Partner Share (USD)"] = txdf["Extended Partner Share USD"]+txdf["Cost Allocation (USD)"]

            summary = txdf.groupby("项目", dropna=False)[
                ["Extended Partner Share USD","Cost Allocation (USD)","Net Partner Share (USD)"]
            ].sum().reset_index()

            st.success("✅ 计算完成")
            st.write(f"财报 USD 合计: {report_total:,.2f} | 分摊总额: {adj_usd:,.2f} | 交易 USD 合计: {total_usd:,.2f}")

            st.subheader("📑 项目汇总")
            st.dataframe(summary)

        except Exception as e:
            st.error(f"⚠️ 出现错误: {e}")
            st.exception(e)
