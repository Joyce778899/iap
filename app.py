import pandas as pd
import streamlit as st

st.set_page_config(page_title="IAP — ORCAT Online (Debug+AutoHeader)", page_icon="🐞", layout="wide")
st.title("🐞 IAP — ORCAT Online Debug + AutoHeader")

with st.expander("输入要求", expanded=False):
    st.markdown("""
**交易明细（CSV/XLSX）**：列 `Extended Partner Share`、`Partner Share Currency`、`SKU`  
**Apple 财报（CSV/XLSX）**：列 `国家或地区 (货币)`、`总欠款`、`收入.1`（或等价）、`调整`、`预扣税`  
**项目-SKU（XLSX）**：列 `项目`、`SKU`
""")

# ---------- 通用读取 ----------
def read_any(file):
    name = file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(file)
    elif name.endswith((".xlsx", ".xls")):
        return pd.read_excel(file, engine="openpyxl")
    else:
        raise ValueError("仅支持 CSV/XLSX 文件")

# ---------- 智能读取 Apple 财报（自动尝试 0–5 行作为表头） ----------
# ---------- 针对“最新 financial 文件”格式的读取 ----------
def read_report(file):
    """
    读取 Apple 财报（最新格式）：
    - 表头固定在第 3 行：header=2
    - 用 on_bad_lines='skip' 跳过异常行（该文件确实有一行字段数不一致）
    - 两个“收入”列中，最后一个“收入.1”为美元收入
    - 保留关键列并转数值，提取币种代码
    """
    # 1) 读取（容错：跳过坏行）
    df = (pd.read_csv(file, header=2, engine="python", on_bad_lines="skip")
          if file.name.lower().endswith(".csv")
          else pd.read_excel(file, header=2, engine="openpyxl"))

    # 2) 标准化列名与清理
    df.columns = [str(c).strip() for c in df.columns]
    # 某些导出会多出一个尾列 Unnamed: xx，这里直接丢掉
    df = df[[c for c in df.columns if not str(c).startswith("Unnamed")]]

    # 3) 校验关键列
    need = ["国家或地区 (货币)", "总欠款", "调整", "预扣税"]
    for c in need:
        if c not in df.columns:
            raise ValueError(f"❌ 财报缺少关键列：{c}")

    # 4) 确认美元收入列：收入.1 是美元收入（你这份文件即如此）
    if "收入.1" not in df.columns:
        # 极端情况兜底：如果没有，就尝试从右往左第一个“收入”
        income_like = [c for c in df.columns if c.startswith("收入")]
        if not income_like:
            raise ValueError("❌ 未找到美元收入列：收入.1")
        # 取最后一个“收入”作为美元收入
        last_income = income_like[-1]
        df["收入.1"] = df[last_income]

    # 5) 数值化
    for c in ["总欠款", "收入.1", "调整", "预扣税"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    # 6) 提取三位币种代码（从 “国家或地区 (货币)”）
    df["Currency"] = df["国家或地区 (货币)"].astype(str).str.extract(r"\((\w{3})\)")
    df = df.dropna(subset=["Currency"])

    # 7) 只返回下游需要的列
    return df[["Currency", "总欠款", "收入.1", "调整", "预扣税"]]

def build_rates(df_report):
    valid = df_report[(df_report["收入.1"].notna()) & (df_report["收入.1"] != 0)]
    if valid.empty:
        raise ValueError("❌ 财报中 '收入.1' 全为 0/空，无法推导汇率")
    # 汇率：总欠款 / 收入.1
    rates = dict(zip(valid["Currency"], (valid["总欠款"] / valid["收入.1"]).astype(float)))
    # 计算分摊总额（调整+预扣税）按币种换算到 USD 再求和
    df = df_report.copy()
    df["rate"] = df["Currency"].map(rates)
    df["AdjTaxUSD"] = (df["调整"].fillna(0) + df["预扣税"].fillna(0)) / df["rate"]
    df["AdjTaxUSD"] = pd.to_numeric(df["AdjTaxUSD"], errors="coerce").fillna(0)
    total_adj_usd = float(df["AdjTaxUSD"].sum())
    report_total_usd = float(pd.to_numeric(df["收入.1"], errors="coerce").sum())
    return rates, total_adj_usd, report_total_usd

# ---------- 交易表 ----------
def read_tx(file):
    df = read_any(file)
    df.columns = [str(c).strip() for c in df.columns]
    st.write("📊 交易表列名：", list(df.columns))
    st.dataframe(df.head())
    need = {"Extended Partner Share", "Partner Share Currency", "SKU"}
    if not need.issubset(df.columns):
        raise ValueError(f"❌ 交易表缺少列：{need - set(df.columns)}")
    df["Extended Partner Share"] = pd.to_numeric(df["Extended Partner Share"], errors="coerce")
    return df

# ---------- 映射表 ----------
def read_map(file):
    df = pd.read_excel(file, engine="openpyxl", dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    st.write("📊 映射表列名：", list(df.columns))
    st.dataframe(df.head())
    if not {"项目", "SKU"}.issubset(df.columns):
        raise ValueError("❌ 映射表缺少列 `项目` 或 `SKU`")
    df = df.assign(SKU=df["SKU"].astype(str).str.split("\n")).explode("SKU")
    df["SKU"] = df["SKU"].str.strip()
    return df[df["SKU"] != ""][["项目", "SKU"]]

# ---------- 上传控件 ----------
c1, c2, c3 = st.columns(3)
with c1: tx = st.file_uploader("① 交易明细（CSV/XLSX）", type=["csv", "xlsx", "xls"], key="tx")
with c2: rp = st.file_uploader("② Apple 财报（CSV/XLSX）", type=["csv", "xlsx", "xls"], key="rp")
with c3: mp = st.file_uploader("③ 项目-SKU（XLSX）", type=["xlsx", "xls"], key="mp")

if st.button("🚀 开始计算 (Debug+AutoHeader)"):
    if not (tx and rp and mp):
        st.error("❌ 三份文件没有全部上传")
    else:
        try:
            # 财报
            rep = read_report(rp)
            st.subheader("📊 Apple 财报预览")
            st.write("列名：", list(rep.columns))
            st.dataframe(rep.head())
            rates, total_adj_usd, report_total_usd = build_rates(rep)

            # 交易
            txdf = read_tx(tx)
            txdf["Extended Partner Share USD"] = txdf.apply(
                lambda r: (r["Extended Partner Share"] / rates.get(str(r["Partner Share Currency"]), 1))
                          if pd.notnull(r["Extended Partner Share"]) else None,
                axis=1
            )
            total_usd = pd.to_numeric(txdf["Extended Partner Share USD"], errors="coerce").sum(min_count=1)
            if not pd.notnull(total_usd) or total_usd == 0:
                st.error("❌ 交易 USD 汇总为 0，可能币种不匹配或金额列为空")
                st.stop()

            # 映射
            mpdf = read_map(mp)
            sku2proj = dict(zip(mpdf["SKU"], mpdf["项目"]))
            txdf["项目"] = txdf["SKU"].map(sku2proj)

            # 分摊与净额
            txdf["Cost Allocation (USD)"] = txdf["Extended Partner Share USD"] / total_usd * total_adj_usd
            txdf["Net Partner Share (USD)"] = txdf["Extended Partner Share USD"] + txdf["Cost Allocation (USD)"]

            # 汇总
            summary = txdf.groupby("项目", dropna=False)[
                ["Extended Partner Share USD", "Cost Allocation (USD)", "Net Partner Share (USD)"]
            ].sum().reset_index()
            total_row = {
                "项目": "__TOTAL__",
                "Extended Partner Share USD": float(summary["Extended Partner Share USD"].sum()),
                "Cost Allocation (USD)": float(summary["Cost Allocation (USD)"].sum()),
                "Net Partner Share (USD)": float(summary["Net Partner Share (USD)"].sum())
            }
            summary = pd.concat([summary, pd.DataFrame([total_row])], ignore_index=True)

            # 结果与下载
            st.success("✅ 计算完成")
            st.markdown(f"- 财报美元收入合计（sum of 收入.1）：**{report_total_usd:,.2f} USD**")
            st.markdown(f"- 分摊总额（调整+预扣税）：**{total_adj_usd:,.2f} USD**")
            st.markdown(f"- 交易毛收入 USD 合计：**{float(total_usd):,.2f} USD**")

            st.download_button("⬇️ 下载 逐单结果 CSV",
                               data=txdf.to_csv(index=False).encode("utf-8-sig"),
                               file_name="transactions_usd_net_project.csv", mime="text/csv")
            st.download_button("⬇️ 下载 项目汇总 CSV",
                               data=summary.to_csv(index=False).encode("utf-8-sig"),
                               file_name="project_summary.csv", mime="text/csv")

            with st.expander("预览：逐单结果", expanded=False):
                st.dataframe(txdf.head(100))
            with st.expander("预览：项目汇总", expanded=True):
                st.dataframe(summary)

        except Exception as e:
            st.error(f"⚠️ 出现错误：{e}")
            st.exception(e)
