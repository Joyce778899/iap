# app.py
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="IAP — ORCAT Online (Debug+AutoHeader)", page_icon="🐞", layout="wide")
st.title("🐞 IAP — ORCAT Online Debug + AutoHeader")

with st.expander("输入要求（必读）", expanded=False):
    st.markdown("""
**交易明细（CSV/XLSX）**：至少包含能表示
- 金额（本币）：如 *Extended Partner Share / Proceeds / Amount*
- 币种：如 *Partner Share Currency / Currency*
- SKU：如 *SKU / Product ID / Item ID*

**Apple 财报（CSV/XLSX）**（支持最新导出格式）：
- 表头在第 3 行（header=2），或前几行；脚本会自动识别
- 两个“收入”列中，**最后一个 `收入.1` 为美元收入**（若不存在则从包含“收入/usd/revenue/proceeds”的列兜底）
- 关键列：`国家或地区 (货币)`、`总欠款`、`收入.1`（美元）、`调整`、`预扣税`

**项目-SKU（XLSX）**：
- 列：`项目`、`SKU`（SKU 可换行分隔多个）
""")

# ---------- 通用读取 ----------
def read_any(file, header=None):
    name = file.name.lower()
    if name.endswith(".csv"):
        # 对 CSV 使用 python 引擎，支持 on_bad_lines
        return pd.read_csv(file, header=header, engine="python", on_bad_lines="skip")
    elif name.endswith((".xlsx", ".xls")):
        return pd.read_excel(file, header=header, engine="openpyxl")
    else:
        raise ValueError("仅支持 CSV 或 Excel (xlsx/xls) 文件")

# ---------- 财报读取（自动识别表头 + 列名变体，适配你的“最新文件”） ----------
def read_report(file):
    """
    自动尝试 header=2（第三行表头，适配你的最新文件），若失败再回退 0..10；
    处理坏行；识别“收入.1”为美元收入（若不存在则从包含‘收入/usd/revenue/proceeds’的列兜底）。
    返回标准列：Currency, 总欠款, 收入.1, 调整, 预扣税
    """
    df = None
    tried_headers = []

    # 先按你最新文件的习惯：header=2
    try:
        file.seek(0)
        df = read_any(file, header=2)
    except Exception:
        df = None

    # 若失败/不含关键列，再尝试 0..10
    def has_currency_col(_df):
        cols = [str(c).strip() for c in _df.columns]
        return any(("货币" in c) or ("国家或地区" in c) for c in cols)

    if df is None or not has_currency_col(df):
        for h in range(0, 11):
            try:
                file.seek(0)
                temp = read_any(file, header=h)
                tried_headers.append(h)
                if has_currency_col(temp):
                    df = temp
                    break
            except Exception:
                continue

    if df is None:
        raise ValueError(f"❌ 无法识别财报表头（已尝试 header={ [2]+tried_headers }）")

    # 清洗列名与无名列
    df.columns = [str(c).strip() for c in df.columns]
    df = df[[c for c in df.columns if not str(c).startswith("Unnamed")]]

    # 找“国家或地区 (货币)”或包含货币信息的列
    currency_source_col = None
    for c in df.columns:
        if ("国家或地区" in c and "货币" in c):
            currency_source_col = c; break
    if currency_source_col is None:
        for c in df.columns:
            if "货币" in c or "currency" in c.lower():
                currency_source_col = c; break
    if currency_source_col is None:
        # 退一步：仅“国家或地区”列
        for c in df.columns:
            if "国家或地区" in c:
                currency_source_col = c; break
    if currency_source_col is None:
        raise ValueError("❌ 财报未找到包含‘货币’或‘国家或地区’的列")

    # 美元收入列：优先使用“收入.1”，否则从右往左找包含“收入/usd/revenue/proceeds”的列
    revenue_col = None
    if "收入.1" in df.columns:
        revenue_col = "收入.1"
    else:
        candidates = []
        for c in df.columns:
            cl = c.lower()
            if ("收入" in c) or ("usd" in cl) or ("revenue" in cl) or ("proceeds" in cl):
                candidates.append(c)
        if candidates:
            revenue_col = candidates[-1]  # 取最右侧一个
    if revenue_col is None:
        raise ValueError("❌ 财报未找到美元收入列（收入.1/包含收入或usd/revenue/proceeds 的列）")

    # 本币总额列（用于汇率）：“总欠款/欠款/本币金额/本地货币/local total/amount local”
    owed_col = None
    for c in df.columns:
        cl = c.lower()
        if ("总欠款" in c) or ("欠款" in c) or ("本币金额" in c) or ("本地货币" in c) or ("local" in cl and ("total" in cl or "amount" in cl)):
            owed_col = c; break
    if owed_col is None:
        # 有些报表“总欠款”就叫“总额/金额”，此处再宽一点（但尽量避免选到美元收入列）
        for c in df.columns:
            if ("总额" in c or "金额" in c) and c != revenue_col:
                owed_col = c; break
    if owed_col is None:
        raise ValueError("❌ 财报未找到本币总额列（如‘总欠款/本币金额/Local Total’）")

    # 调整/预扣税列（无则置 0）
    adj_col = None
    tax_col = None
    for c in df.columns:
        cl = c.lower()
        if ("调整" in c) or ("adjust" in cl):
            adj_col = c; break
    for c in df.columns:
        cl = c.lower()
        if ("预扣税" in c) or ("withholding" in cl) or ("wht" in cl):
            tax_col = c; break
    if adj_col is None:
        df["__adj__"] = 0; adj_col = "__adj__"
    if tax_col is None:
        df["__tax__"] = 0; tax_col = "__tax__"

    # 数值化
    for c in [owed_col, revenue_col, adj_col, tax_col]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    # 提取三位币种（(XXX) 或 直接三字母）
    def extract_ccy(val):
        s = str(val)
        m = re.search(r"\(([A-Za-z]{3})\)", s)
        if m: return m.group(1).upper()
        m = re.search(r"\b([A-Za-z]{3})\b", s)
        if m: return m.group(1).upper()
        return None

    cur_vals = df[currency_source_col].apply(extract_ccy)
    df = df.assign(
        Currency = cur_vals
    ).dropna(subset=["Currency"])

    out = df.assign(
        **{
            "总欠款": df[owed_col],
            "收入.1": df[revenue_col],
            "调整": df[adj_col],
            "预扣税": df[tax_col],
        }
    )[["Currency", "总欠款", "收入.1", "调整", "预扣税"]]

    # 调试展示
    st.subheader("📊 财报预览")
    st.write("识别列：", {
        "currency_source_col": currency_source_col,
        "revenue_col": revenue_col,
        "owed_col": owed_col,
        "adjust_col": adj_col if adj_col != "__adj__" else "(none→0)",
        "withholding_col": tax_col if tax_col != "__tax__" else "(none→0)",
    })
    st.dataframe(out.head())
    return out

def build_rates(df_report):
    """
    汇率：rate = 本币总额 / 美元收入（逐币种）
    分摊总额(USD)：((调整+预扣税)/rate) 合计
    """
    valid = df_report[(df_report["收入.1"].notna()) & (df_report["收入.1"] != 0)]
    if valid.empty:
        raise ValueError("❌ 财报 '收入.1' 全为 0/空，无法推导汇率")
    rates = dict(zip(valid["Currency"], (valid["总欠款"] / valid["收入.1"]).astype(float)))

    df = df_report.copy()
    df["rate"] = df["Currency"].map(rates)
    df["AdjTaxUSD"] = (df["调整"].fillna(0) + df["预扣税"].fillna(0)) / df["rate"]
    df["AdjTaxUSD"] = pd.to_numeric(df["AdjTaxUSD"], errors="coerce").fillna(0)

    total_adj_usd = float(df["AdjTaxUSD"].sum())
    report_total_usd = float(pd.to_numeric(df["收入.1"], errors="coerce").sum())
    return rates, total_adj_usd, report_total_usd

# ---------- 交易表（自动识别 + 手动映射兜底） ----------
def _norm(s: str) -> str:
    s = str(s)
    s = s.strip().lower()
    s = re.sub(r'[\s\-\_\/\.\(\):，,]+', '', s)
    return s

def _auto_guess_columns(cols):
    norm_map = {c: _norm(c) for c in cols}

    # 金额列（优先 Extended Partner Share / Proceeds / Amount）
    eps = None
    for c, n in norm_map.items():
        if ('extended' in n and 'partner' in n and ('share' in n or 'proceeds' in n or 'amount' in n)) \
           or ('partnershareextended' in n) \
           or (('partnershare' in n or 'partnerproceeds' in n) and ('amount' in n or 'gross' in n or 'net' in n)):
            eps = c; break
    if eps is None:
        for c, n in norm_map.items():
            if ('proceeds' in n or 'revenue' in n or 'amount' in n) and ('partner' in n or 'publisher' in n):
                eps = c; break

    # 币种列
    cur = None
    for c, n in norm_map.items():
        if ('currency' in n) or ('iso' in n and 'code' in n) or n == 'currency':
            cur = c; break
    if cur is None:
        for c, n in norm_map.items():
            if n.endswith('currencycode') or n.endswith('currency'):
                cur = c; break
    if cur is None and 'Currency' in cols:
        cur = 'Currency'

    # SKU 列
    sku = None
    for c, n in norm_map.items():
        if n == 'sku' or n.endswith('sku') or 'productid' in n or n == 'productid' or n == 'itemid':
            sku = c; break
    if sku is None and 'SKU' in cols:
        sku = 'SKU'

    return eps, cur, sku

def read_tx(file):
    df = read_any(file)
    df.columns = [str(c).strip() for c in df.columns]
    st.subheader("📊 交易表预览")
    st.write("列名：", list(df.columns))
    st.dataframe(df.head())

    eps, cur, sku = _auto_guess_columns(df.columns)
    with st.expander("🛠 手动列映射（自动识别不准时请调整）", expanded=not (eps and cur and sku)):
        cols = list(df.columns)
        eps = st.selectbox("金额列（Extended Partner Share / Proceeds / Amount）", cols, index=(cols.index(eps) if eps in cols else 0))
        cur = st.selectbox("币种列（Partner Share Currency / Currency）", cols, index=(cols.index(cur) if cur in cols else 0))
        sku = st.selectbox("SKU 列（SKU / Product ID / Item ID）", cols, index=(cols.index(sku) if sku in cols else 0))

    rename_map = {eps: "Extended Partner Share", cur: "Partner Share Currency", sku: "SKU"}
    df = df.rename(columns=rename_map)

    need = {"Extended Partner Share", "Partner Share Currency", "SKU"}
    missing = need - set(df.columns)
    if missing:
        raise ValueError(f"❌ 交易表仍缺少列：{missing}；请在“手动列映射”里选择正确的列。")

    # 金额转数值（去逗号）
    df["Extended Partner Share"] = df["Extended Partner Share"].astype(str).str.replace(",", "", regex=False)
    df["Extended Partner Share"] = pd.to_numeric(df["Extended Partner Share"], errors="coerce")
    return df

# ---------- 映射表 ----------
def read_map(file):
    df = pd.read_excel(file, engine="openpyxl", dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    st.subheader("📊 映射表预览")
    st.write("列名：", list(df.columns))
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
            rep = read_report(rp)
            rates, total_adj_usd, report_total_usd = build_rates(rep)

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

            mpdf = read_map(mp)
            sku2proj = dict(zip(mpdf["SKU"], mpdf["项目"]))
            txdf["项目"] = txdf["SKU"].map(sku2proj)

            txdf["Cost Allocation (USD)"] = txdf["Extended Partner Share USD"] / total_usd * total_adj_usd
            txdf["Net Partner Share (USD)"] = txdf["Extended Partner Share USD"] + txdf["Cost Allocation (USD)"]

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
