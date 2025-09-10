# app.py — IAP ORCAT Online（矩阵财报专用，含去重与向量化修复）
import re
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="IAP — ORCAT Online (Matrix Edition)", page_icon="💼", layout="wide")
st.title("💼 IAP — ORCAT Online（矩阵财报专用）")

with st.expander("使用说明", expanded=False):
    st.markdown("""
**请上传 3 个文件：**  
1) **交易表**（CSV/XLSX）：包含 金额（本币）、币种、SKU  
2) **Apple 财报（矩阵格式）**（CSV/XLSX）：第一列为“国家或地区 (货币)”，右侧为各币种列；下方多行是指标（总欠款/收入(美元)或收入.1/调整/预扣税/汇率）  
3) **项目-SKU 映射**（XLSX）：列 `项目`、`SKU`（SKU 可换行多个）

**计算概览：**  
- 从矩阵财报抽取每币种：`总欠款`、`收入.1(USD)`、`调整`、`预扣税`、`汇率(本币/美元)`；若缺“汇率”则用 `总欠款/收入.1` 推导  
- 分摊美元：`(调整+预扣税)/汇率` 的总额按交易 USD 占比分摊到每条记录  
- 结果输出：逐单（毛收入USD、分摊、净额、项目）与项目汇总，可下载
""")

# ------------------ 基础读取 ------------------
def _read_any(uploaded, header=None):
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded, header=header, engine="python", on_bad_lines="skip")
    elif name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded, header=header, engine="openpyxl")
    else:
        raise ValueError("仅支持 CSV 或 Excel (xlsx/xls) 文件")

def _norm_colkey(s: str) -> str:
    s = str(s).strip().lower()
    s = re.sub(r'[\s\-\_\/\.\(\):，,]+', '', s)
    return s

# ------------------ 矩阵财报解析 ------------------
def find_header_index(raw: pd.DataFrame) -> int:
    """定位‘国家或地区 (货币)’这一行的行号（无表头 DataFrame）。"""
    col0 = raw.iloc[:, 0].astype(str).str.replace("\u3000", " ").str.strip()
    idx = col0[col0 == "国家或地区 (货币)"].index.tolist()
    if idx:
        return idx[0]
    # 宽松匹配
    pattern = re.compile(r"国家或地区\s*[\(\（]货币[\)\）]")
    idx = col0[col0.str.contains(pattern)].index.tolist()
    if idx:
        return idx[0]
    # 再宽：包含 (XXX) 的行
    idx = col0[col0.str.contains(r"\([A-Za-z]{3}\)")].index.tolist()
    if idx:
        return idx[0]
    raise ValueError("未找到表头行：需要第一列为“国家或地区 (货币)”的一行。")

def _normalize_metric_name(s: str) -> str:
    s2 = re.sub(r"\s+", "", str(s))
    # 常见别名：收入(美元) → 收入.1
    s2 = s2.replace("收入（美元）", "收入.1").replace("收入(美元)", "收入.1")
    return s2

def parse_matrix_report(uploaded) -> pd.DataFrame:
    """
    将“国家或地区 (货币)”横向矩阵财报标准化为长表：
    返回列：Currency, 总欠款, 收入.1(USD), 汇率(本币/美元), 调整, 预扣税, AdjTaxUSD
    """
    raw = _read_any(uploaded, header=None)
    hdr = find_header_index(raw)

    headers = raw.iloc[hdr, :].tolist()                  # 第一行：国家(货币) + 各币种列名
    data_block = raw.iloc[hdr + 1 :, :].copy()
    data_block.columns = [f"col{i}" for i in range(data_block.shape[1])]
    metric_names = data_block["col0"].astype(str).str.strip()

    wanted = {"总欠款", "收入.1", "收入", "调整", "预扣税", "汇率"}

    # 币种列（跳过第一格标题）
    currencies_headers = []
    for h in headers[1:]:
        hs = str(h).strip()
        if hs and hs.lower() != "nan":
            currencies_headers.append(hs)

    # 逐币种抽取指标
    records = []
    for j, cur in enumerate(currencies_headers, start=1):
        colname = f"col{j}"
        values = {}
        for idx, s in enumerate(metric_names):
            nm = _normalize_metric_name(s)
            if nm in wanted:
                try:
                    values[nm] = pd.to_numeric(data_block.iloc[idx][colname], errors="coerce")
                except Exception:
                    values[nm] = pd.NA

        usd_rev = values.get("收入.1", pd.NA)
        if (pd.isna(usd_rev) or usd_rev is pd.NA) and ("收入" in values):
            usd_rev = values.get("收入", pd.NA)

        records.append({
            "CurrencyHeader": cur,
            "总欠款": values.get("总欠款", pd.NA),
            "收入.1": usd_rev,      # 视为美元收入
            "调整": values.get("调整", 0),
            "预扣税": values.get("预扣税", 0),
            "汇率": values.get("汇率", pd.NA),  # 若缺则后续用 总欠款/收入.1 推导
        })

    tidy = pd.DataFrame(records)

    # —— 修复点 1：去重列名，避免重复列名引发“多列赋单列”错误
    tidy = tidy.loc[:, ~tidy.columns.duplicated(keep="first")]

    # 提取 3 位币种代码
    tidy["Currency"] = tidy["CurrencyHeader"].astype(str).str.extract(r"\(([A-Za-z]{3})\)").iloc[:, 0]
    tidy = tidy.dropna(subset=["Currency"]).reset_index(drop=True)

    # 数值化
    for c in ["总欠款", "收入.1", "调整", "预扣税", "汇率"]:
        if c in tidy.columns:
            tidy[c] = pd.to_numeric(tidy[c], errors="coerce")

    # —— 修复点 2：向量化推导汇率，避免 apply 返回 DataFrame
    income = tidy["收入.1"].fillna(0).to_numpy(dtype="float64")
    base_local = tidy["总欠款"].fillna(0).to_numpy(dtype="float64")
    rate_calc = np.where(income != 0, base_local / income, np.nan)

    if "汇率" in tidy.columns:
        rate_given = tidy["汇率"].to_numpy(dtype="float64")
    else:
        rate_given = np.full(len(tidy), np.nan, dtype="float64")

    rate = np.where(np.isnan(rate_given), rate_calc, rate_given)
    tidy["rate"] = rate

    # —— 修复点 3：分摊美元（避免除 0）
    adj = tidy["调整"].fillna(0).to_numpy(dtype="float64")
    wht = tidy["预扣税"].fillna(0).to_numpy(dtype="float64")
    denom = rate.copy()
    denom[denom == 0] = np.nan
    adj_usd = (adj + wht) / denom
    tidy["AdjTaxUSD"] = pd.to_numeric(adj_usd, errors="coerce")

    # 输出列
    out = tidy[["Currency", "总欠款", "收入.1", "rate", "调整", "预扣税", "AdjTaxUSD"]].rename(
        columns={"收入.1": "收入.1(USD)", "rate": "汇率(本币/美元)"}
    )

    # 调试展示
    st.subheader("📊 财报（矩阵→标准化）预览")
    st.write("识别的币种列数：", len(out))
    st.dataframe(out.head(20))
    return out

def build_rates_and_totals(cleaned_report: pd.DataFrame):
    rates = dict(zip(cleaned_report["Currency"], cleaned_report["汇率(本币/美元)"]))
    total_adj_usd = float(pd.to_numeric(cleaned_report["AdjTaxUSD"], errors="coerce").sum())
    report_total_usd = float(pd.to_numeric(cleaned_report["收入.1(USD)"], errors="coerce").sum())
    return rates, total_adj_usd, report_total_usd

# ------------------ 交易表（自动识别 + 手动映射兜底） ------------------
def _auto_guess_tx_cols(cols):
    norm_map = {c: _norm_colkey(c) for c in cols}

    # 金额列（优先 extended partner share / proceeds / amount）
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

def read_tx(uploaded):
    df = _read_any(uploaded)
    df.columns = [str(c).strip() for c in df.columns]
    st.subheader("📊 交易表预览")
    st.write("列名：", list(df.columns))
    st.dataframe(df.head())

    eps, cur, sku = _auto_guess_tx_cols(df.columns)
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

    # 金额去逗号 → 数值
    df["Extended Partner Share"] = df["Extended Partner Share"].astype(str).str.replace(",", "", regex=False)
    df["Extended Partner Share"] = pd.to_numeric(df["Extended Partner Share"], errors="coerce")

    return df

# ------------------ 映射表 ------------------
def read_map(uploaded):
    df = pd.read_excel(uploaded, engine="openpyxl", dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    st.subheader("📊 映射表预览")
    st.write("列名：", list(df.columns))
    st.dataframe(df.head())
    if not {"项目", "SKU"}.issubset(df.columns):
        raise ValueError("❌ 映射表缺少列 `项目` 或 `SKU`")
    df = df.assign(SKU=df["SKU"].astype(str).str.split("\n")).explode("SKU")
    df["SKU"] = df["SKU"].str.strip()
    return df[df["SKU"] != ""][["项目", "SKU"]]

# ------------------ 上传区 ------------------
c1, c2, c3 = st.columns(3)
with c1: tx = st.file_uploader("① 交易表（CSV/XLSX）", type=["csv", "xlsx", "xls"], key="tx")
with c2: rp = st.file_uploader("② Apple 财报（矩阵格式，CSV/XLSX）", type=["csv", "xlsx", "xls"], key="rp")
with c3: mp = st.file_uploader("③ 项目-SKU（XLSX）", type=["xlsx", "xls"], key="mp")

if st.button("🚀 开始计算（矩阵财报专用）"):
    if not (tx and rp and mp):
        st.error("❌ 请先上传三份文件")
    else:
        try:
            # 1) 财报 → 标准化
            cleaned_report = parse_matrix_report(rp)
            rates, total_adj_usd, report_total_usd = build_rates_and_totals(cleaned_report)

            # 2) 交易 + 映射
            txdf = read_tx(tx)
            mpdf = read_map(mp)
            sku2proj = dict(zip(mpdf["SKU"], mpdf["项目"]))

            # 3) 交易换算 USD + 分摊
            txdf["Extended Partner Share USD"] = txdf.apply(
                lambda r: (r["Extended Partner Share"] / rates.get(str(r["Partner Share Currency"]), 1))
                          if pd.notnull(r["Extended Partner Share"]) else None,
                axis=1
            )
            total_usd = pd.to_numeric(txdf["Extended Partner Share USD"], errors="coerce").sum(min_count=1)
            if not pd.notnull(total_usd) or total_usd == 0:
                st.error("❌ 交易 USD 汇总为 0，可能币种不匹配或金额列为空")
                st.stop()

            txdf["Cost Allocation (USD)"] = txdf["Extended Partner Share USD"] / total_usd * total_adj_usd
            txdf["Net Partner Share (USD)"] = txdf["Extended Partner Share USD"] + txdf["Cost Allocation (USD)"]
            txdf["项目"] = txdf["SKU"].map(sku2proj)

            # 4) 汇总
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

            # 5) 展示 & 下载
            st.success("✅ 计算完成")
            st.markdown(f"- 财报美元收入总额（sum of 收入.1）：**{report_total_usd:,.2f} USD**")
            st.markdown(f"- 分摊总额（调整+预扣税）：**{total_adj_usd:,.2f} USD**")
            st.markdown(f"- 交易毛收入 USD 合计：**{float(total_usd):,.2f} USD**")

            st.download_button("⬇️ 下载：逐单结果 CSV",
                               data=txdf.to_csv(index=False).encode("utf-8-sig"),
                               file_name="transactions_usd_net_project.csv", mime="text/csv")
            st.download_button("⬇️ 下载：项目汇总 CSV",
                               data=summary.to_csv(index=False).encode("utf-8-sig"),
                               file_name="project_summary.csv", mime="text/csv")

            with st.expander("预览：财报（标准化后）", expanded=False):
                st.dataframe(cleaned_report)
            with st.expander("预览：逐单结果", expanded=False):
                st.dataframe(txdf.head(200))
            with st.expander("预览：项目汇总", expanded=True):
                st.dataframe(summary)

        except Exception as e:
            st.error(f"⚠️ 出错：{e}")
            st.exception(e)
