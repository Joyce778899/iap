# app.py — IAP ORCAT Online（矩阵财报 | USD/本币 | 强校验 + 防串列）
import re
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="IAP — ORCAT Online (Matrix USD/Local, Safe)", page_icon="💼", layout="wide")
st.title("💼 IAP — ORCAT Online（矩阵财报 | USD/本币 | 强校验）")

with st.expander("使用说明", expanded=False):
    st.markdown("""
**上传 3 个文件：**
1) 交易表（CSV/XLSX）：需含 金额（本币）、币种（建议 3 位代码）、SKU  
2) Apple 财报（CSV/XLSX，矩阵格式）：**第三行**为表头，包含 `国家或地区 (货币) / 收入 / 总欠款 / 汇率 / 收入.1 / 预扣税 / 调整` 等  
3) 项目-SKU 映射（XLSX）：列 `项目`、`SKU`（SKU 可换行多个）

**核心逻辑（与你的模板匹配）**
- 从 `国家或地区 (货币)` 提取币种（三位代码）
- 币种聚合：`汇率(USD/本币) = ∑(收入.1, USD) / ∑(总欠款, 本币)`
- `(调整+预扣税)` 折美元：**乘法** `(调整+预扣税) * 汇率(USD/本币)`
- 交易 USD：`Extended Partner Share * 汇率(USD/本币)`（USD 自身=1）
- 分摊按交易 USD 占比；最终**净额合计 ≈ 财报美元收入合计**（可设容差）

**防呆/自检**
- 强制展开“手动列映射”
- 金额单位选择：元/分(÷100)/厘(÷1000)
- 金额分布体检（p90/p99/max；疑似ID列自动排除），异常直接阻断
- 币种值标准化（中文币名/括号代码 → 3 位代码），对不上直接阻断
""")

# ---------------- 基础读取 ----------------
def _read_any(uploaded, header=None):
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded, header=header, engine="python", on_bad_lines="skip")
    elif name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded, header=header, engine="openpyxl")
    else:
        raise ValueError("仅支持 CSV 或 Excel 文件")

def _norm_colkey(s: str) -> str:
    s = str(s).strip().lower()
    s = re.sub(r'[\s\-\_\/\.\(\):，,]+', '', s)
    return s

# ---------------- 财报解析（矩阵 → 币种聚合；USD/本币） ----------------
def read_report_matrix(uploaded) -> pd.DataFrame:
    # 你的财报为第三行(索引=2)是表头
    df = _read_any(uploaded, header=2)
    df = df[[c for c in df.columns if not str(c).startswith("Unnamed")]]
    df.columns = [str(c).strip() for c in df.columns]
    if "国家或地区 (货币)" not in df.columns:
        raise ValueError("财报缺少列：国家或地区 (货币)")

    # 提取三位币种代码
    df["Currency"] = df["国家或地区 (货币)"].astype(str).str.extract(r"\(([A-Za-z]{3})\)").iloc[:, 0]

    # 数值化（若列缺失则补 NaN）
    for c in ["收入", "收入.1", "调整", "预扣税", "总欠款", "汇率"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
        else:
            df[c] = np.nan

    # 币种聚合
    grp = df.dropna(subset=["Currency"]).groupby("Currency", dropna=False).agg(
        local_sum=("总欠款","sum"),
        usd_sum=("收入.1","sum"),
        adj_sum=("调整","sum"),
        wht_sum=("预扣税","sum"),
    ).reset_index()

    # 汇率(USD/本币)：∑USD / ∑本币
    grp["rate_usd_per_local"] = np.where(
        grp["local_sum"].abs() > 0,
        grp["usd_sum"].abs() / grp["local_sum"].abs(),
        np.nan
    )

    # (调整+预扣税) 折美元（乘法）
    grp["AdjTaxUSD"] = (grp["adj_sum"].fillna(0) + grp["wht_sum"].fillna(0)) * grp["rate_usd_per_local"]

    audit = grp.rename(columns={
        "local_sum": "本币总欠款",
        "usd_sum": "美元收入合计(收入.1)",
        "adj_sum": "调整(本币)合计",
        "wht_sum": "预扣税(本币)合计",
        "rate_usd_per_local": "汇率(USD/本币)"
    })
    return audit

def build_rates_and_totals(audit_df: pd.DataFrame):
    rates = dict(zip(audit_df["Currency"], audit_df["汇率(USD/本币)"]))
    report_total_usd = float(pd.to_numeric(audit_df["美元收入合计(收入.1)"], errors="coerce").sum())
    total_adj_usd = float(pd.to_numeric(audit_df["AdjTaxUSD"], errors="coerce").sum())
    return rates, total_adj_usd, report_total_usd

# ---------------- 交易表：自动识别 + 强制人工确认 + 自检 ----------------
# 中文币名 → 3 位代码（可按需扩充）
_CNY_MAP = {
    "人民币": "CNY","美元": "USD","欧元": "EUR","日元": "JPY","英镑": "GBP","港币": "HKD",
    "新台币": "TWD","台币":"TWD","韩元": "KRW","澳元": "AUD","加元":"CAD","新西兰元":"NZD",
    "卢布":"RUB","里拉":"TRY","兰特":"ZAR","瑞郎":"CHF","新加坡元":"SGD","沙特里亚尔":"SAR",
    "阿联酋迪拉姆":"AED","泰铢":"THB","新谢克尔":"ILS","匈牙利福林":"HUF","捷克克朗":"CZK",
    "丹麦克朗":"DKK","挪威克朗":"NOK","瑞典克朗":"SEK","波兰兹罗提":"PLN","罗马尼亚列伊":"RON",
    "墨西哥比索":"MXN","巴西雷亚尔":"BRL","智利比索":"CLP","新台幣":"TWD"
}

def _parse_numeric(s: pd.Series) -> pd.Series:
    t = s.astype(str).str.replace(",", "", regex=False).str.replace(r"[^\d\.\-\+]", "", regex=True)
    return pd.to_numeric(t, errors="coerce")

def _auto_guess_tx_cols_by_values(df: pd.DataFrame):
    cols = list(df.columns)
    norm_map = {c: _norm_colkey(c) for c in cols}

    # ===== 1) 金额列（防串列：自动排除“像ID”的长整型列） =====
    # 先按关键词命中
    amount = None
    for c, n in norm_map.items():
        if ('extended' in n and 'partner' in n and ('share' in n or 'proceeds' in n or 'amount' in n)) \
           or ('partnershareextended' in n) \
           or (('partnershare' in n or 'partnerproceeds' in n) and ('amount' in n or 'gross' in n or 'net' in n)) \
           or (('proceeds' in n or 'revenue' in n or 'amount' in n) and ('partner' in n or 'publisher' in n)):
            amount = c; break

    # 兜底：分布评分 + 排除“疑似ID”
    candidates = []
    for c in cols:
        v = _parse_numeric(df[c])
        if v.notna().mean() < 0.3:
            continue
        # 疑似ID：大多是整数 & p99 >= 1e9（10位级）
        ints_ratio = (v.dropna() == np.floor(v.dropna())).mean() if v.notna().any() else 0
        p99 = v.quantile(0.99) if v.notna().any() else 0
        if ints_ratio > 0.95 and p99 >= 1e9:
            continue  # 排除ID样式列

        # 评分：非空数、多样性、中位量级（过大降权）
        score = (
            v.notna().sum(),
            float(np.nanmedian(np.abs(v))) if v.notna().any() else 0.0,
            -float(np.nanquantile(np.abs(v), 0.99)) if v.notna().any() else 0.0
        )
        candidates.append((score, c))
    if amount is None:
        if candidates:
            candidates.sort(reverse=True)
            amount = candidates[0][1]
        else:
            # 万不得已：回退到“非空多&总和大”
            best = None; best_score = (-1, -1)
            for c in cols:
                v = _parse_numeric(df[c])
                score = (v.notna().sum(), v.abs().sum(skipna=True))
                if score > best_score:
                    best, best_score = c, score
            amount = best

    # ===== 2) 币种列 =====
    def ccy_score(series: pd.Series) -> float:
        s = series.dropna().astype(str).str.strip()
        token = s.str.extract(r"([A-Z]{3})")[0]
        rate = token.notna().mean() if len(s) else 0
        bonus = 0.15 if 'currency' in _norm_colkey(series.name) or 'isocode' in _norm_colkey(series.name) else 0.0
        return rate + bonus
    c_scores = {c: ccy_score(df[c]) for c in cols}
    currency = max(c_scores, key=c_scores.get)

    # ===== 3) SKU 列 =====
    sku = None
    for c, n in norm_map.items():
        if n == 'sku' or n.endswith('sku') or 'productid' in n or n == 'productid' or n == 'itemid':
            sku = c; break
    if sku is None and 'SKU' in cols:
        sku = 'SKU'
    if sku is None:
        text_scores = {}
        for c in cols:
            s = df[c].astype(str)
            v = _parse_numeric(s)
            nonnum_ratio = v.isna().mean()
            nunique = s.nunique(dropna=True)
            text_scores[c] = (nonnum_ratio, nunique)
        sku = max(text_scores, key=lambda c: (text_scores[c][0], text_scores[c][1]))

    return amount, currency, sku

def read_tx(uploaded, report_rates: dict):
    df = _read_any(uploaded)
    df.columns = [str(c).strip() for c in df.columns]
    st.subheader("📊 交易表预览")
    st.write("列名：", list(df.columns))
    st.dataframe(df.head())

    a, c, s = _auto_guess_tx_cols_by_values(df)
    cols = list(df.columns)
    with st.expander("🛠 手动列映射（请确认/修正）", expanded=True):
        a = st.selectbox("金额列（Extended Partner Share / Proceeds / Amount）", cols, index=(cols.index(a) if a in cols else 0))
        c = st.selectbox("币种列（3位代码或中文币名）", cols, index=(cols.index(c) if c in cols else 0))
        s = st.selectbox("SKU 列（SKU / Product ID / Item ID）", cols, index=(cols.index(s) if s in cols else 0))
        unit = st.radio("金额单位", ["单位元（不用换）", "单位分（÷100）", "单位厘（÷1000）"], index=0, horizontal=True)

    df = df.rename(columns={a:"Extended Partner Share", c:"Partner Share Currency", s:"SKU"})

    need = {"Extended Partner Share","Partner Share Currency","SKU"}
    missing = need - set(df.columns)
    if missing:
        st.error(f"系统猜测：金额={a} 币种={c} SKU={s}")
        raise ValueError(f"❌ 交易表缺列：{missing}")

    # 金额清洗 + 单位换算
    amt = _parse_numeric(df["Extended Partner Share"])
    if unit == "单位分（÷100）":
        amt = amt / 100.0
    elif unit == "单位厘（÷1000）":
        amt = amt / 1000.0
    df["Extended Partner Share"] = amt

    # 币种标准化（中文名/括号内代码 → 3位代码 → 大写）
    cval = df["Partner Share Currency"].astype(str).str.strip()
    code_from_paren = cval.str.extract(r"\(([A-Za-z]{3})\)", expand=False)
    final_ccy = cval.str.upper()
    final_ccy = np.where(code_from_paren.notna(), code_from_paren.str.upper(), final_ccy)
    final_ccy = pd.Series(final_ccy).replace(_CNY_MAP).str.upper()
    df["Partner Share Currency"] = final_ccy

    # —— 自检 1：金额分布（大额阻断）
    desc = amt.describe(percentiles=[0.5,0.9,0.99])
    p99, vmax = float(desc.get("99%", np.nan)), float(desc.get("max", np.nan))
    st.info(f"金额统计：min={desc.get('min',np.nan):.2f}, median={desc.get('50%',np.nan):.2f}, "
            f"p90={desc.get('90%',np.nan):.2f}, p99={p99:.2f}, max={vmax:.2f}")
    big_idx = np.argsort(-amt.fillna(0).to_numpy())[:20]
    st.caption("Top 20 大额样本（用于自检）")
    st.dataframe(df.iloc[big_idx][["Extended Partner Share","Partner Share Currency","SKU"]])
    if p99 > 1e6 or vmax > 1e8:
        st.error("⚠️ 金额分布异常大：可能金额列选错或金额单位不是“元”。请检查映射与“金额单位”。")
        st.stop()

    # —— 自检 2：币种集合对齐（缺失阻断）
    tx_ccy = set(df["Partner Share Currency"].dropna().unique().tolist())
    report_ccy = set(k for k,v in report_rates.items() if np.isfinite(v))
    st.write("交易表币种个数：", len(tx_ccy), "；财报可用币种个数：", len(report_ccy))
    st.write("交集样例：", sorted(list(tx_ccy & report_ccy))[:20])
    missing_in_report = sorted(tx_ccy - report_ccy)
    if missing_in_report:
        st.error(f"⚠️ 以下币种在财报中不存在或无法计算汇率：{missing_in_report}。请修正交易表币种或财报。")
        st.stop()

    return df

# ---------------- 映射表 ----------------
def read_map(uploaded):
    mp = _read_any(uploaded, header=0)
    mp.columns = [str(c).strip() for c in mp.columns]
    st.subheader("📊 映射表预览")
    st.write("列名：", list(mp.columns))
    st.dataframe(mp.head())
    if not {"项目","SKU"}.issubset(mp.columns):
        raise ValueError("❌ 映射表缺列：项目 或 SKU")
    mp = mp.assign(SKU=mp["SKU"].astype(str).str.split("\n")).explode("SKU")
    mp["SKU"] = mp["SKU"].str.strip()
    mp = mp[mp["SKU"] != ""]
    return mp[["项目","SKU"]]

# ---------------- 页面上传 ----------------
c1, c2, c3 = st.columns(3)
with c1: tx = st.file_uploader("① 交易表（CSV/XLSX）", type=["csv","xlsx","xls"], key="tx")
with c2: rp = st.file_uploader("② Apple 财报（矩阵，CSV/XLSX）", type=["csv","xlsx","xls"], key="rp")
with c3: mp = st.file_uploader("③ 项目-SKU（XLSX）", type=["xlsx","xls"], key="mp")

strict_check = st.checkbox("严格校验：净额总和≈财报美元收入（容差 $0.5 USD）", value=True)

if st.button("🚀 开始计算（USD/本币 | 强校验）"):
    if not (tx and rp and mp):
        st.error("❌ 请先上传三份文件")
    else:
        try:
            # 1) 财报 → 审计表（USD/本币）
            audit = read_report_matrix(rp)
            rates, total_adj_usd, report_total_usd = build_rates_and_totals(audit)

            # 2) 交易 + 自检
            txdf = read_tx(tx, rates)
            mpdf = read_map(mp)
            sku2proj = dict(zip(mpdf["SKU"], mpdf["项目"]))

            # 3) 交易换算美元（乘以 USD/本币）
            txdf["rate_usd_per_local"] = txdf["Partner Share Currency"].map(rates)
            txdf["Extended Partner Share USD"] = txdf["Extended Partner Share"] * txdf["rate_usd_per_local"]

            tx_total_usd = float(pd.to_numeric(txdf["Extended Partner Share USD"], errors="coerce").sum())
            if not np.isfinite(tx_total_usd) or tx_total_usd == 0:
                st.error("❌ 交易 USD 合计为 0：检查金额列/金额单位/币种映射")
                st.stop()

            # 4) 成本分摊（按交易 USD 占比）
            txdf["Cost Allocation (USD)"] = txdf["Extended Partner Share USD"] / tx_total_usd * total_adj_usd
            txdf["Net Partner Share (USD)"] = txdf["Extended Partner Share USD"] + txdf["Cost Allocation (USD)"]
            txdf["项目"] = txdf["SKU"].astype(str).map(sku2proj)

            # 5) 项目汇总
            summary = txdf.groupby("项目", dropna=False)[
                ["Extended Partner Share USD", "Cost Allocation (USD)", "Net Partner Share (USD)"]
            ].sum().reset_index()

            # 6) 校验：净额 ≈ 财报美元收入
            net_total = float(pd.to_numeric(txdf["Net Partner Share (USD)"], errors="coerce").sum())
            diff = net_total - report_total_usd
            if strict_check and (not np.isfinite(diff) or abs(diff) > 0.5):
                st.error(f"❌ 对账失败：交易净额 {net_total:,.2f} USD 与财报 {report_total_usd:,.2f} USD 差异 {diff:,.2f}。"
                         "请检查金额列/金额单位/币种。")
                st.stop()

            # 7) 总行与下载
            total_row = {
                "项目": "__TOTAL__",
                "Extended Partner Share USD": float(summary["Extended Partner Share USD"].sum()),
                "Cost Allocation (USD)": float(summary["Cost Allocation (USD)"].sum()),
                "Net Partner Share (USD)": float(summary["Net Partner Share (USD)"].sum()),
            }
            summary = pd.concat([summary, pd.DataFrame([total_row])], ignore_index=True)

            st.success("✅ 计算完成")
            st.markdown(f"- 财报美元收入合计（∑收入.1）：**{report_total_usd:,.2f} USD**")
            st.markdown(f"- 分摊总额（调整+预扣税 → USD）：**{total_adj_usd:,.2f} USD**")
            st.markdown(f"- 交易毛收入 USD 合计：**{tx_total_usd:,.2f} USD**")
            st.markdown(f"- 交易净额 USD 合计：**{net_total:,.2f} USD**，差异：**{diff:,.2f} USD**")

            st.download_button("⬇️ 审计：每币种汇率与分摊 (CSV)",
                               data=audit.to_csv(index=False).encode("utf-8-sig"),
                               file_name="financial_report_currency_rates.csv", mime="text/csv")
            st.download_button("⬇️ 逐单结果 (CSV)",
                               data=txdf.to_csv(index=False).encode("utf-8-sig"),
                               file_name="transactions_usd_net_project.csv", mime="text/csv")
            st.download_button("⬇️ 项目汇总 (CSV)",
                               data=summary.to_csv(index=False).encode("utf-8-sig"),
                               file_name="project_summary.csv", mime="text/csv")

            with st.expander("预览：财报审计（USD/本币）", expanded=False):
                st.dataframe(audit)
            with st.expander("预览：逐单结果", expanded=False):
                st.dataframe(txdf.head(200))
            with st.expander("预览：项目汇总", expanded=True):
                st.dataframe(summary)

        except Exception as e:
            st.error(f"⚠️ 出错：{e}")
            st.exception(e)
