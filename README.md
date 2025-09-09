# iap — ORCAT Online Pipeline

网页处理 Apple 报表与交易明细：逐单 USD、分摊（调整+预扣税）、净收入、项目映射与汇总。

## 本地运行
```bash
pip install -r requirements.txt
streamlit run app.py
```

访问 http://localhost:8501

## 部署到 Streamlit Community Cloud
1. 新建 GitHub 仓库（名称例如：**iap**），上传本项目文件：`app.py`、`requirements.txt`
2. 打开 https://share.streamlit.io/ → New app
3. 选择你的仓库、分支，App file 选 `app.py` → Deploy

## 部署到 Hugging Face Spaces
1. https://huggingface.co → Create new Space → SDK: **Streamlit**
2. 上传 `app.py` 与 `requirements.txt` → 保存后会自动构建并生成在线地址

## 输入文件要求（上传到网页端）
- 交易明细（CSV/XLSX）：列 `Extended Partner Share`、`Partner Share Currency`、`SKU`
- Apple 报表（CSV/XLSX）：列 `国家或地区 (货币)`、`总欠款`、`收入.1`（或等价列）、`调整`、`预扣税`
- 项目-SKU 映射（XLSX）：列 `项目`、`SKU`（SKU 支持换行多个）
