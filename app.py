import streamlit as st
import pandas as pd
import spacy
from spacy.cli import download
import io
from zipfile import ZipFile

# 嘗試載入 spaCy 官方日文模型，如果沒安裝就自動下載
@st.cache_resource
def load_ja_model():
    model_name = "ja_core_news_sm"
    try:
        return spacy.load(model_name)
    except OSError:
        download(model_name)
        return spacy.load(model_name)

nlp = load_ja_model()
st.sidebar.success(f"📦 使用 spaCy 日文 NER 模型：{nlp.meta['name']}")

# --- NER 抽取專有名詞 ---
def extract_entities(text):
    doc = nlp(text)
    return "、".join([ent.text for ent in doc.ents]) if doc.ents else ""

# --- UI ---
st.title("🔍 專有名詞（NER）自動抽取工具")
uploaded_files = st.file_uploader("請上傳 Excel 檔案（可多選）", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    colname = st.text_input("請輸入語句內容欄位名稱", value="語句內容")
    if st.button("開始抽取 NER"):
        with st.spinner("分析中..."):
            combined_df = pd.DataFrame()
            output_files = []
            all_dates = []

            def extract_date(filename):
                # 例如：2023Q1_XXX.xlsx 或 2024_YYY.xlsx 會取 2023Q1 或 2024
                import re
                m = re.match(r"^([0-9]{4}(Q[1-4])?)", filename)
                return m.group(1) if m else "未知"

            for file in uploaded_files:
                fname = file.name
                date_tag = extract_date(fname)
                all_dates.append(date_tag)
                df = pd.read_excel(file)
                if colname not in df.columns:
                    st.warning(f"{fname} 缺少 {colname} 欄位，已略過")
                    continue
                df = df.dropna(subset=[colname])
                sentences = df[colname].astype(str)
                df["NER專有名詞"] = sentences.apply(extract_entities)
                df.insert(0, "資料日期", date_tag)
                buffer = io.BytesIO()
                df.to_excel(buffer, index=False)
                output_files.append((f"{fname[:-5]}_ner.xlsx", buffer))
                combined_df = pd.concat([combined_df, df], ignore_index=True)

            # 合併輸出
            try:
                valid_dates = [d for d in all_dates if d != "未知"]
                if valid_dates:
                    min_date, max_date = min(valid_dates), max(valid_dates)
                else:
                    min_date = max_date = "未知"
            except:
                min_date = max_date = "未知"
            merged_filename = f"ner_總表合併_{min_date}-{max_date}.xlsx"
            merged_buffer = io.BytesIO()
            combined_df.to_excel(merged_buffer, index=False)
            merged_buffer.seek(0)
            output_files.append((merged_filename, merged_buffer))

            # 壓縮
            zip_buffer = io.BytesIO()
            with ZipFile(zip_buffer, "w") as zipf:
                for fname, buf in output_files:
                    buf.seek(0)
                    zipf.writestr(fname, buf.read())
            zip_buffer.seek(0)

        st.success("專有名詞（NER）抽取完成 ✅")
        st.download_button("📥 下載結果（ZIP壓縮包）", zip_buffer, file_name="ner_outputs.zip")
