import streamlit as st
import pandas as pd
import spacy
import subprocess
import io
from zipfile import ZipFile

# 嘗試載入 GiNZA 模型，若尚未安裝則自動下載
try:
    nlp = spacy.load("ja_ginza")
except OSError:
    subprocess.run(["python", "-m", "spacy", "download", "ja_ginza"], check=True)
    nlp = spacy.load("ja_ginza")

def extract_named_entities(text):
    doc = nlp(text)
    entities = [ent.text for ent in doc.ents if ent.label_ not in {"PUNCT", "SYM"}]
    return "、".join(entities) if entities else "（無）"

# Streamlit app
st.title("🧠 命名實體辨識（NER）專有名詞萃取工具")
uploaded_files = st.file_uploader("請上傳 Excel 檔案（可多選）", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    column_name = st.text_input("請輸入要分析的欄位名稱（預設為：語句內容）", value="語句內容")

    if st.button("開始分析"):
        with st.spinner("分析中..."):
            output_files = []
            combined_df = pd.DataFrame()

            for file in uploaded_files:
                filename = file.name
                df = pd.read_excel(file)

                if column_name not in df.columns:
                    st.warning(f"❌ 檔案 {filename} 中找不到欄位：{column_name}，已跳過")
                    continue

                df = df.dropna(subset=[column_name])
                df[column_name] = df[column_name].astype(str)

                df["命名實體"] = df[column_name].apply(extract_named_entities)
                combined_df = pd.concat([combined_df, df], ignore_index=True)

                buffer = io.BytesIO()
                df.to_excel(buffer, index=False)
                output_files.append((f"{filename[:-5]}_NER.xlsx", buffer))

            merged_buffer = io.BytesIO()
            with pd.ExcelWriter(merged_buffer, engine="openpyxl") as writer:
                combined_df.to_excel(writer, sheet_name="NER合併總表", index=False)
            merged_buffer.seek(0)
            output_files.append(("NER_總表合併.xlsx", merged_buffer))

            zip_buffer = io.BytesIO()
            with ZipFile(zip_buffer, "w") as zipf:
                for fname, buffer in output_files:
                    buffer.seek(0)
                    zipf.writestr(fname, buffer.read())
            zip_buffer.seek(0)

        st.success("分析完成 ✅")
        st.download_button("📥 下載所有檔案（ZIP 壓縮包）", zip_buffer, file_name="NER_outputs.zip")
