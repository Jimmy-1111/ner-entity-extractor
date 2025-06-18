import streamlit as st
import pandas as pd
import spacy
import openpyxl
import io

# 直接 load（不用自動下載）
nlp = spacy.load("ja_core_news_sm")

st.title("日文 NER 專有名詞抽取工具")

uploaded_files = st.file_uploader("請上傳 Excel 檔案（可多選）", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    column_name = st.text_input("請輸入要分析的欄位名稱（預設：語句內容）", value="語句內容")

    if st.button("開始抽取專有名詞"):
        with st.spinner("分析中..."):
            output_files = []
            combined_df = pd.DataFrame()

            for file in uploaded_files:
                df = pd.read_excel(file)
                if column_name not in df.columns:
                    st.warning(f"{file.name} 缺少欄位：{column_name}，已略過")
                    continue

                # 專有名詞抽取
                ner_keywords = []
                for text in df[column_name].astype(str):
                    doc = nlp(text)
                    entities = [ent.text for ent in doc.ents if ent.label_ in ["PERSON", "ORG", "GPE", "LOC", "PRODUCT", "EVENT", "WORK_OF_ART"]]
                    ner_keywords.append("、".join(entities) if entities else "（無）")

                df["NER專有名詞"] = ner_keywords

                # 儲存檔案
                buffer = io.BytesIO()
                df.to_excel(buffer, index=False)
                output_files.append((file.name.replace('.xlsx', '_ner.xlsx'), buffer))

                # 合併用
                df["_來源檔名"] = file.name
                combined_df = pd.concat([combined_df, df], ignore_index=True)

            # 合併所有結果
            merged_buffer = io.BytesIO()
            combined_df.to_excel(merged_buffer, index=False)
            merged_buffer.seek(0)
            output_files.append(("ner_總表合併.xlsx", merged_buffer))

            # 打包ZIP
            from zipfile import ZipFile
            zip_buffer = io.BytesIO()
            with ZipFile(zip_buffer, "w") as zipf:
                for fname, buf in output_files:
                    buf.seek(0)
                    zipf.writestr(fname, buf.read())
            zip_buffer.seek(0)

        st.success("分析完成 ✅")
        st.download_button("📥 下載所有結果（ZIP 壓縮包）", zip_buffer, file_name="ner_outputs.zip")
