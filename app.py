import streamlit as st
import pandas as pd
import spacy
from spacy.cli import download
import io
from zipfile import ZipFile

# ─────────────────────────────────────────────
# 確保日文 GiNZA 模型可用；若尚未安裝則自動下載
try:
    nlp = spacy.load("ja_ginza")
except OSError:
    download("ja_ginza")       # 一次性下載
    nlp = spacy.load("ja_ginza")
# ─────────────────────────────────────────────

def extract_named_entities(text: str) -> str:
    """回傳句子中的命名實體，以 '、' 串接。無則回傳（無）"""
    doc = nlp(text)
    ents = [ent.text for ent in doc.ents if ent.label_ not in {"PUNCT", "SYM"}]
    return "、".join(ents) if ents else "（無）"

# ────────────── Streamlit UI ──────────────
st.title("🧠 日文 GiNZA – NER 專有名詞萃取工具")

uploaded_files = st.file_uploader(
    "請上傳 Excel 檔案（可多選）", type=["xlsx"], accept_multiple_files=True
)

if uploaded_files:
    column_name = st.text_input("請輸入要分析的欄位名稱（預設：語句內容）", value="語句內容")

    if st.button("開始分析"):
        with st.spinner("分析中…"):
            combined_df = pd.DataFrame()
            output_files = []

            for file in uploaded_files:
                df = pd.read_excel(file)
                if column_name not in df.columns:
                    st.warning(f"❌ {file.name} 缺少欄位『{column_name}』，已略過")
                    continue

                df = df.dropna(subset=[column_name]).copy()
                df["命名實體"] = df[column_name].astype(str).apply(extract_named_entities)

                # 個別檔案輸出
                buf = io.BytesIO()
                df.to_excel(buf, index=False)
                output_files.append((f"{file.name[:-5]}_NER.xlsx", buf))

                df["_來源檔"] = file.name
                combined_df = pd.concat([combined_df, df], ignore_index=True)

            # 合併總表
            merge_buf = io.BytesIO()
            with pd.ExcelWriter(merge_buf, engine="openpyxl") as writer:
                combined_df.to_excel(writer, sheet_name="NER合併總表", index=False)
            merge_buf.seek(0)
            output_files.append(("NER_總表合併.xlsx", merge_buf))

            # 壓縮成 ZIP
            zip_buf = io.BytesIO()
            with ZipFile(zip_buf, "w") as zf:
                for fname, buf in output_files:
                    buf.seek(0)
                    zf.writestr(fname, buf.read())
            zip_buf.seek(0)

        st.success("分析完成 ✅")
        st.download_button("📥 下載結果（ZIP）", zip_buf, file_name="NER_outputs.zip")
