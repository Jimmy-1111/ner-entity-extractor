import streamlit as st
import pandas as pd
import spacy
import subprocess, importlib
import io
from zipfile import ZipFile

# ---------- 1.  確保 GiNZA 模型可用 ----------
try:
    nlp = spacy.load("ja_ginza")
except OSError:
    # 在雲端環境靜默安裝 ja_ginza wheel
    subprocess.run(["pip", "install", "-q", "ja_ginza>=5.1.0"], check=True)
    importlib.invalidate_caches()
    nlp = spacy.load("ja_ginza")

# ---------- 2.  擷取命名實體 ----------
def extract_named_entities(text: str) -> str:
    doc = nlp(text)
    ents = [ent.text for ent in doc.ents if ent.label_ not in {"PUNCT", "SYM"}]
    return "、".join(ents) if ents else "（無）"

# ---------- 3.  Streamlit UI ----------
st.title("🧠 日文 GiNZA - NER 專有名詞萃取器")

uploaded_files = st.file_uploader(
    "請上傳 Excel 檔案（可多選）", type=["xlsx"], accept_multiple_files=True
)

if uploaded_files:
    column_name = st.text_input(
        "請輸入要分析的欄位名稱（預設：語句內容）", value="語句內容"
    )

    if st.button("開始分析"):
        with st.spinner("處理中…"):
            output_files, combined_df = [], pd.DataFrame()

            for f in uploaded_files:
                df = pd.read_excel(f)
                if column_name not in df.columns:
                    st.warning(f"❌ {f.name} 缺少欄位「{column_name}」，已略過")
                    continue

                df = df.dropna(subset=[column_name]).copy()
                df["命名實體"] = df[column_name].astype(str).apply(extract_named_entities)

                buf = io.BytesIO()
                df.to_excel(buf, index=False)
                output_files.append((f"{f.name[:-5]}_NER.xlsx", buf))

                df["_來源檔"] = f.name
                combined_df = pd.concat([combined_df, df], ignore_index=True)

            # 合併總表
            merge_buf = io.BytesIO()
            with pd.ExcelWriter(merge_buf, engine="openpyxl") as w:
                combined_df.to_excel(w, sheet_name="NER合併總表", index=False)
            merge_buf.seek(0)
            output_files.append(("NER_總表合併.xlsx", merge_buf))

            # 打包 ZIP
            zip_buf = io.BytesIO()
            with ZipFile(zip_buf, "w") as z:
                for fname, buf in output_files:
                    buf.seek(0)
                    z.writestr(fname, buf.read())
            zip_buf.seek(0)

        st.success("分析完成 ✅")
        st.download_button("📥 下載結果（ZIP）", zip_buf, file_name="NER_outputs.zip")

