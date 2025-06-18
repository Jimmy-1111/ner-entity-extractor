import streamlit as st
import pandas as pd
import spacy
from spacy.cli import download
import io
from zipfile import ZipFile

# ──────────────────── 1.  智慧載入日文 NER 模型 ────────────────────
MODEL_CANDIDATES = ["ja_ginza", "ja_ginza_electra", "ja_core_news_sm"]

def load_ja_model():
    for m in MODEL_CANDIDATES:
        try:
            return spacy.load(m)
        except OSError:
            try:
                download(m)        # 在線下載相容版本
                return spacy.load(m)
            except Exception:
                continue  # 換下一個候選
    raise RuntimeError("❌ 無法安裝任何日文 spaCy 模型。")

nlp = load_ja_model()
st.sidebar.success(f"📦 使用模型：{nlp.meta['name']}")

# ──────────────────── 2.  NER 萃取函式 ────────────────────────────
def extract_named_entities(text: str) -> str:
    doc = nlp(text)
    ents = [ent.text for ent in doc.ents if ent.label_ not in {"PUNCT", "SYM"}]
    return "、".join(ents) if ents else "（無）"

# ──────────────────── 3.  Streamlit UI ────────────────────────────
st.title("🧠 日文 NER 專有名詞萃取工具 (spaCy)")

uploaded_files = st.file_uploader(
    "請上傳 Excel 檔案（可多選）", type=["xlsx"], accept_multiple_files=True
)

if uploaded_files:
    column_name = st.text_input("請輸入要分析的欄位名稱（預設：語句內容）", value="語句內容")

    if st.button("開始分析"):
        with st.spinner("分析中…"):
            combined_df = pd.DataFrame()
            output_files = []

            for f in uploaded_files:
                df = pd.read_excel(f)
                if column_name not in df.columns:
                    st.warning(f"❌ {f.name} 缺少欄位『{column_name}』，已跳過")
                    continue

                df = df.dropna(subset=[column_name]).copy()
                df["命名實體"] = df[column_name].astype(str).apply(extract_named_entities)

                # 個別檔案
                buf = io.BytesIO()
                df.to_excel(buf, index=False)
                output_files.append((f"{f.name[:-5]}_NER.xlsx", buf))

                df["_來源檔"] = f.name
                combined_df = pd.concat([combined_df, df], ignore_index=True)

            # 合併總表
            merge_buf = io.BytesIO()
            with pd.ExcelWriter(merge_buf, engine="openpyxl") as writer:
                combined_df.to_excel(writer, sheet_name="NER合併總表", index=False)
            merge_buf.seek(0)
            output_files.append(("NER_總表合併.xlsx", merge_buf))

            # 打包 ZIP
            zip_buf = io.BytesIO()
            with ZipFile(zip_buf, "w") as zf:
                for fname, buf in output_files:
                    buf.seek(0)
                    zf.writestr(fname, buf.read())
            zip_buf.seek(0)

        st.success("分析完成 ✅")
        st.download_button("📥 下載結果（ZIP）", zip_buf, file_name="NER_outputs.zip")
