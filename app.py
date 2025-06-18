import streamlit as st
import pandas as pd
from transformers import pipeline

@st.cache_resource
def load_ner_pipeline():
    return pipeline(
        "ner", 
        model="ku-nlp/roberta-base-japanese-ner", 
        aggregation_strategy="simple"
    )

def extract_entities(text, ner):
    if not isinstance(text, str) or not text.strip():
        return "（無）"
    try:
        entities = ner(text)
        if not entities:
            return "（無）"
        return "、".join(f"{e['word']}({e['entity_group']})" for e in entities)
    except Exception as e:
        return "（無）"

st.title("🚀 日文 BERT NER Streamlit 工具")

uploaded_file = st.file_uploader("請上傳要分析的 Excel 檔案（含句子內容）", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    # 讓使用者選擇語句內容欄位
    column_name = st.selectbox("選擇句子內容的欄位", options=df.columns, index=0)
    if st.button("開始辨識實體"):
        with st.spinner("載入模型與分析中，請稍候（首次下載會較久）..."):
            ner = load_ner_pipeline()
            df["NER專有名詞"] = df[column_name].apply(lambda x: extract_entities(x, ner))
        st.success("辨識完成！")
        st.write(df[["NER專有名詞"] + [column_name]].head(20))  # 預覽前20筆
        # 下載結果
        out = df.to_excel(index=False)
        st.download_button("📥 下載全部辨識結果 Excel", out, file_name="ner_結果.xlsx")

st.info("本工具使用 BERT NER 抽取日文專有名詞，推薦本地執行（首次執行需下載模型，請耐心等候）。")
