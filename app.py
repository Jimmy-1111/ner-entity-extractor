import streamlit as st
import pandas as pd
import io
from janome.tokenizer import Tokenizer

# ==== Stopwords 客製 ====
STOPWORDS = set([
    "こと", "ため", "もの", "よう", "ところ", "これ", "それ", "あれ",
    "の", "に", "は", "が", "を", "と", "で", "も", "へ", "から", "まで",
    "及び", "より", "として", "について", "など", "における", "により", "一方", "また", "さらに", "なお", "当面", "まず", "次に", "および"
])

# ==== 名詞抽取函數 ====
janome_tokenizer = Tokenizer()

def extract_nouns(text):
    tokens = janome_tokenizer.tokenize(text)
    nouns = [t.surface for t in tokens if t.part_of_speech.split(',')[0] == "名詞"]
    nouns = [w for w in nouns if w not in STOPWORDS and len(w) > 1]
    return nouns

# ==== Streamlit App ====
st.title("📋 日文每句全名詞抽取工具（已排除 stopwords）")
uploaded = st.file_uploader("請上傳 Excel（需有『語句內容』欄）", type=["xlsx"])

if uploaded:
    df = pd.read_excel(uploaded)
    if "語句內容" not in df.columns:
        st.error("缺少『語句內容』欄位")
        st.stop()
    sentences = df["語句內容"].astype(str).tolist()
    all_nouns = []
    for sent in sentences:
        nouns = extract_nouns(sent)
        all_nouns.append("、".join(sorted(set(nouns))) if nouns else "（無）")
    df["句內名詞"] = all_nouns
    st.dataframe(df)
    output = io.BytesIO()
    df.to_excel(output, index=False)
    st.download_button("📥 下載名詞結果", data=output.getvalue(), file_name="每句名詞結果.xlsx")
