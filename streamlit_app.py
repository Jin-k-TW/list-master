import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="G-Change", layout="wide")

st.markdown("""
    <style>
    h1 { color: #800000; }
    </style>
""", unsafe_allow_html=True)

st.title("🚗 G-Change｜企業情報自動整形ツール（Ver3.1）")

uploaded_file = st.file_uploader("📤 編集前のExcelファイルをアップロード", type=["xlsx"])

# 抽出ルールに使うキーワード
review_keywords = ["楽しい", "親切", "人柄", "感じ", "スタッフ", "雰囲気", "交流", "お世話", "ありがとう", "です", "ました", "🙇"]
ignore_keywords = ["ウェブサイト", "ルート", "営業中", "閉店", "口コミ"]

def normalize(text):
    text = str(text).strip().replace(" ", " ").replace("　", " ")
    return re.sub(r'[−–—―]', '-', text)

def extract_info(lines):
    company = normalize(lines[0]) if lines else ""
    industry, address, phone = "", "", ""

    for line in lines[1:]:
        line = normalize(line)
        if any(kw in line for kw in ignore_keywords):
            continue
        if any(kw in line for kw in review_keywords):
            continue
        if "·" in line or "⋅" in line:
            parts = re.split(r"[·⋅]", line)
            industry = parts[-1].strip()
        elif re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line):
            phone = re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line).group()
        elif not address and any(x in line for x in ["丁目", "町", "番", "区", "−", "-"]):
            address = line

    return pd.Series([company, industry, address, phone])

def is_company_line(line):
    line = normalize(str(line))
    return not any(kw in line for kw in ignore_keywords + review_keywords) and not re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line)

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=None)
    lines = df[0].dropna().tolist()

    groups = []
    current = []
    for line in lines:
        line = normalize(str(line))
        if is_company_line(line):
            if current:
                groups.append(current)
            current = [line]
        else:
            current.append(line)
    if current:
        groups.append(current)

    result_df = pd.DataFrame([extract_info(group) for group in groups],
                             columns=["企業名", "業種", "住所", "電話番号"])

    st.success(f"✅ 整形完了：{len(result_df)}件の企業データを取得しました。")
    st.dataframe(result_df, use_container_width=True)

    # Excel保存
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        result_df.to_excel(writer, index=False, sheet_name="整形済みデータ")
    st.download_button("📥 Excelファイルをダウンロード", data=output.getvalue(),
                       file_name="整形済み_企業リスト.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")