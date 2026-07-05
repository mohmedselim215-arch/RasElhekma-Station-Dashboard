import streamlit as st
import pandas as pd
import plotly.express as px
import os

# إعداد الصفحة
st.set_page_config(page_title="Ras Elhekma Dashboard", layout="wide")

st.title("📊 لوحة تحكم مشروع رأس الحكمة")

# 1. اختيار الملف (سواء من السيرفر أو رفع يدوي)
uploaded_file = st.sidebar.file_uploader("ارفع ملف الإكسيل هنا:", type=["xlsm", "xlsx"])

if uploaded_file is not None:
    file_path = uploaded_file
else:
    # لو مفيش ملف مرفوع، ابحث عن الملف الموجود في نفس فولدر الكود
    file_path = "RAAS ELHEKMA STATION LOG.xlsm"

if os.path.exists(file_path) or uploaded_file:
    try:
        # قراءة الملف
        xl = pd.ExcelFile(file_path)
        sheet_names = xl.sheet_names
        selected_sheet = st.sidebar.selectbox("اختار الشيت:", sheet_names)
        
        # قراءة البيانات
        df = pd.read_excel(file_path, sheet_name=selected_sheet, header=2)
        
        st.write(f"### البيانات الحالية: {selected_sheet}")
        st.dataframe(df, use_container_width=True)
        
        # رسومات بيانية سريعة عشان الداش بورد تبان
        st.subheader("تحليل سريع")
        numeric_cols = df.select_dtypes(include=['number']).columns
        if len(numeric_cols) > 0:
            fig = px.bar(df, x=df.index, y=numeric_cols[0], title="رسم بياني لأول عمود أرقام")
            st.plotly_chart(fig, use_container_width=True)
            
    except Exception as e:
        st.error(f"حصلت مشكلة في قراءة البيانات: {e}")
else:
    st.warning("⚠️ يا هندسة، لازم ترفع ملف الإكسيل من القائمة اللي على الشمال عشان الداش بورد تشتغل!")
