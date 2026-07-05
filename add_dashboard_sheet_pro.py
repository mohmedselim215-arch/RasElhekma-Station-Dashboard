import streamlit as st
import pandas as pd

st.title("Ras Elhekma Dashboard")

file_path = "RAAS ELHEKMA STATION LOG.xlsm"

try:
    # هنعرض كل أسماء الشيتات عشان نعرف فين الـ 400 سطر
    xl = pd.ExcelFile(file_path)
    sheet_names = xl.sheet_names
    selected_sheet = st.selectbox("اختار الشيت اللي فيه البيانات:", sheet_names)
    
    # قراءة الشيت اللي هتختاره
    df = pd.read_excel(file_path, sheet_name=selected_sheet)
    
    st.write(f"بيانات الشيت: {selected_sheet}")
    st.dataframe(df)
    
except Exception as e:
    st.error(f"مشكلة: {e}")
