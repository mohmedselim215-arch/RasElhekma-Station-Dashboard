import streamlit as st
import pandas as pd

st.title("Ras Elhekma Dashboard")

# قراءة أول شيت في ملف الإكسيل مباشرة
file_path = "RAAS ELHEKMA STATION LOG.xlsm"

try:
    # قراءة البيانات
    df = pd.read_excel(file_path, sheet_name=0) # sheet_name=0 يعني أول شيت
    
    st.write("بيانات المشروع:")
    st.dataframe(df) # ده هيعرض الـ 400 سطر قدامك بجدول تقدر تتحكم فيه
    
except Exception as e:
    st.error(f"مش عارف أفتح الملف: {e}")
