import streamlit as st
import pandas as pd

st.title("Ras Elhekma Dashboard")

file_path = "RAAS ELHEKMA STATION LOG.xlsm"

try:
    # هنقرأ البيانات ونقوله يتجاهل أول صفين (أو حسب مكان بياناتك)
    df = pd.read_excel(file_path, sheet_name=0, header=1) 
    
    st.write("بيانات المشروع:")
    st.dataframe(df) # ده هيعرض كل البيانات اللي في الشيت
    
except Exception as e:
    st.error(f"مشكلة في قراءة البيانات: {e}")
