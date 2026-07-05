import streamlit as st
import pandas as pd
import os

def add_dashboard_sheet_pro():
    file_path = "RAAS ELHEKMA STATION LOG.xlsm"
    
    st.write("بيحاول يفتح الملف...") # دي هتظهرلك في الصفحة عشان نعرف وصل لحد فين
    
    if not os.path.exists(file_path):
        st.error("الملف مش موجود في السيرفر!")
        return

    try:
        st.write("جاري الكتابة في ملف الإكسيل...")
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df = pd.DataFrame({'Status': ['Active'], 'Value': [100]})
            df.to_excel(writer, sheet_name='Dashboard', index=False)
        st.success("تم التحديث!")
    except Exception as e:
        st.error(f"خطأ تقني: {e}")

st.title("Ras Elhekma Dashboard Manager")
if st.button("تحديث الداش بورد"):
    add_dashboard_sheet_pro()
