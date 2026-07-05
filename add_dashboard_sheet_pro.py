import streamlit as st
import pandas as pd
import os

st.title("Ras Elhekma Dashboard Manager")

if st.button("تحديث الداش بورد"):
    st.write("بيحاول يفتح الملف...")
    file_path = "RAAS ELHEKMA STATION LOG.xlsm"
    
    if not os.path.exists(file_path):
        st.error("الملف مش موجود في السيرفر!")
    else:
        try:
            st.write("جاري الكتابة في ملف الإكسيل...")
            writer = pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace')
            df = pd.DataFrame({'Status': ['Active'], 'Value': [100]})
            df.to_excel(writer, sheet_name='Dashboard', index=False)
            writer.close() 
            st.success("تم التحديث بنجاح! 🚀")
        except Exception as e:
            st.error(f"خطأ: {e}")
