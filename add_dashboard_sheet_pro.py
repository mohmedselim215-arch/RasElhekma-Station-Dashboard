import streamlit as st
import pandas as pd
import os

def add_dashboard_sheet_pro():
    file_path = "RAAS ELHEKMA STATION LOG.xlsm"
    
    if not os.path.exists(file_path):
        st.error("الملف غير موجود!")
        return

    try:
        # بنستخدم engine='openpyxl' وده الأفضل للتعامل مع ملفات xlsm
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df = pd.DataFrame({'Status': ['Active'], 'Value': [100]})
            df.to_excel(writer, sheet_name='Dashboard', index=False)
        st.success("تم التحديث بنجاح! 🚀")
    except Exception as e:
        st.error(f"حدث خطأ أثناء الكتابة: {e}")

st.title("Ras Elhekma Dashboard Manager")
if st.button("تحديث الداش بورد"):
    add_dashboard_sheet_pro()
