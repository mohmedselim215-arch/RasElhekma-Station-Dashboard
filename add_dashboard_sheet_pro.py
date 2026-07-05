import streamlit as st
import pandas as pd
import os

st.title("Ras Elhekma Dashboard Manager")

# إضافة زرار مباشر للتحديث
if st.button("تحديث الداش بورد"):
    st.write("تم الضغط على الزرار... جاري التنفيذ!")
    file_path = "RAAS ELHEKMA STATION LOG.xlsm"
    
    if os.path.exists(file_path):
        try:
            # محاولة الكتابة بـ openpyxl
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df = pd.DataFrame({'Status': ['Active'], 'Value': [100]})
                df.to_excel(writer, sheet_name='Dashboard', index=False)
            st.success("تم التحديث بنجاح!")
            st.write("تمت العملية.")
        except Exception as e:
            st.error(f"خطأ: {e}")
    else:
        st.error("الملف غير موجود!")
