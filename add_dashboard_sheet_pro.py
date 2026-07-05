import streamlit as st
import pandas as pd
import plotly.express as px
import os

# إعداد الصفحة
st.set_page_config(page_title="Ras Elhekma Dashboard", layout="wide")

st.title("📊 لوحة تحكم مشروع رأس الحكمة")

# 1. Sidebar لرفع الملف واختيار الشيت
with st.sidebar:
    st.markdown("### 🗂️ إعدادات البيانات")
    uploaded_file = st.file_uploader("ارفع ملف الإكسيل هنا:", type=["xlsm", "xlsx"])
    
    # تحديد مصدر البيانات
    data_source = None
    if uploaded_file is not None:
        data_source = uploaded_file
        st.success("✅ تم تحميل الملف بنجاح!")
    elif os.path.exists("RAAS ELHEKMA STATION LOG.xlsm"):
        data_source = "RAAS ELHEKMA STATION LOG.xlsm"
        st.info("📂 تم استخدام الملف الموجود في السيرفر")
    else:
        st.warning("⚠️ يرجى رفع ملف الإكسيل من هنا")

# 2. معالجة البيانات وعرضها
if data_source:
    try:
        # قراءة الملف
        xl = pd.ExcelFile(data_source)
        sheet_names = xl.sheet_names
        selected_sheet = st.sidebar.selectbox("اختار الشيت المطلوبة:", sheet_names)
        
        # قراءة البيانات (بافتراض إن الهيدر في الصف الثالث)
        df = pd.read_excel(data_source, sheet_name=selected_sheet, header=2)
        
        # عرض البيانات
        st.write(f"### 📋 عرض بيانات: {selected_sheet}")
        st.dataframe(df, use_container_width=True)
        
        # تحليل سريع (رسومات)
        st.subheader("📈 تحليل الأداء")
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        
        if numeric_cols:
            selected_col = st.selectbox("اختار عمود للتحليل:", numeric_cols)
            fig = px.bar(df, x=df.index, y=selected_col, title=f"رسم بياني لـ {selected_col}")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("لا توجد أعمدة رقمية للرسم البياني.")
            
    except Exception as e:
        st.error(f"حدث خطأ أثناء معالجة الملف: {e}")
else:
    st.write("---")
    st.markdown("### 💡 ملاحظة:")
    st.write("هذه الداش بورد مصممة لعرض بيانات مشروع رأس الحكمة. يرجى التأكد من رفع ملف الإكسيل الصحيح.")
