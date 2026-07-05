import pandas as pd
import streamlit as st

def add_dashboard_sheet_pro():
    file_path = r"D:\log\RAAS ELHEKMA STATION LOG.xlsm"
    
    if not os.path.exists(file_path):
        print("❌ الملف غير موجود!")
        return

    print("🚀 جاري فتح إكسيل لإضافة الشيت...")
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        wb = excel.Workbooks.Open(file_path)
        
        # التأكد من عدم وجود شيت بنفس الاسم عشان ميعملش خطأ
        try:
            wb.Sheets("Dashboard").Delete()
        except:
            pass
            
        # إنشاء شيت جديد ووضعه في البداية
        ws = wb.Sheets.Add(Before=wb.Sheets(1))
        ws.Name = "Dashboard"
        
        # إضافة زر تشغيل الداش بورد (رابط)
        ws.Range("B2").Value = "📊 لوحة التحكم التفاعلية"
        ws.Range("B2").Font.Size = 16
        ws.Range("B2").Font.Bold = True
        
        ws.Range("B4").Value = "اضغط هنا لتشغيل الداش بورد:"
        
        # إضافة رابط (Hyperlink)
        ws.Hyperlinks.Add(Anchor=ws.Range("B5"), Address="http://localhost:8501", TextToDisplay="🚀 تشغيل الداش بورد الآن")
        
        ws.Range("B5").Font.Size = 14
        ws.Range("B5").Font.Color = 0xFF0000 # لون أحمر
        
        wb.Save()
        wb.Close()
        print("✅ تم بنجاح إضافة شيت 'Dashboard' في ملف الإكسيل!")

    except Exception as e:
        print(f"❌ حدث خطأ: {e}")
    finally:
        excel.Quit()

if __name__ == "__main__":
    add_dashboard_sheet_pro()
