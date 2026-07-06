import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import shutil
import tempfile
import requests

# -----------------------------------------------------------------------------
# 1. PAGE CONFIGURATION & THEME STYLING
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Ras Elhekma S-19 - Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Corporate CSS for Premium Design & Arabic Support
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;600;700&display=swap');
    
    /* Global Typography */
    html, body, [class*="css"] {
        font-family: 'Cairo', sans-serif;
        text-align: right;
    }
    
    /* Title and Header Adjustments */
    .main-title {
        color: #0F172A;
        font-weight: 700;
        font-size: 2.5rem;
        margin-bottom: 0.2rem;
        text-align: center;
    }
    
    .sub-title {
        color: #475569;
        font-size: 1.1rem;
        text-align: center;
        margin-bottom: 2rem;
    }
    
    /* KPI Card Styling */
    .kpi-card {
        background-color: #FFFFFF;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
        border-right: 5px solid #3B82F6;
        text-align: center;
        transition: transform 0.2s;
    }
    .kpi-card:hover {
        transform: translateY(-4px);
    }
    .kpi-title {
        font-size: 0.95rem;
        color: #64748B;
        font-weight: 600;
        margin-bottom: 10px;
    }
    .kpi-value {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E293B;
    }
    
    /* Sidebar Sheet Tab Styling */
    .sheet-tab-active {
        background-color: #2563EB !important;
        color: white !important;
        font-weight: bold !important;
        border-radius: 8px;
    }
    
    /* Status Colors Definition */
    .status-a { border-right-color: #10B981 !important; }   /* Approved - Green */
    .status-awc { border-right-color: #F59E0B !important; } /* Approved with Comment - Orange */
    .status-r { border-right-color: #EF4444 !important; }   /* Rejected - Red */
    .status-p { border-right-color: #64748B !important; }   /* Pending - Gray */
    </style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 2. DEFINITIONS & PATHS
# -----------------------------------------------------------------------------
EXCEL_PATH = r"D:\log\RAAS ELHEKMA STATION LOG.xlsm"

# ضع هنا الـ ID بتاع ملف الجوجل درايف (من رابط المشاركة)
# مثال رابط: https://drive.google.com/file/d/1AbCdEfGhIjKlMnOpQrSt/view?usp=sharing
# الـ ID هو الجزء ده: 1AbCdEfGhIjKlMnOpQrSt
GDRIVE_FILE_ID = "1s4MPmPSLcxgvo75fccAQ7lwVthIEtU3c"
GDRIVE_DOWNLOAD_URL = f"https://drive.google.com/uc?export=download&id={GDRIVE_FILE_ID}"

SHEETS_LIST = [
    "Inspection Request",
    "Material Inspection Request",
    "Pour Concrete",
    "Test Result",
    "Release Finish",
    "Start New Activity",
    "DOCUMENT SUBMITTAL",
    "Material Submittal",
    "Request for Info.",
    "Request for Proposal",
    "Transmittals",
    "Shop Drawing",
    "As-Built",
    "NCR",
    "QVN",
    "SOR",
    "EMAILS"
]

STATUS_MAP = {
    'A': 'APPROVED',
    'AWC': 'APPROVED WITH COMMENT',
    'R': 'REJECTED',
    'P': 'PENDING'
}

STATUS_COLORS = {
    'APPROVED': '#10B981',
    'APPROVED WITH COMMENT': '#F59E0B',
    'REJECTED': '#EF4444',
    'PENDING': '#64748B'
}

# -----------------------------------------------------------------------------
# 3. ROBUST DATA LOADING FUNCTIONS
# -----------------------------------------------------------------------------
def download_from_gdrive(url):
    """
    يحمل ملف الإكسيل من رابط جوجل درايف المباشر ويحفظه في مجلد مؤقت.
    """
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        temp_dir = tempfile.gettempdir()
        temp_file_path = os.path.join(temp_dir, "online_dashboard_log.xlsm")
        with open(temp_file_path, "wb") as f:
            f.write(response.content)
        # تأكد إن الملف فعلاً إكسيل مش صفحة خطأ من جوجل
        if os.path.getsize(temp_file_path) < 5000:
            return None, "الملف المحمّل صغير جداً، تأكد إن رابط المشاركة صحيح ومتاح لأي حد."
        return temp_file_path, None
    except Exception as e:
        return None, str(e)


def get_safe_excel_reader(file_path):
    """
    Copies the excel file to a temp directory to prevent 'Permission Denied' 
    if the file is currently open in Excel by the user.
    """
    try:
        temp_dir = tempfile.gettempdir()
        temp_file_path = os.path.join(temp_dir, "temp_dashboard_log.xlsm")
        shutil.copy2(file_path, temp_file_path)
        return temp_file_path
    except Exception as e:
        st.error(f"حدث خطأ أثناء محاولة نسخ الملف المؤقت: {e}")
        return file_path

@st.cache_data(show_spinner=False)
def load_excel_sheet(file_source, sheet_name):
    """Loads a specific sheet. Source can be file path or uploaded file object."""
    try:
        # Load the sheet using openpyxl
        df = pd.read_excel(file_source, sheet_name=sheet_name, header=None, engine='openpyxl')
        return df, None
    except Exception as e:
        return None, str(e)

# Helper to generate Mock Data if Excel is not found (for demonstration/failsafe)
def generate_mock_data():
    import numpy as np
    categories = ['ARCH', 'MECH', 'SURVEY', 'ELEC', 'CIVIL']
    locations = ['Platform', 'Holding Tank', 'Water Tank', 'Main Building', 'Substation']
    subjects = [
        "Waterproofing test for roof slab",
        "Installation of fire alarm system first fix",
        "Excavation and leveling works",
        "Formwork and steel reinforcement check",
        "Cable tray installation and grounding"
    ]
    
    rows = []
    for i in range(3895, 3914):
        cat = np.random.choice(categories)
        loc = np.random.choice(locations)
        subj = np.random.choice(subjects) + f" (Item {i})"
        
        # Simulating revision codes
        rev00 = np.random.choice(['A', 'AWC', 'R', 'P', None], p=[0.3, 0.2, 0.1, 0.3, 0.1])
        rev01 = np.random.choice(['A', 'AWC', 'R', 'P', None], p=[0.4, 0.2, 0.05, 0.2, 0.15]) if rev00 in ['R', 'P', None] else None
        rev02 = np.random.choice(['A', 'AWC', 'R', 'P', None], p=[0.6, 0.1, 0.01, 0.1, 0.19]) if rev01 in ['R', 'P', None] else None
        rev03 = np.random.choice(['A', 'AWC', 'R', 'P', None], p=[0.8, 0.1, 0.0, 0.05, 0.05]) if rev02 in ['R', 'P', None] else None
        
        row = [None] * 24
        row[1] = f"RHS-KK-ECG-IR-{cat}-{i}-REV00" # Col B (IR No)
        row[2] = cat # Col C
        row[3] = loc # Col D
        row[4] = subj # Col E
        row[8] = rev00 # Col I
        row[12] = rev01 # Col M
        row[16] = rev02 # Col Q
        row[20] = rev03 # Col U
        rows.append(row)
        
    return pd.DataFrame(rows)

# -----------------------------------------------------------------------------
# 4. SIDEBAR NAVIGATION & FILE RETRIEVAL
# -----------------------------------------------------------------------------
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/000000/microsoft-excel.png", width=70)
    st.markdown("### 🗂️ متصفح صفحات اللوج")
    st.markdown("اضغط على اسم الشيت لتنتقل إليه مباشرة:")
    
    # Simple selection system mimicking excel tabs
    selected_sheet = st.radio(
        "الشيت الحالي:",
        SHEETS_LIST,
        index=0,
        label_visibility="collapsed"
    )
    
    st.markdown("---")
    st.markdown("⚙️ **إعدادات ومصدر البيانات**")
    
    data_source = None
    use_mock = False

    # 1) المحاولة الأولى: تحميل الملف أونلاين من جوجل درايف (يعمل مع أي حد فتح الرابط)
    with st.spinner("🔄 جاري جلب أحدث نسخة من ملف اللوج من جوجل درايف..."):
        gdrive_path, gdrive_err = download_from_gdrive(GDRIVE_DOWNLOAD_URL)

    if gdrive_path:
        st.success("✅ تم جلب أحدث نسخة من اللوج أونلاين بنجاح.")
        data_source = get_safe_excel_reader(gdrive_path) if False else gdrive_path
    else:
        st.warning(f"⚠️ لم يتم الوصول لملف جوجل درايف: {gdrive_err}")

        # 2) لو مش شغال، جرب المسار المحلي (للاستخدام على جهاز معين فقط)
        if os.path.exists(EXCEL_PATH):
            st.info("ℹ️ تم استخدام النسخة المحلية من الملف بدلاً من ذلك.")
            data_source = get_safe_excel_reader(EXCEL_PATH)
        else:
            # 3) وإلا اسمح برفع الملف يدوياً
            uploaded_file = st.file_uploader("قم برفع ملف الـ LOG يدوياً هنا:", type=["xlsm", "xlsx"])
            if uploaded_file is not None:
                data_source = uploaded_file
                st.success("✅ تم تحميل الملف المرفوع بنجاح!")
            else:
                st.info("💡 سيتم عرض بيانات افتراضية تفاعلية للمعاينة الآن.")
                use_mock = True

# -----------------------------------------------------------------------------
# 5. MAIN HEADER
# -----------------------------------------------------------------------------
st.markdown("<div class='main-title'>📊 لوحة متابعة مشروع محطة رأس الحكمة (S-19)</div>", unsafe_allow_html=True)
st.markdown(f"<div class='sub-title'>تحديث تفاعلي فوري لملف الحسابات | الشيت النشط الحالي: <b>{selected_sheet}</b></div>", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 6. SHEET ROUTING & LOGIC
# -----------------------------------------------------------------------------
if use_mock:
    raw_df = generate_mock_data()
    err = None
else:
    raw_df, err = load_excel_sheet(data_source, selected_sheet)

if err:
    st.error(f"حدث خطأ أثناء تحميل الشيت '{selected_sheet}': {err}")
    st.info("تأكد من أن اسم الشيت داخل ملف الإكسيل مطابق تماماً للاسم المكتوب.")
else:
    # -------------------------------------------------------------------------
    # MAIN SHEET LOGIC: INSPECTION REQUEST
    # -------------------------------------------------------------------------
    if selected_sheet == "Inspection Request":
        
        # Data Extraction (Rows B6 to B3913)
        # In python: index 5 is Row 6. Index 3912 is Row 3913.
        try:
            # We filter up to the length of df if it's smaller, safely
            end_row = min(3913, len(raw_df))
            df_subset = raw_df.iloc[5:end_row].copy()
            
            # Extract and map our key columns safely using positional indices
            # B=1, C=2, D=3, E=4, I=8, M=12, Q=16, U=20
            df_clean = pd.DataFrame()
            df_clean['IR_No'] = df_subset.iloc[:, 1]
            df_clean['Category'] = df_subset.iloc[:, 2]
            df_clean['Location'] = df_subset.iloc[:, 3]
            df_clean['Subject'] = df_subset.iloc[:, 4]
            df_clean['Rev00'] = df_subset.iloc[:, 8]
            df_clean['Rev01'] = df_subset.iloc[:, 12]
            df_clean['Rev02'] = df_subset.iloc[:, 16]
            df_clean['Rev03'] = df_subset.iloc[:, 20]
            
            # Filter out rows where IR No is empty
            df_clean = df_clean.dropna(subset=['IR_No'])
            df_clean['IR_No'] = df_clean['IR_No'].astype(str).str.strip()
            df_clean = df_clean[df_clean['IR_No'] != '']
            
            # Clean categories & locations
            df_clean['Category'] = df_clean['Category'].fillna('Other').astype(str).str.strip().str.upper()
            df_clean['Location'] = df_clean['Location'].fillna('NOT SPECIFIED').astype(str).str.strip()
            
            # ---------------------------------------------------------
            # RESOLVE STATUS TO LATEST REVISION (Logic backwards: Rev03 -> Rev00)
            # ---------------------------------------------------------
            def get_latest_rev_status(row):
                for rev_col in ['Rev03', 'Rev02', 'Rev01', 'Rev00']:
                    val = str(row[rev_col]).strip().upper() if pd.notna(row[rev_col]) else ""
                    if val in ['A', 'AWC', 'R', 'P']:
                        return STATUS_MAP[val]
                return 'PENDING' # Default fallback
            
            df_clean['Latest_Status'] = df_clean.apply(get_latest_rev_status, axis=1)
            
            # ---------------------------------------------------------
            # KPI METRICS SECTION
            # ---------------------------------------------------------
            total_irs = len(df_clean)
            status_counts = df_clean['Latest_Status'].value_counts()
            
            cnt_a = status_counts.get('APPROVED', 0)
            cnt_awc = status_counts.get('APPROVED WITH COMMENT', 0)
            cnt_r = status_counts.get('REJECTED', 0)
            cnt_p = status_counts.get('PENDING', 0)
            
            col_tot, col_a, col_awc, col_r, col_p = st.columns(5)
            
            with col_tot:
                st.markdown(f"""
                    <div class="kpi-card">
                        <div class="kpi-title">📋 إجمالي طلبات الفحص (IRs)</div>
                        <div class="kpi-value">{total_irs:,}</div>
                    </div>
                """, unsafe_allow_html=True)
                
            with col_a:
                st.markdown(f"""
                    <div class="kpi-card status-a">
                        <div class="kpi-title" style="color: #10B981;">✅ Approved (A)</div>
                        <div class="kpi-value">{cnt_a:,}</div>
                    </div>
                """, unsafe_allow_html=True)
                
            with col_awc:
                st.markdown(f"""
                    <div class="kpi-card status-awc">
                        <div class="kpi-title" style="color: #F59E0B;">⚠️ Approved with Comment (AWC)</div>
                        <div class="kpi-value">{cnt_awc:,}</div>
                    </div>
                """, unsafe_allow_html=True)
                
            with col_r:
                st.markdown(f"""
                    <div class="kpi-card status-r">
                        <div class="kpi-title" style="color: #EF4444;">❌ Rejected (R)</div>
                        <div class="kpi-value">{cnt_r:,}</div>
                    </div>
                """, unsafe_allow_html=True)
                
            with col_p:
                st.markdown(f"""
                    <div class="kpi-card status-p">
                        <div class="kpi-title" style="color: #64748B;">⏳ Pending (P)</div>
                        <div class="kpi-value">{cnt_p:,}</div>
                    </div>
                """, unsafe_allow_html=True)
                
            st.markdown("<br>", unsafe_allow_html=True)
            
            # ---------------------------------------------------------
            # INTERACTIVE FILTERS
            # ---------------------------------------------------------
            st.markdown("### 🔍 فلترة البيانات الذكية")
            filter_col1, filter_col2, filter_col3 = st.columns(3)
            
            with filter_col1:
                selected_cats = st.multiselect(
                    "تصفية بالقسم (Category):",
                    options=sorted(df_clean['Category'].unique()),
                    default=[]
                )
            with filter_col2:
                selected_locs = st.multiselect(
                    "تصفية بالموقع (Location):",
                    options=sorted(df_clean['Location'].unique()),
                    default=[]
                )
            with filter_col3:
                selected_statuses = st.multiselect(
                    "تصفية بالحالة النهائية (Status):",
                    options=list(STATUS_MAP.values()),
                    default=[]
                )
                
            # Apply filters dynamically
            filtered_df = df_clean.copy()
            if selected_cats:
                filtered_df = filtered_df[filtered_df['Category'].isin(selected_cats)]
            if selected_locs:
                filtered_df = filtered_df[filtered_df['Location'].isin(selected_locs)]
            if selected_statuses:
                filtered_df = filtered_df[filtered_df['Latest_Status'].isin(selected_statuses)]
                
            # ---------------------------------------------------------
            # CHARTS SECTION
            # ---------------------------------------------------------
            st.markdown("### 📊 التحليل البياني والإحصائيات")
            chart_col1, chart_col2 = st.columns([1, 1])
            
            with chart_col1:
                # 1. Doughnut Chart of Overall Status Distribution
                dist_data = filtered_df['Latest_Status'].value_counts().reset_index()
                dist_data.columns = ['Status', 'Count']
                
                fig_pie = px.pie(
                    dist_data,
                    names='Status',
                    values='Count',
                    hole=0.4,
                    title="📊 النسبة المئوية لتوزيع حالات طلبات الفحص الحالية",
                    color='Status',
                    color_discrete_map=STATUS_COLORS,
                )
                fig_pie.update_layout(legend_orientation="h")
                st.plotly_chart(fig_pie, use_container_width=True)
                
            with chart_col2:
                # 2. Stacked Bar Chart by Category
                cat_status = filtered_df.groupby(['Category', 'Latest_Status']).size().reset_index(name='Count')
                fig_bar = px.bar(
                    cat_status,
                    x='Category',
                    y='Count',
                    color='Latest_Status',
                    title="📈 مقارنة توزيع الحالات بين الأقسام المختلفة (ARCH, MECH, ELEC, etc.)",
                    color_discrete_map=STATUS_COLORS,
                    barmode='stack'
                )
                st.plotly_chart(fig_bar, use_container_width=True)
                
            # 3. Location Analysis Chart
            st.markdown("---")
            loc_data = filtered_df.groupby(['Location', 'Latest_Status']).size().reset_index(name='Count')
            # Sort by total IRs in that location
            loc_totals = filtered_df['Location'].value_counts().index
            fig_loc = px.bar(
                loc_data,
                y='Location',
                x='Count',
                color='Latest_Status',
                title="🗺️ توزيع طلبات الفحص حسب الموقع والمنطقة الجغرافية بالمشروع",
                color_discrete_map=STATUS_COLORS,
                orientation='h',
                category_orders={'Location': loc_totals}
            )
            st.plotly_chart(fig_loc, use_container_width=True)
            
            # ---------------------------------------------------------
            # INTERACTIVE DATA TABLE
            # ---------------------------------------------------------
            st.markdown("### 📋 جدول البيانات التفصيلي المتكامل")
            st.dataframe(
                filtered_df[['IR_No', 'Category', 'Location', 'Subject', 'Latest_Status', 'Rev00', 'Rev01', 'Rev02', 'Rev03']],
                use_container_width=True,
                height=400
            )
            
            # Export CSV Button
            csv = filtered_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📥 تحميل جدول التصفية الحالي كملف Excel (CSV)",
                data=csv,
                file_name='Filtered_Inspection_Request.csv',
                mime='text/csv',
            )

        except Exception as e:
            st.error(f"حدث خطأ غير متوقع أثناء معالجة البيانات: {e}")
            st.info("يرجى مراجعة تركيبة شيت 'Inspection Request' للتأكد من موافقتها للأعمدة المحددة.")

    # -------------------------------------------------------------------------
    # GENERIC PREVIEW FOR THE OTHER 16 SHEETS
    # -------------------------------------------------------------------------
    else:
        st.markdown(f"### 👀 استعراض بيانات شيت: {selected_sheet}")
        st.info("💡 تم تحميل هذا الشيت ديناميكياً. يمكنك تصفح واستيراد البيانات الكاملة له من الجدول أدناه:")
        
        # Clean up empty rows and columns for display
        df_display = raw_df.dropna(how='all').dropna(axis=1, how='all')
        
        # Display Search bar for this specific sheet
        search_query = st.text_input("🔍 ابحث عن أي قيمة داخل هذا الشيت:")
        if search_query:
            # Filter rows where any cell contains the query
            mask = df_display.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
            df_filtered = df_display[mask]
        else:
            df_filtered = df_display
            
        st.dataframe(df_filtered, use_container_width=True, height=500)
        
        # Simple download button for this sheet
        csv_sheet = df_filtered.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label=f"📥 تصدير بيانات {selected_sheet} الحالية",
            data=csv_sheet,
            file_name=f'{selected_sheet}_data.csv',
            mime='text/csv',
        )