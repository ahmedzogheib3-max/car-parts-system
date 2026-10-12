import streamlit as st
import pandas as pd
import sqlite3
import os
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

DB_NAME = "auto_parts_system.db"
EXCEL_BASE = "اجمالي البضاعة 03 - 09 - 2026.xlsx"

# تهيئة قاعدة البيانات واستيراد ملف الإكسيل المرفق كقاعدة أساسية
def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='products'")
    if not cursor.fetchone():
        if os.path.exists(EXCEL_BASE):
            try:
                xls = pd.ExcelFile(EXCEL_BASE)
                if 'Products' in xls.sheet_names:
                    df_prod = pd.read_excel(EXCEL_BASE, sheet_name='Products')
                    df_prod.columns = [str(c).strip() for c in df_prod.columns]
                    products_clean = pd.DataFrame()
                    products_clean['name'] = df_prod.get('الصنف', df_prod.iloc[:, 1])
                    products_clean['category'] = df_prod.get('ماركة السياراة', df_prod.iloc[:, 2])
                    products_clean['purchase_price'] = pd.to_numeric(df_prod.get('سعر التكلفه', df_prod.iloc[:, 3]), errors='coerce').fillna(0)
                    products_clean['wholesale_price'] = pd.to_numeric(df_prod.get('سعر بيع الجملة', df_prod.iloc[:, 4]), errors='coerce').fillna(0)
                    products_clean['sale_price'] = pd.to_numeric(df_prod.get('سعر البيع', df_prod.iloc[:, 5]), errors='coerce').fillna(0)
                    products_clean['stock_quantity'] = pd.to_numeric(df_prod.get('الكمية', df_prod.iloc[:, 6]), errors='coerce').fillna(0)
                    products_clean['min_stock_alert'] = 3
                    products_clean.to_sql('products', conn, if_exists='replace', index=False)
            except Exception as e:
                print("Error loading base excel:", e)
                
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS financial_ledger (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            transaction_type TEXT,
            amount REAL,
            category TEXT,
            description TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS sales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_no TEXT,
            product_name TEXT,
            quantity INTEGER,
            unit_price REAL,
            total_price REAL,
            customer_name TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

init_db()

st.set_page_config(page_title="نظام قطع غيار السيارات الذكي", page_icon="🚗", layout="wide")
st.title("🚗 نظام إدارة قطع الغيار بالذكاء الاصطناعي والتصدير المتقدم")

menu = st.sidebar.selectbox("📂 القائمة الرئيسية", ["🤖 المساعد الذكي", "📦 المخزون والبضاعة", "💰 الخزنة والماليات", "📊 التقارير والتصدير (Excel & PDF)"])

# 1. المساعد الذكي
if menu == "🤖 المساعد الذكي":
    st.subheader("💬 المساعد الذكي لتنفيذ العمليات")
    user_q = st.text_input("أدخل أمرك (مثال: تقرير المخزون، حالة الخزنة، تسعير...)")
    if user_q:
        if "مخزون" in user_q or "بضاعة" in user_q:
            conn = sqlite3.connect(DB_NAME)
            df = pd.read_sql("SELECT * FROM products", conn)
            conn.close()
            st.success(f"إجمالي عدد أصناف البضاعة في المخزون: {len(df)} صنف.")
            st.dataframe(df)
        elif "خزنة" in user_q or "مالية" in user_q:
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute("SELECT SUM(amount) FROM financial_ledger WHERE transaction_type='IN'")
            total_in = cursor.fetchone()[0] or 0.0
            conn.close()
            st.info(f"إجمالي إيرادات الخزنة الحالية: {total_in:,.2f} ج.م")
        else:
            st.write("أهلاً بك! يمكنك الاستعلام عن المخزون أو الخزنة، أو الانتقال لقسم التقارير لتصدير الملفات.")

# 2. إدارة المخزون وبضاعة الإكسيل
elif menu == "📦 المخزون والبضاعة":
    st.subheader("📦 بضاعة المحل الحالية (المستوردة من ملف الإكسيل)")
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql("SELECT * FROM products", conn)
    conn.close()
    st.dataframe(df, use_container_width=True)

# 3. الخزنة والماليات
elif menu == "💰 الخزنة والماليات":
    st.subheader("💰 الخزنة وحركة النقدية")
    conn = sqlite3.connect(DB_NAME)
    df_ledger = pd.read_sql("SELECT * FROM financial_ledger", conn)
    conn.close()
    if not df_ledger.empty:
        st.dataframe(df_ledger, use_container_width=True)
    else:
        st.warning("لا توجد حركات مالية مسجلة حتى الآن.")

# 4. التقارير والتصدير (Excel & PDF)
elif menu == "📊 التقارير والتصدير (Excel & PDF)":
    st.subheader("📊 تصدير التقارير بصيغ مختلفة (Excel / PDF)")
    
    conn = sqlite3.connect(DB_NAME)
    df_report = pd.read_sql("SELECT * FROM products", conn)
    conn.close()
    
    col1, col2 = st.columns(2)
    
    # تصدير إكسيل
    with col1:
        st.markdown("### 📥 تصدير كملف Excel")
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_report.to_excel(writer, index=False, sheet_name='Inventory')
        excel_data = output.getvalue()
        
        st.download_button(
            label="تحميل تقرير المخزون بصيغة Excel (.xlsx)",
            data=excel_data,
            file_name="Auto_Parts_Inventory_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # تصدير PDF
    with col2:
        st.markdown("### 📄 تصدير كملف PDF")
        if st.button("توليد وتحميل تقرير المخزون بصيغة PDF"):
            pdf_buffer = BytesIO()
            doc = SimpleDocTemplate(pdf_buffer, pagesize=letter)
            elements = []
            
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle('ArabicTitle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=16, alignment=1)
            
            elements.append(Paragraph("Auto Parts Inventory Report", title_style))
            elements.append(Spacer(1, 15))
            
            # جدول البيانات للـ PDF
            table_data = [['ID', 'Item Name', 'Car Brand', 'Qty', 'Sale Price']]
            for idx, row in df_report.head(30).iterrows():
                table_data.append([str(idx+1), str(row['name'])[:20], str(row['category']), str(row['stock_quantity']), str(row['sale_price'])])
                
            t = Table(table_data)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.grey),
                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0,0), (-1,0), 6),
                ('GRID', (0,0), (-1,-1), 0.5, colors.black)
            ]))
            elements.append(t)
            doc.build(elements)
            pdf_data = pdf_buffer.getvalue()
            
            st.download_button(
                label="اضغط هنا لتحميل ملف الـ PDF",
                data=pdf_data,
                file_name="Auto_Parts_Report.pdf",
                mime="application/pdf"
            )