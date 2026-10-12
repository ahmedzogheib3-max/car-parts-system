import streamlit as st
import pandas as pd
import sqlite3
import os
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

DB_NAME = "auto_parts_shop.db"
EXCEL_FILE = "اجمالي البضاعة 03 - 09 - 2026.xlsx"

# 1. إعداد قاعدة البيانات من ملف الإكسيل
def init_database():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='products'")
    if not cursor.fetchone():
        if os.path.exists(EXCEL_FILE):
            try:
                df_prod = pd.read_excel(EXCEL_FILE, sheet_name='Products')
                df_prod.columns = [str(c).strip() for c in df_prod.columns]
                products_clean = pd.DataFrame()
                products_clean['name'] = df_prod.get('الصنف', df_prod.iloc[:, 1])
                products_clean['car_brand'] = df_prod.get('ماركة السياراة', df_prod.iloc[:, 2])
                products_clean['purchase_price'] = pd.to_numeric(df_prod.get('سعر التكلفه', df_prod.iloc[:, 3]), errors='coerce').fillna(0)
                products_clean['wholesale_price'] = pd.to_numeric(df_prod.get('سعر بيع الجملة', df_prod.iloc[:, 4]), errors='coerce').fillna(0)
                products_clean['sale_price'] = pd.to_numeric(df_prod.get('سعر البيع', df_prod.iloc[:, 5]), errors='coerce').fillna(0)
                products_clean['stock_quantity'] = pd.to_numeric(df_prod.get('الكمية', df_prod.iloc[:, 6]), errors='coerce').fillna(0)
                products_clean['min_stock_alert'] = 3
                products_clean.to_sql('products', conn, if_exists='replace', index=False)
            except Exception as e:
                print("Error loading products:", e)

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS sales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_no TEXT,
            product_name TEXT,
            quantity INTEGER,
            unit_price REAL,
            total_price REAL,
            customer_name TEXT,
            payment_type TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS expenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT,
            amount REAL,
            user_name TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

init_database()

# 2. واجهة المستخدم والتطبيق
st.set_page_config(page_title="برنامج قطع الغيار الذكي", page_icon="🚗", layout="wide")
st.title("🚗 نظام إدارة محلات قطع غيار السيارات الذكي")

menu = st.sidebar.selectbox("📂 القائمة الرئيسية", [
    "🤖 المحادثة مع المساعد الذكي", 
    "📦 المخزون وقطع الغيار", 
    "🛒 عمليات البيع", 
    "💰 الخزنة والمصروفات", 
    "📊 التقارير والتصدير (Excel & PDF)"
])

def get_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

if menu == "🤖 المحادثة مع المساعد الذكي":
    st.subheader("💬 المساعد الذكي التفاعلي")
    user_msg = st.text_input("اسأل المساعد عن المخزون أو الخزنة:")
    if user_msg:
        msg_lower = user_msg.lower()
        if "مخزون" in msg_lower or "بضاعة" in msg_lower:
            conn = get_connection()
            df = pd.read_sql("SELECT name, stock_quantity, sale_price FROM products", conn)
            conn.close()
            st.success("🤖 حالة المخزون الحالي:")
            st.dataframe(df)
        elif "خزنة" in msg_lower or "أرباح" in msg_lower:
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT SUM(total_price) FROM sales")
            total_sales = cursor.fetchone()[0] or 0.0
            cursor.execute("SELECT SUM(amount) FROM expenses")
            total_exp = cursor.fetchone()[0] or 0.0
            conn.close()
            st.success(f"🤖 إجمالي المبيعات: {total_sales:,.2f} ج.م | المصروفات: {total_exp:,.2f} ج.م | الصافي: {total_sales - total_exp:,.2f} ج.م")
        else:
            st.write("🤖 أهلاً بك! يمكنك سؤالي عن المخزون أو الخزنة.")

elif menu == "📦 المخزون وقطع الغيار":
    st.subheader("📦 رصيد المخزون والبضاعة الحالية")
    conn = get_connection()
    df = pd.read_sql("SELECT * FROM products", conn)
    conn.close()
    st.dataframe(df, use_container_width=True)

elif menu == "🛒 عمليات البيع":
    st.subheader("🛒 تسجيل فاتورة مبيعات جديدة")
    conn = get_connection()
    df_prod = pd.read_sql("SELECT name, stock_quantity, sale_price FROM products", conn)
    conn.close()
    
    with st.form("sale_form"):
        selected_product = st.selectbox("اختر القطعة", df_prod['name'].tolist() if not df_prod.empty else [])
        quantity = st.number_input("الكمية المباعة", min_value=1, value=1, step=1)
        customer_name = st.text_input("اسم العميل", value="عميل نقدي")
        payment_type = st.selectbox("طريقة الدفع", ["نقدي", "تحويل", "أجل"])
        
        submit_sale = st.form_submit_button("إتمام البيع وتحديث المخزون")
        if submit_sale:
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM products WHERE name = ?", (selected_product,))
            prod = cursor.fetchone()
            if prod and prod['stock_quantity'] >= quantity:
                total_price = prod['sale_price'] * quantity
                new_stock = prod['stock_quantity'] - quantity
                cursor.execute("UPDATE products SET stock_quantity = ? WHERE name = ?", (new_stock, selected_product))
                cursor.execute("INSERT INTO sales (product_name, quantity, unit_price, total_price, customer_name, payment_type) VALUES (?, ?, ?, ?, ?, ?)",
                               (selected_product, quantity, prod['sale_price'], total_price, customer_name, payment_type))
                conn.commit()
                conn.close()
                st.success(f"تمت عملية البيع بنجاح! الإجمالي: {total_price} ج.م")
                st.rerun()
            else:
                conn.close()
                st.error("الكمية غير متوفرة في المخزون أو المنتج غير موجود.")

elif menu == "💰 الخزنة والمصروفات":
    st.subheader("💰 إدارة المصروفات وحركة الخزنة")
    with st.form("expense_form"):
        title = st.text_input("بند المصروف (مثال: إيجار، مرتبات)")
        amount = st.number_input("المبلغ", min_value=0.0)
        sub = st.form_submit_button("تسجيل المصروف")
        if sub:
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("INSERT INTO expenses (title, amount, user_name) VALUES (?, ?, ?)", (title, amount, "Admin"))
            conn.commit()
            conn.close()
            st.success("تم تسجيل المصروف بنجاح.")
            st.rerun()

elif menu == "📊 التقارير والتصدير (Excel & PDF)":
    st.subheader("📊 تصدير التقارير بصيغ مختلفة")
    conn = get_connection()
    df_inv = pd.read_sql("SELECT * FROM products", conn)
    conn.close()
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### 📥 تصدير إلى Excel")
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_inv.to_excel(writer, index=False, sheet_name='Inventory')
        st.download_button("تحميل Excel (.xlsx)", output.getvalue(), "Inventory_Report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
    with col2:
        st.markdown("### 📄 تصدير إلى PDF")
        if st.button("تحميل تقرير PDF"):
            pdf_buffer = BytesIO()
            doc = SimpleDocTemplate(pdf_buffer, pagesize=letter)
            elements = []
            styles = getSampleStyleSheet()
            elements.append(Paragraph("Auto Parts Inventory Report", styles['Heading1']))
            elements.append(Spacer(1, 10))
            table_data = [['Item Name', 'Brand', 'Qty', 'Price']]
            for _, row in df_inv.head(30).iterrows():
                table_data.append([str(row['name'])[:20], str(row['car_brand']), str(row['stock_quantity']), str(row['sale_price'])])
            t = Table(table_data)
            t.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.grey), ('GRID', (0,0), (-1,-1), 0.5, colors.black)]))
            elements.append(t)
            doc.build(elements)
            st.download_button("اضغط لتحميل الـ PDF", pdf_buffer.getvalue(), "Inventory_Report.pdf", "application/pdf")
