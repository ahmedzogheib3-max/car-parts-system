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

# 1. تهيئة قاعدة البيانات واستيراد الإكسيل الأصلي
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
        CREATE TABLE IF NOT EXISTS purchases (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supplier_name TEXT,
            product_name TEXT,
            quantity INTEGER,
            purchase_price REAL,
            total_price REAL,
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

st.set_page_config(page_title="برنامج قطع الغيار الذكي", page_icon="🚗", layout="wide")
st.title("🚗 نظام إدارة محلات قطع غيار السيارات الذكي والمتكامل")

def get_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

# القائمة الجانبية
menu = st.sidebar.selectbox("📂 القائمة الرئيسية", [
    "🤖 المساعد الذكي التفاعلي", 
    "📦 المخزون والبحث", 
    "🛒 عمليات البيع والشراء", 
    "💰 الخزنة والمصروفات", 
    "📊 التقارير والتصدير (Excel & PDF)"
])

# ==========================================
# 1. المساعد الذكي التفاعلي (محادثة وتنفيذ)
# ==========================================
if menu == "🤖 المساعد الذكي التفاعلي":
    st.subheader("💬 المساعد الذكي المتقدم")
    st.info("💡 يمكنك كتابة ما تريده بوضوح (مثال: 'تقرير المخزون'، 'حالة الخزنة'، أو البحث عن قطعة معينة).")

    if "chat_history" not in st.session_state:
        st.session_state.chat_history = []

    for message in st.session_state.chat_history:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

    if user_input := st.chat_input("اكتب أمرك للمساعد الذكي..."):
        st.session_state.chat_history.append({"role": "user", "content": user_input})
        with st.chat_message("user"):
            st.markdown(user_input)

        with st.chat_message("assistant"):
            query_lower = user_input.lower()
            response = ""
            
            if "مخزون" in query_lower or "بضاعة" in query_lower:
                conn = get_connection()
                df = pd.read_sql("SELECT name, car_brand, stock_quantity, sale_price FROM products", conn)
                conn.close()
                response = f"📊 **تقرير المخزون الحالي (إجمالي الأصناف: {len(df)}):**\n"
                for idx, row in df.iterrows():
                    response += f"- **{row['name']}** ({row['car_brand']}) | المتبقي: {row['stock_quantity']} | السعر: {row['sale_price']} ج.م\n"
            
            elif "خزنة" in query_lower or "مالية" in query_lower or "أرباح" in query_lower:
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("SELECT SUM(total_price) FROM sales")
                total_sales = cursor.fetchone()[0] or 0.0
                cursor.execute("SELECT SUM(amount) FROM expenses")
                total_exp = cursor.fetchone()[0] or 0.0
                conn.close()
                net = total_sales - total_exp
                response = f"💰 **الملخص المالي للخزنة:**\n- إجمالي المبيعات: {total_sales:,.2f} ج.م\n- إجمالي المصروفات: {total_exp:,.2f} ج.م\n- صافي الخزنة: {net:,.2f} ج.م"
            
            else:
                # بحث عام شامل في المنتجات
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("SELECT name, car_brand, stock_quantity, sale_price FROM products WHERE name LIKE ?", ('%' + user_input + '%',))
                results = cursor.fetchall()
                conn.close()
                
                if results:
                    response = f"🔍 **نتائج البحث عن '{user_input}':**\n"
                    for r in results:
                        response += f"- **{r['name']}** | الماركة: {r['car_brand']} | الكمية: {r['stock_quantity']} | السعر: {r['sale_price']} ج.م\n"
                else:
                    response = "عذراً، لم أتمكن من فهم الأمر بدقة أو العثور على العنصر المطابق. يمكنك استخدام الأقسام المخصصة للبيع والمخزون مباشرة."

            st.markdown(response)
            st.session_state.chat_history.append({"role": "assistant", "content": response})

# ==========================================
# 2. المخزون والبحث المفصل
# ==========================================
elif menu == "📦 المخزون والبحث":
    st.subheader("📦 إدارة المخزون والبحث الشامل")
    search_term = st.text_input("🔍 ابحث عن اسم القطعة أو ماركة السيارة:")
    
    conn = get_connection()
    if search_term:
        df = pd.read_sql(f"SELECT * FROM products WHERE name LIKE '%{search_term}%' OR car_brand LIKE '%{search_term}%'", conn)
    else:
        df = pd.read_sql("SELECT * FROM products", conn)
    conn.close()
    
    st.write(f"عدد النتائج المعروضة: **{len(df)}** صنف")
    st.dataframe(df, use_container_width=True)

# ==========================================
# 3. عمليات البيع والشراء الحقيقية
# ==========================================
elif menu == "🛒 عمليات البيع والشراء":
    tab1, tab2 = st.tabs(["🛒 تسجيل عملية بيع", "📦 تسجيل عملية شراء (توريد بضاعة)"])
    
    with tab1:
        st.subheader("تسجيل فاتورة مبيعات وتحديث المخزون والخزنة")
        conn = get_connection()
        df_prod = pd.read_sql("SELECT name, stock_quantity, sale_price FROM products", conn)
        conn.close()
        
        with st.form("sale_form"):
            prod_name = st.selectbox("اختر القطعة المباعة", df_prod['name'].tolist() if not df_prod.empty else [])
            qty = st.number_input("الكمية", min_value=1, value=1, step=1)
            cust_name = st.text_input("اسم العميل", value="عميل نقدي")
            pay_type = st.selectbox("طريقة الدفع", ["نقدي", "تحويل", "أجل"])
            
            submit_sale = st.form_submit_button("إتمام عملية البيع")
            if submit_sale:
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM products WHERE name = ?", (prod_name,))
                p = cursor.fetchone()
                if p and p['stock_quantity'] >= qty:
                    total = p['sale_price'] * qty
                    new_stock = p['stock_quantity'] - qty
                    invoice_no = f"INV-{os.urandom(3).hex().upper()}"
                    
                    # خصم المخزون
                    cursor.execute("UPDATE products SET stock_quantity = ? WHERE name = ?", (new_stock, prod_name))
                    # تسجيل المبيعات
                    cursor.execute("INSERT INTO sales (invoice_no, product_name, quantity, unit_price, total_price, customer_name, payment_type) VALUES (?, ?, ?, ?, ?, ?, ?)",
                                   (invoice_no, prod_name, qty, p['sale_price'], total, cust_name, pay_type))
                    conn.commit()
                    conn.close()
                    st.success(f"✅ تمت عملية البيع بنجاح! رقم الفاتورة: {invoice_no} | الإجمالي: {total:,.2f} ج.م (تم خصم الكمية من المخزون)")
                    st.rerun()
                else:
                    conn.close()
                    st.error("❌ عذراً، الكمية المطلوبة غير متوفرة في المخزون الحالي.")

    with tab2:
        st.subheader("تسجيل مشتريات وإضافة بضاعة للمخزون")
        conn = get_connection()
        df_prod = pd.read_sql("SELECT name FROM products", conn)
        conn.close()
        
        with st.form("purchase_form"):
            supplier = st.text_input("اسم المورد")
            p_name = st.text_input("اسم القطعة أو اختر من القائمة")
            p_qty = st.number_input("الكمية المشتراة", min_value=1, value=1, step=1)
            p_price = st.number_input("سعر شراء القطعة الواحدة", min_value=0.0)
            
            submit_purch = st.form_submit_button("إضافة بضاعة للمخزون")
            if submit_purch:
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM products WHERE name = ?", (p_name,))
                existing = cursor.fetchone()
                
                if existing:
                    new_qty = existing['stock_quantity'] + p_qty
                    cursor.execute("UPDATE products SET stock_quantity = ? WHERE name = ?", (new_qty, p_name))
                else:
                    cursor.execute("INSERT INTO products (name, car_brand, purchase_price, sale_price, stock_quantity) VALUES (?, ?, ?, ?, ?)",
                                   (p_name, supplier, p_price, p_price * 1.3, p_qty))
                
                cursor.execute("INSERT INTO purchases (supplier_name, product_name, quantity, purchase_price, total_price) VALUES (?, ?, ?, ?, ?)",
                               (supplier, p_name, p_qty, p_price, p_price * p_qty))
                conn.commit()
                conn.close()
                st.success(f"✅ تم إضافة {p_qty} قطعة من '{p_name}' للمخزون وتحديث الحسابات بنجاح!")
                st.rerun()

# ==========================================
# 4. الخزنة والمصروفات
# ==========================================
elif menu == "💰 الخزنة والمصروفات":
    st.subheader("💰 إدارة الخزنة والمصروفات اليومية")
    with st.form("exp_form"):
        title = st.text_input("بند المصروف (إيجار، مرتبات، نثريات...)")
        amount = st.number_input("المبلغ المدفوع", min_value=0.0)
        submit_exp = st.form_submit_button("حفظ المصروف وصرفه من الخزنة")
        if submit_exp:
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("INSERT INTO expenses (title, amount, user_name) VALUES (?, ?, ?)", (title, amount, "Admin"))
            conn.commit()
            conn.close()
            st.success("✅ تم تسجيل وصرف المصروف بنجاح.")
            st.rerun()
            
    st.markdown("---")
    st.write("📋 **سجل المصروفات المسجلة:**")
    conn = get_connection()
    df_exps = pd.read_sql("SELECT * FROM expenses", conn)
    conn.close()
    st.dataframe(df_exps, use_container_width=True)

# ==========================================
# 5. التقارير والتصدير (Excel & PDF)
# ==========================================
elif menu == "📊 التقارير والتصدير (Excel & PDF)":
    st.subheader("📊 استخراج التقارير وتصديرها")
    conn = get_connection()
    df_inv = pd.read_sql("SELECT * FROM products", conn)
    conn.close()
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### 📥 تصدير المخزون لملف Excel")
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_inv.to_excel(writer, index=False, sheet_name='Inventory')
        st.download_button("تحميل تقرير Excel (.xlsx)", output.getvalue(), "Inventory_Report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
    with col2:
        st.markdown("### 📄 تصدير تقرير المخزون PDF")
        if st.button("توليد وتحميل الـ PDF"):
            pdf_buffer = BytesIO()
            doc = SimpleDocTemplate(pdf_buffer, pagesize=letter)
            elements = []
            styles = getSampleStyleSheet()
            elements.append(Paragraph("Auto Parts Inventory Report", styles['Heading1']))
            elements.append(Spacer(1, 10))
            table_data = [['Item Name', 'Brand', 'Qty', 'Price']]
            for _, row in df_inv.head(50).iterrows():
                table_data.append([str(row['name'])[:20], str(row['car_brand']), str(row['stock_quantity']), str(row['sale_price'])])
            t = Table(table_data)
            t.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.grey), ('GRID', (0,0), (-1,-1), 0.5, colors.black)]))
            elements.append(t)
            doc.build(elements)
            st.download_button("اضغط لتحميل ملف الـ PDF", pdf_buffer.getvalue(), "Inventory_Report.pdf", "application/pdf")
