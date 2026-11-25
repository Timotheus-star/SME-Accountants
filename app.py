
import streamlit as st
import pandas as pd
import sqlite3
from datetime import date
from pathlib import Path
from io import BytesIO

from openpyxl.styles import Font, PatternFill, Alignment

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="SME Accountant", page_icon="📊", layout="wide")

BUSINESS_NAME = "Grace Provision Store"
OPENING_CAPITAL = 80000
DB_PATH = Path(__file__).parent / "sme_books.db"


# DATABASE
def get_connection():
    return sqlite3.connect(DB_PATH)


def initialize_database():
    with get_connection() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS sales (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sales_date TEXT NOT NULL,
                product TEXT NOT NULL,
                amount REAL NOT NULL,
                payment_method TEXT NOT NULL
            )
        """)

        conn.execute("""
            CREATE TABLE IF NOT EXISTS expenses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                expense_date TEXT NOT NULL,
                category TEXT NOT NULL,
                description TEXT NOT NULL,
                amount REAL NOT NULL,
                payment_method TEXT NOT NULL
            )
        """)

        conn.execute("""
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_name TEXT NOT NULL UNIQUE,
                quantity REAL NOT NULL DEFAULT 0,
                average_cost REAL NOT NULL DEFAULT 0,
                selling_price REAL NOT NULL DEFAULT 0,
                reorder_level REAL NOT NULL DEFAULT 5
            )
        """)

        conn.execute("""
            CREATE TABLE IF NOT EXISTS stock_purchases (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                purchase_date TEXT NOT NULL,
                product_id INTEGER NOT NULL,
                quantity REAL NOT NULL,
                unit_cost REAL NOT NULL,
                total_cost REAL NOT NULL,
                payment_method TEXT NOT NULL,
                FOREIGN KEY(product_id) REFERENCES products(id)
            )
        """)

        conn.execute("""
            CREATE TABLE IF NOT EXISTS liability_payments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                liability_id INTEGER NOT NULL,
                payment_date TEXT NOT NULL,
                amount REAL NOT NULL,
                payment_method TEXT NOT NULL,
                description TEXT,
                FOREIGN KEY(liability_id) REFERENCES liabilities(id)
            )
        """)



# SALES AND EXPENSE FUNCTIONS
def save_sale(sales_date, product, amount, payment_method):
    with get_connection() as conn:
        conn.execute("""
            INSERT INTO sales
            (sales_date, product, amount, payment_method)
            VALUES (?, ?, ?, ?)
        """, (str(sales_date), product, amount, payment_method))

def record_inventory_sale(sales_date, product_id, quantity, payment):
    if quantity <= 0:
        raise ValueError("Quantity must be greater than zero.")

    with get_connection() as conn:

        product = conn.execute("""
            SELECT product_name, quantity, average_cost, selling_price
            FROM products
            WHERE id = ?
        """, (product_id,)).fetchone()

        if product is None:
            raise ValueError("Product not found.")

        product_name, available_stock, average_cost, selling_price = product

        if quantity > available_stock:
            raise ValueError(
                f"Insufficient stock. Available stock: {available_stock}"
            )

        sales_amount = quantity * selling_price
        cost_of_goods_sold = quantity * average_cost
        gross_profit = sales_amount - cost_of_goods_sold

        conn.execute("""
            INSERT INTO sales
            (sales_date, product, amount, payment_method,
             product_id, quantity, unit_price,
             cost_of_goods_sold, gross_profit)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            str(sales_date), product_name, sales_amount, payment,
            product_id, quantity, selling_price,
            cost_of_goods_sold, gross_profit
        ))

        conn.execute("""
            UPDATE products
            SET quantity = quantity - ?
            WHERE id = ?
        """, (quantity, product_id))

        return sales_amount, cost_of_goods_sold, gross_profit


def save_expense(expense_date, category, description, amount, payment):
    with get_connection() as conn:
        conn.execute("""
            INSERT INTO expenses
            (expense_date, category, description, amount, payment_method)
            VALUES (?, ?, ?, ?, ?)
        """, (str(expense_date), category, description, amount, payment))


def get_sales():
    with get_connection() as conn:
        return pd.read_sql_query("SELECT * FROM sales ORDER BY id DESC", conn)


def get_expenses():
    with get_connection() as conn:
        return pd.read_sql_query("SELECT * FROM expenses ORDER BY id DESC", conn)

def get_stock_purchases():
    with get_connection() as conn:
        return pd.read_sql_query("""
            SELECT
                sp.id,
                sp.purchase_date,
                p.product_name,
                sp.quantity,
                sp.unit_cost,
                sp.total_cost,
                sp.payment_method
            FROM stock_purchases sp
            JOIN products p ON sp.product_id = p.id
            ORDER BY sp.id DESC
        """, conn)

def record_liability_payment(
    liability_id,
    payment_date,
    amount,
    payment_method,
    description
):
    with get_connection() as conn:
        liability = conn.execute("""
            SELECT amount, amount_paid
            FROM liabilities
            WHERE id = ?
        """, (liability_id,)).fetchone()

        if liability is None:
            raise ValueError("Liability record not found.")

        outstanding = liability[0] - liability[1]

        if amount <= 0:
            raise ValueError("Payment must be greater than zero.")

        if amount > outstanding:
            raise ValueError("Payment exceeds outstanding liability.")

        conn.execute("""
            INSERT INTO liability_payments (
                liability_id,
                payment_date,
                amount,
                payment_method,
                description
            )
            VALUES (?, ?, ?, ?, ?)
        """, (
            liability_id,
            payment_date.isoformat(),
            amount,
            payment_method,
            description
        ))

        conn.execute("""
            UPDATE liabilities
            SET amount_paid = amount_paid + ?
            WHERE id = ?
        """, (amount, liability_id))


# INVENTORY FUNCTIONS
def get_products():
    with get_connection() as conn:
        return pd.read_sql_query(
            "SELECT * FROM products ORDER BY product_name", conn
        )


def add_product(name, quantity, cost, selling, reorder):
    with get_connection() as conn:
        conn.execute("""
            INSERT INTO products
            (product_name, quantity, average_cost, selling_price, reorder_level)
            VALUES (?, ?, ?, ?, ?)
        """, (name, quantity, cost, selling, reorder))


def purchase_stock(product_id, purchase_date, quantity, unit_cost, payment):
    with get_connection() as conn:

        product = conn.execute("""
            SELECT quantity, average_cost
            FROM products WHERE id = ?
        """, (product_id,)).fetchone()

        if product is None:
            raise ValueError("Product not found.")

        old_quantity, old_cost = product

        new_quantity = old_quantity + quantity

        new_average_cost = (
            (old_quantity * old_cost) + (quantity * unit_cost)
        ) / new_quantity

        total_cost = quantity * unit_cost

        conn.execute("""
            UPDATE products
            SET quantity = ?, average_cost = ?
            WHERE id = ?
        """, (new_quantity, new_average_cost, product_id))

        conn.execute("""
            INSERT INTO stock_purchases
            (purchase_date, product_id, quantity, unit_cost,
             total_cost, payment_method)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            str(purchase_date), product_id, quantity,
            unit_cost, total_cost, payment
        ))
initialize_database()
# SALES DATABASE UPGRADE
with get_connection() as conn:
    existing_columns = {
        row[1] for row in conn.execute(
            "PRAGMA table_info(sales)"
        ).fetchall()
    }

    new_columns = [
        ("product_id", "INTEGER"),
        ("quantity", "REAL NOT NULL DEFAULT 0"),
        ("unit_price", "REAL NOT NULL DEFAULT 0"),
        ("cost_of_goods_sold", "REAL NOT NULL DEFAULT 0"),
        ("gross_profit", "REAL NOT NULL DEFAULT 0")
    ]

    for column_name, column_type in new_columns:
        if column_name not in existing_columns:
            conn.execute(
                f"ALTER TABLE sales ADD COLUMN "
                f"{column_name} {column_type}"
            )
# INITIALIZE

sales_df = get_sales()
expenses_df = get_expenses()
products_df = get_products()
purchases_df = get_stock_purchases()
conn.execute("""
    CREATE TABLE IF NOT EXISTS owner_drawings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        withdrawal_date TEXT NOT NULL,
        amount REAL NOT NULL,
        description TEXT,
        payment_method TEXT NOT NULL
    )
""")

conn.execute("""
          CREATE TABLE IF NOT EXISTS liabilities (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              liability_date TEXT NOT NULL,
              creditor TEXT NOT NULL,
              description TEXT NOT NULL,
              amount REAL NOT NULL,
              amount_paid REAL NOT NULL DEFAULT 0,
              due_date TEXT,
              liability_type TEXT NOT NULL
          )
      """)


total_sales = sales_df["amount"].sum() if not sales_df.empty else 0

total_expenses = (
    expenses_df["amount"].sum() if not expenses_df.empty else 0
)

total_gross_profit = (
    sales_df["gross_profit"].sum()
    if not sales_df.empty else 0
)

inventory_value = (
    (products_df["quantity"] * products_df["average_cost"]).sum()
    if not products_df.empty else 0
)

# DASHBOARD
st.title("📊 SME Accountant")
st.subheader(BUSINESS_NAME)
st.caption("Simple Accounting for Small Businesses")

c1, c2, c3, c4, c5 = st.columns(5)

c1.metric("Opening Capital", f"₦{OPENING_CAPITAL:,.2f}")
c2.metric("Total Sales", f"₦{total_sales:,.2f}")
c3.metric("Total Expenses", f"₦{total_expenses:,.2f}")
c4.metric("Inventory Value", f"₦{inventory_value:,.2f}")
c5.metric("Inventory Gross Profit", f"₦{total_gross_profit:,.2f}")



st.divider()



# MODULE TABS
sales_tab, expense_tab, inventory_tab, purchase_tab, drawings_tab, liability_tab, history_tab, reports_tab = st.tabs([
    "🛒 Sales",
    "💰 Expenses",
    "📦 Inventory",
    "🚚 Stock Purchases",
    "👤 Owner's Drawings",
    "📑 Liabilities",
    "📋 History",
    "📊 Financial Reports"
])



# SALES
with sales_tab:
    st.header("Record Daily Sales")

    sales_mode = st.radio(
        "Select Sales Type",
        ["Inventory Sale", "Manual Sale"],
        horizontal=True
    )

    if sales_mode == "Inventory Sale":

        products_df = get_products()

        if products_df.empty:
            st.warning("Register products in Inventory before recording sales.")

        else:
            product_options = {
                row["product_name"]: int(row["id"])
                for _, row in products_df.iterrows()
            }

            with st.form("inventory_sales_form"):

                sales_date = st.date_input(
                    "Sales Date", value=date.today()
                )

                selected_product = st.selectbox(
                    "Select Product",
                    list(product_options.keys())
                )

                quantity = st.number_input(
                    "Quantity Sold",
                    min_value=0.0,
                    step=1.0
                )

                payment = st.selectbox(
                    "Payment Method",
                    ["Cash", "Bank Transfer", "POS"],
                    key="inventory_sales_payment"
                )

                submitted = st.form_submit_button("Record Sale")

                if submitted:
                    try:
                        if quantity <= 0:
                            st.warning("Enter a valid quantity.")
                        else:
                            sales_amount, cost, profit = record_inventory_sale(
                                sales_date,
                                product_options[selected_product],
                                quantity,
                                payment
                            )

                            st.success("Inventory sale recorded successfully!")

                            st.write(f"Sales Amount: ₦{sales_amount:,.2f}")
                            st.write(f"Cost of Goods Sold: ₦{cost:,.2f}")
                            st.write(f"Gross Profit: ₦{profit:,.2f}")

                            st.rerun()

                    except ValueError as e:
                        st.error(str(e))

    else:
        with st.form("sales_form"):

            sales_date = st.date_input(
                "Sales Date", value=date.today()
            )

            product = st.text_input("Product / Description")

            amount = st.number_input(
                "Sales Amount (₦)",
                min_value=0.0,
                step=100.0
            )

            payment = st.selectbox(
                "Payment Method",
                ["Cash", "Bank Transfer", "POS"],
                key="sales_payment"
            )

            submitted = st.form_submit_button("Record Sale")

            if submitted:
                if product.strip() and amount > 0:
                    save_sale(
                        sales_date,
                        product.strip(),
                        amount,
                        payment
                    )

                    st.success("Sale saved successfully!")
                    st.rerun()

                else:
                    st.warning("Enter a product and valid amount.")


# EXPENSES
with expense_tab:
    st.header("Record Business Expenses")

    categories = [
        "Rent", "Electricity", "Transportation",
        "Salaries", "Maintenance", "Communication",
        "Packaging", "Other Expenses"
    ]

    with st.form("expense_form"):
        expense_date = st.date_input("Expense Date", value=date.today())
        category = st.selectbox("Expense Category", categories)
        description = st.text_input("Expense Description")
        amount = st.number_input("Expense Amount (₦)", min_value=0.0, step=100.0)
        payment = st.selectbox(
            "Payment Method",
            ["Cash", "Bank Transfer", "POS"],
            key="expense_payment"
        )

        submitted = st.form_submit_button("Save Expense")

        if submitted:
            if description.strip() and amount > 0:
                save_expense(expense_date, category, description, amount, payment)
                st.success("Expense saved successfully!")
                st.rerun()
            else:
                st.warning("Enter a description and valid amount.")


# INVENTORY
with inventory_tab:
    st.header("📦 Product Inventory")

    st.subheader("Register New Product")

    with st.form("product_form"):
        name = st.text_input("Product Name")

        quantity = st.number_input(
            "Opening Stock Quantity",
            min_value=0.0,
            step=1.0
        )

        cost = st.number_input(
            "Unit Cost Price (₦)",
            min_value=0.0,
            step=100.0
        )

        selling = st.number_input(
            "Unit Selling Price (₦)",
            min_value=0.0,
            step=100.0
        )

        reorder = st.number_input(
            "Low Stock Alert Level",
            min_value=0.0,
            value=5.0,
            step=1.0
        )

        submitted = st.form_submit_button("Register Product")

        if submitted:
            if name.strip() and quantity >= 0 and cost > 0 and selling > 0:
                try:
                    add_product(name.strip(), quantity, cost, selling, reorder)
                    st.success("Product registered successfully!")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("This product already exists.")
            else:
                st.warning("Enter a product name and valid prices.")

    st.divider()
    st.subheader("Available Stock")

    products_df = get_products()

    if not products_df.empty:
        display = products_df.copy()

        display["Stock Value (₦)"] = (
            display["quantity"] * display["average_cost"]
        )

        display["Stock Status"] = display.apply(
            lambda row: "LOW STOCK"
            if row["quantity"] <= row["reorder_level"]
            else "Available",
            axis=1
        )

        display = display.rename(columns={
            "product_name": "Product",
            "quantity": "Quantity",
            "average_cost": "Average Cost (₦)",
            "selling_price": "Selling Price (₦)",
            "reorder_level": "Reorder Level"
        })

        st.dataframe(
            display[[
                "Product", "Quantity", "Average Cost (₦)",
                "Selling Price (₦)", "Stock Value (₦)", "Stock Status"
            ]],
            use_container_width=True,
            hide_index=True
        )

    else:
        st.info("No products registered yet.")


# STOCK PURCHASES
with purchase_tab:
    st.header("🚚 Record New Stock Purchases")

    products_df = get_products()

    if products_df.empty:
        st.info("Register a product in Inventory before recording purchases.")

    else:
        product_options = {
            row["product_name"]: int(row["id"])
            for _, row in products_df.iterrows()
        }

        with st.form("purchase_form"):
            purchase_date = st.date_input(
                "Purchase Date", value=date.today()
            )

            selected_product = st.selectbox(
                "Select Product", list(product_options.keys())
            )

            quantity = st.number_input(
                "Quantity Purchased", min_value=0.0, step=1.0
            )

            unit_cost = st.number_input(
                "Purchase Cost per Unit (₦)",
                min_value=0.0,
                step=100.0
            )

            payment = st.selectbox(
                "Payment Method",
                ["Cash", "Bank Transfer", "POS"],
                key="purchase_payment"
            )

            submitted = st.form_submit_button("Save Stock Purchase")

            if submitted:
                if quantity > 0 and unit_cost > 0:
                    purchase_stock(
                        product_options[selected_product],
                        purchase_date,
                        quantity,
                        unit_cost,
                        payment
                    )
                    st.success("Stock purchase recorded!")
                    st.rerun()
                else:
                    st.warning("Enter a valid quantity and unit cost.")

# OWNER'S DRAWINGS
with drawings_tab:
    st.header("👤 Record Owner's Drawings")

    with st.form("drawings_form"):

        withdrawal_date = st.date_input(
            "Withdrawal Date",
            value=date.today()
        )

        amount = st.number_input(
            "Amount Withdrawn (₦)",
            min_value=0.0,
            step=500.0
        )

        description = st.text_input(
            "Description / Purpose",
            placeholder="e.g. Personal expenses"
        )

        payment_method = st.selectbox(
            "Payment Method",
            ["Cash", "Bank Transfer", "POS"],
            key="drawings_payment"
        )

        submitted = st.form_submit_button(
            "Save Owner's Withdrawal"
        )

        if submitted:
            if amount > 0:
                with get_connection() as conn:
                    conn.execute("""
                        INSERT INTO owner_drawings (
                            withdrawal_date,
                            amount,
                            description,
                            payment_method
                        )
                        VALUES (?, ?, ?, ?)
                    """, (
                        withdrawal_date.isoformat(),
                        amount,
                        description,
                        payment_method
                    ))

                st.success("Owner's withdrawal recorded!")
                st.rerun()

            else:
                st.warning("Enter a valid withdrawal amount.")

# LIABILITIES
with liability_tab:
    st.header("📑 Record Business Liability")

    with st.form("liability_form"):
        liability_date = st.date_input(
            "Liability Date",
            value=date.today()
        )

        creditor = st.text_input(
            "Creditor / Supplier Name"
        )

        description = st.text_input(
            "Description of Debt",
            placeholder="e.g. Goods purchased on credit"
        )

        amount = st.number_input(
            "Total Liability (₦)",
            min_value=0.0,
            step=1000.0
        )

        amount_paid = st.number_input(
            "Amount Already Paid (₦)",
            min_value=0.0,
            step=1000.0
        )

        due_date = st.date_input(
            "Expected Repayment Date",
            value=None
        )

        liability_type = st.selectbox(
            "Liability Classification",
            ["Current Liability", "Non-current Liability"]
        )

        submitted = st.form_submit_button(
            "Save Liability"
        )

        if submitted:
            if creditor.strip() and amount > 0 and amount_paid <= amount:
                with get_connection() as conn:
                    conn.execute("""
                        INSERT INTO liabilities (
                            liability_date,
                            creditor,
                            description,
                            amount,
                            amount_paid,
                            due_date,
                            liability_type
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, (
                        liability_date.isoformat(),
                        creditor.strip(),
                        description.strip(),
                        amount,
                        amount_paid,
                        due_date.isoformat() if due_date else None,
                        liability_type
                    ))

                st.success("Liability recorded successfully!")
                st.rerun()

            else:
                st.warning(
                    "Enter a creditor, a valid liability amount, "
                    "and ensure amount paid does not exceed the total liability."
                )


# TRANSACTION HISTORY
with history_tab:
    st.header("Sales History")

    sales_df = get_sales()

    if not sales_df.empty:
        st.dataframe(sales_df, use_container_width=True, hide_index=True)
    else:
        st.info("No sales recorded.")

    st.divider()
    st.header("Expense History")

    expenses_df = get_expenses()

    if not expenses_df.empty:
        st.dataframe(expenses_df, use_container_width=True, hide_index=True)
    else:
        st.info("No expenses recorded.")

    st.divider()
    st.header("Stock Purchase History")

    with get_connection() as conn:
        purchases_df = pd.read_sql_query("""
            SELECT
                sp.id,
                sp.purchase_date,
                p.product_name,
                sp.quantity,
                sp.unit_cost,
                sp.total_cost,
                sp.payment_method
            FROM stock_purchases sp
            JOIN products p ON sp.product_id = p.id
            ORDER BY sp.id DESC
        """, conn)

    if not purchases_df.empty:
        st.dataframe(
            purchases_df,
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("No stock purchases recorded.")

st.divider()


# FINANCIAL REPORTS
with reports_tab:
    st.header("📊 Profit & Loss Statement")
    st.caption("Grace Provision Store")

    # REPORTING PERIOD
    st.subheader("Select Reporting Period")

    start_date, end_date = st.date_input(
        "Reporting Date Range",
        value=(
            date.today().replace(day=1),
            date.today()
        )
    )

    if start_date > end_date:
        st.error("Start date cannot be later than end date.")

    else:
        # LOAD TRANSACTIONS
        sales_df = get_sales()
        expenses_df = get_expenses()

        # FILTER SALES BY DATE
        if not sales_df.empty:
            sales_df["sales_date"] = pd.to_datetime(
                sales_df["sales_date"]
            ).dt.date

            sales_df = sales_df[
                (sales_df["sales_date"] >= start_date) &
                (sales_df["sales_date"] <= end_date)
            ]

        # FILTER EXPENSES BY DATE
        if not expenses_df.empty:
            expenses_df["expense_date"] = pd.to_datetime(
                expenses_df["expense_date"]
            ).dt.date

            expenses_df = expenses_df[
                (expenses_df["expense_date"] >= start_date) &
                (expenses_df["expense_date"] <= end_date)
            ]

        
        # FILTER STOCK PURCHASES BY DATE
        if not purchases_df.empty:
            purchases_df["purchase_date"] = pd.to_datetime(
                purchases_df["purchase_date"]
            ).dt.date

            purchases_df = purchases_df[
                (purchases_df["purchase_date"] >= start_date) &
                (purchases_df["purchase_date"] <= end_date)
            ]
    

        # CALCULATE REVENUE
        total_revenue = (
            sales_df["amount"].sum()
            if not sales_df.empty else 0
        )
        # CALCULATE CASH INFLOW
        cash_inflow = (
            sales_df["amount"].sum()
            if not sales_df.empty else 0
        )

        # CALCULATE CASH OUTFLOW
        cash_expenses = (
            expenses_df["amount"].sum()
            if not expenses_df.empty else 0
        )

        cash_stock_purchases = (
            purchases_df["total_cost"].sum()
            if not purchases_df.empty else 0
        )

        cash_outflow = cash_expenses + cash_stock_purchases
        # CALCULATE NET CASH FLOW
        net_cash_flow = cash_inflow - cash_outflow

        st.divider()

        # CASH FLOW ANALYSIS
        st.subheader("💰 Cash Flow Analysis")

        cf1, cf2, cf3 = st.columns(3)

        cf1.metric(
            "Cash Inflow",
            f"₦{cash_inflow:,.2f}"
        )

        cf2.metric(
            "Cash Outflow",
            f"₦{cash_outflow:,.2f}"
        )

        cf3.metric(
            "Net Cash Flow",
            f"₦{net_cash_flow:,.2f}"
        )



        # COST OF GOODS SOLD
        total_cogs = (
            sales_df["cost_of_goods_sold"].sum()
            if not sales_df.empty else 0
        )

        # GROSS PROFIT
        gross_profit = total_revenue - total_cogs

        # OPERATING EXPENSES
        operating_expenses = (
            expenses_df["amount"].sum()
            if not expenses_df.empty else 0
        )
        # NET PROFIT
        net_profit = gross_profit - operating_expenses

        # PROFIT MARGINS
        gross_profit_margin = (
            (gross_profit / total_revenue) * 100
            if total_revenue > 0 else 0
        )

        net_profit_margin = (
            (net_profit / total_revenue) * 100
            if total_revenue > 0 else 0
        )
        # MONTHLY PROFITABILITY TREND
        st.subheader("Monthly Profitability Trend")

        monthly_sales = sales_df.copy()
        monthly_expenses = expenses_df.copy()

        if not monthly_sales.empty:
            monthly_sales["sales_date"] = pd.to_datetime(
                monthly_sales["sales_date"]
            )

            monthly_sales["month"] = (
                monthly_sales["sales_date"].dt.to_period("M").astype(str)
            )

            monthly_summary = (
                monthly_sales.groupby("month")
                .agg(
                    Revenue=("amount", "sum"),
                    COGS=("cost_of_goods_sold", "sum")
                )
                .reset_index()
            )
        else:
            monthly_summary = pd.DataFrame(
                columns=["month", "Revenue", "COGS"]
            )

        if not monthly_expenses.empty:
            monthly_expenses["expense_date"] = pd.to_datetime(
                monthly_expenses["expense_date"]
            )

            monthly_expenses["month"] = (
                monthly_expenses["expense_date"].dt.to_period("M").astype(str)
            )

            monthly_expense_summary = (
                monthly_expenses.groupby("month")["amount"]
                .sum()
                .reset_index(name="Operating Expenses")
            )
        else:
            monthly_expense_summary = pd.DataFrame(
                columns=["month", "Operating Expenses"]
            )

        monthly_profitability = pd.merge(
            monthly_summary,
            monthly_expense_summary,
            on="month",
            how="outer"
        ).fillna(0)

        monthly_profitability["Gross Profit"] = (
            monthly_profitability["Revenue"]
            - monthly_profitability["COGS"]
        )

        monthly_profitability["Net Profit"] = (
            monthly_profitability["Gross Profit"]
            - monthly_profitability["Operating Expenses"]
        )

        monthly_profitability = monthly_profitability.sort_values("month")

        if not monthly_profitability.empty:
            st.dataframe(
                monthly_profitability,
                use_container_width=True
            )

            st.line_chart(
                monthly_profitability.set_index("month")[
                    ["Revenue", "Gross Profit", "Net Profit"]
                ]
            )
            # REVENUE VS OPERATING EXPENSES
            st.subheader("Revenue vs Operating Expenses")

            comparison_data = monthly_profitability.set_index("month")[
                ["Revenue", "Operating Expenses"]
            ]

            st.bar_chart(comparison_data)
        else:
            st.info("No transactions available for the selected period.")
        # PROFITABILITY ANALYSIS
        st.subheader("Profitability Analysis")

        col1, col2, col3 = st.columns(3)

        col1.metric(
            "Total Revenue",
            f"₦{total_revenue:,.2f}"
        )

        col2.metric(
            "Cost of Goods Sold",
            f"₦{total_cogs:,.2f}"
        )

        col3.metric(
            "Gross Profit",
            f"₦{gross_profit:,.2f}"
        )

        col4, col5, col6 = st.columns(3)

        col4.metric(
            "Operating Expenses",
            f"₦{operating_expenses:,.2f}"
        )

        col5.metric(
            "Net Profit",
            f"₦{net_profit:,.2f}"
        )

        col6.metric(
            "Net Profit Margin",
            f"{net_profit_margin:.2f}%"
        )

        st.caption(
            f"Gross Profit Margin: {gross_profit_margin:.2f}%"
        )
        # EXPENSE BREAKDOWN BY CATEGORY
        st.subheader("Operating Expenses Breakdown")

        if not expenses_df.empty:
            expense_summary = (
                expenses_df.groupby("category")["amount"]
                .sum()
                .reset_index()
            )
            total_operating_expenses = expense_summary["amount"].sum()

            expense_summary = expense_summary.sort_values(
                by="amount", ascending=False
            )

            largest_expense_category = expense_summary.iloc[0]["category"]
            largest_expense_amount = expense_summary.iloc[0]["amount"]
            st.markdown("### Expense Summary")

            col1, col2, col3 = st.columns(3)

            col1.metric(
                "Total Operating Expenses",
                f"₦{total_operating_expenses:,.2f}"
            )

            col2.metric(
                "Largest Expense Category",
                largest_expense_category
            )

            col3.metric(
                "Amount in Largest Category",
                f"₦{largest_expense_amount:,.2f}"
            )

            expense_summary = expense_summary.sort_values(
                by="amount",
                ascending=False
            )

            expense_summary.columns = [
                "Expense Category",
                "Amount (₦)"
            ]

            st.dataframe(
                expense_summary,
                use_container_width=True,
                hide_index=True
            )

            st.bar_chart(
                expense_summary.set_index("Expense Category")
            )

        else:
            st.info("No expenses recorded for this period.")




    sales_df = get_sales()
    expenses_df = get_expenses()
    purchases_df = get_stock_purchases()

    # Revenue
    total_revenue = (
        sales_df["amount"].sum()
        if not sales_df.empty else 0
    )

    # Cost of Goods Sold
    total_cogs = (
        sales_df["cost_of_goods_sold"].sum()
        if not sales_df.empty else 0
    )

    # Gross Profit
    gross_profit = total_revenue - total_cogs

    # Operating Expenses
    operating_expenses = (
        expenses_df["amount"].sum()
        if not expenses_df.empty else 0
    )

    # Net Profit
    net_profit = gross_profit - operating_expenses

    st.divider()

    st.subheader("Income Statement")

    col1, col2 = st.columns(2)

    col1.metric(
        "Total Revenue",
        f"₦{total_revenue:,.2f}"
    )

    col2.metric(
        "Cost of Goods Sold",
        f"₦{total_cogs:,.2f}"
    )

    st.divider()

    col3, col4 = st.columns(2)

    col3.metric(
        "Gross Profit",
        f"₦{gross_profit:,.2f}"
    )

    col4.metric(
        "Operating Expenses",
        f"₦{operating_expenses:,.2f}"
    )

    st.divider()

    st.metric(
        "NET PROFIT",
        f"₦{net_profit:,.2f}"
    )

    
    # EXCEL EXPORT - PROFIT & LOSS

    st.subheader("Download Profit & Loss Report")

    pl_export = pd.DataFrame({
        "Description": [
            "Total Revenue",
            "Cost of Goods Sold",
            "Gross Profit",
            "Operating Expenses",
            "Net Profit"
        ],
        "Amount (₦)": [
            total_revenue,
            total_cogs,
            gross_profit,
            operating_expenses,
            net_profit
        ]
    })

    excel_buffer = BytesIO()

    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:

        pl_export.to_excel(
            writer,
            index=False,
            sheet_name="Profit & Loss",
            startrow=4
        )

        worksheet = writer.sheets["Profit & Loss"]

        # BUSINESS NAME
        worksheet.merge_cells("A1:B1")
        worksheet["A1"] = BUSINESS_NAME
        worksheet["A1"].font = Font(
            bold=True, size=16
        )

        # REPORT TITLE
        worksheet.merge_cells("A2:B2")
        worksheet["A2"] = "PROFIT & LOSS STATEMENT"
        worksheet["A2"].font = Font(
            bold=True, size=14
        )

        # REPORTING PERIOD
        worksheet.merge_cells("A3:B3")
        worksheet["A3"] = (
            f"Period: {start_date} to {end_date}"
        )
        worksheet["A3"].font = Font(
            italic=True, size=11
        )

        # COLUMN HEADINGS
        for cell in worksheet[5]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(
                start_color="1F4E78",
                end_color="1F4E78",
                fill_type="solid"
            )
            cell.alignment = Alignment(horizontal="center")

        # CURRENCY FORMATTING
        for row in range(6, 11):
            worksheet.cell(row=row, column=2).number_format = '"₦"#,##0.00'

        # HIGHLIGHT GROSS PROFIT AND NET PROFIT
        for row in [8, 10]:
            for cell in worksheet[row]:
                cell.font = Font(bold=True)

        # COLUMN WIDTHS
        worksheet.column_dimensions["A"].width = 30
        worksheet.column_dimensions["B"].width = 20

        # ALIGNMENT
        for row in range(6, 11):
            worksheet.cell(row=row, column=2).alignment = Alignment(
                horizontal="right"
            )

        worksheet.freeze_panes = "A6"

    excel_buffer.seek(0)

    st.download_button(
        label="📥 Download Profit & Loss (Excel)",
        data=excel_buffer,
        file_name="Profit_and_Loss_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # PDF EXPORT - PROFIT & LOSS

    st.subheader("Download Profit & Loss (PDF)")

    pdf_buffer = BytesIO()

    pdf = SimpleDocTemplate(
        pdf_buffer,
        pagesize=A4,
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=40
    )

    styles = getSampleStyleSheet()

    elements = []

    # BUSINESS NAME
    elements.append(
        Paragraph(f"<b>{BUSINESS_NAME}</b>", styles["Title"])
    )

    elements.append(Spacer(1, 10))

    # REPORT TITLE
    elements.append(
        Paragraph("PROFIT & LOSS STATEMENT", styles["Heading2"])
    )

    elements.append(Spacer(1, 8))

    # REPORTING PERIOD
    elements.append(
        Paragraph(
            f"Reporting Period: {start_date} to {end_date}",
            styles["Normal"]
        )
    )

    elements.append(Spacer(1, 20))

    # FINANCIAL DATA
    pdf_data = [
        ["Description", "Amount (NGN)"],
        ["Total Revenue", f"{total_revenue:,.2f}"],
        ["Cost of Goods Sold", f"{total_cogs:,.2f}"],
        ["Gross Profit", f"{gross_profit:,.2f}"],
        ["Operating Expenses", f"{operating_expenses:,.2f}"],
        ["Net Profit", f"{net_profit:,.2f}"]
    ]

    pdf_table = Table(
        pdf_data,
        colWidths=[300, 170]
    )

    pdf_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 3), (-1, 3), "Helvetica-Bold"),
        ("FONTNAME", (0, 5), (-1, 5), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (1, 1), (1, -1), "RIGHT"),
        ("BACKGROUND", (0, 5), (-1, 5), colors.lightgrey),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
    ]))

    elements.append(pdf_table)

    # GENERATE PDF
    pdf.build(elements)

    pdf_buffer.seek(0)

    st.download_button(
        label="📄 Download Profit & Loss (PDF)",
        data=pdf_buffer,
        file_name="Profit_and_Loss_Report.pdf",
        mime="application/pdf"
    )    

st.caption("SME Accountant | Grace Provision Store")
