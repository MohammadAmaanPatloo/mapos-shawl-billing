import streamlit as st
from reportlab.lib import colors
import pandas as pd
from openpyxl import Workbook, load_workbook
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    HRFlowable,
)
from urllib.parse import quote
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(page_title="Shawl Billing App", page_icon="🧣", layout="wide")


def number_to_words(number):
    ones = [
        "",
        "One",
        "Two",
        "Three",
        "Four",
        "Five",
        "Six",
        "Seven",
        "Eight",
        "Nine",
        "Ten",
        "Eleven",
        "Twelve",
        "Thirteen",
        "Fourteen",
        "Fifteen",
        "Sixteen",
        "Seventeen",
        "Eighteen",
        "Nineteen",
    ]

    tens = [
        "",
        "",
        "Twenty",
        "Thirty",
        "Forty",
        "Fifty",
        "Sixty",
        "Seventy",
        "Eighty",
        "Ninety",
    ]

    def convert_less_than_thousand(n):

        words = ""

        if n >= 100:
            words += ones[n // 100] + " Hundred"
            n %= 100

            if n:
                words += " "

        if n >= 20:
            words += tens[n // 10]
            n %= 10

            if n:
                words += " " + ones[n]

        elif n > 0:
            words += ones[n]

        return words

    number = float(number)

    rupees = int(number)
    paise = round((number - rupees) * 100)

    if rupees == 0:
        rupees_words = "Zero"
    else:
        parts = []

        crore = rupees // 10000000
        rupees %= 10000000

        lakh = rupees // 100000
        rupees %= 100000

        thousand = rupees // 1000
        rupees %= 1000

        if crore:
            parts.append(convert_less_than_thousand(crore) + " Crore")

        if lakh:
            parts.append(convert_less_than_thousand(lakh) + " Lakh")

        if thousand:
            parts.append(convert_less_than_thousand(thousand) + " Thousand")

        if rupees:
            parts.append(convert_less_than_thousand(rupees))

        rupees_words = " ".join(parts)

    if paise > 0:
        paise_words = convert_less_than_thousand(paise)

        return f"{rupees_words} Rupees and {paise_words} Paise Only"

    return f"{rupees_words} Rupees Only"


# ========================================================
# INDIA DATE & TIME
# ========================================================


def get_india_datetime():
    return datetime.now(ZoneInfo("Asia/Kolkata"))


# ============================================================
# SESSION STATE
# ============================================================

if "products" not in st.session_state:
    st.session_state.products = []

if "bill_no" not in st.session_state:
    st.session_state.bill_no = 1

if "product_form_reset" not in st.session_state:
    st.session_state.product_form_reset = 0

if "bill_number_reset" not in st.session_state:
    st.session_state.bill_number_reset = 0

if "bill_records" not in st.session_state:
    st.session_state.bill_records = []


# ============================================================
# PDF GENERATOR
# ============================================================


def generate_pdf(
    shop_name,
    shop_location,
    shop_phone,
    bill_no,
    customer_name,
    customer_phone,
    payment_method,
    products,
    subtotal,
    gst_rate,
    gst_amount,
    grand_total,
    total_items,
    instagram,
):

    buffer = BytesIO()

    # Receipt-style width similar to a thermal printer
    receipt_width = 80 * mm

    doc = SimpleDocTemplate(
        buffer,
        pagesize=(receipt_width, 250 * mm),
        rightMargin=5 * mm,
        leftMargin=5 * mm,
        topMargin=5 * mm,
        bottomMargin=5 * mm,
    )

    styles = getSampleStyleSheet()

    shop_style = ParagraphStyle(
        "ShopName",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=14,
        alignment=TA_CENTER,
        spaceAfter=2,
    )

    center_style = ParagraphStyle(
        "Center",
        parent=styles["Normal"],
        fontSize=8,
        alignment=TA_CENTER,
        leading=10,
    )
    normal_style = ParagraphStyle(
        "NormalCustom",
        parent=styles["Normal"],
        fontSize=8,
        leading=10,
    )

    right_style = ParagraphStyle(
        "RightStyle",
        parent=normal_style,
        alignment=TA_RIGHT,
    )

    bold_style = ParagraphStyle(
        "BoldCustom",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8,
        leading=10,
    )

    total_style = ParagraphStyle(
        "Total",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=10,
        leading=12,
    )

    story = []

    # ========================================================
    # SHOP HEADER
    # ========================================================

    # ========================================================
    # CLICKABLE SHOP NAME → GOOGLE MAPS
    # ========================================================

    maps_url = "https://www.google.com/maps/search/?api=1&query=" + quote(shop_location)

    shop_name_link = f'<link href="{maps_url}" color="black">{shop_name.upper()}</link>'

    story.append(Paragraph(shop_name_link, shop_style))

    story.append(Paragraph(shop_location, center_style))

    story.append(Paragraph(f"SHOP PHONE: {shop_phone}", center_style))

    story.append(Spacer(1, 5))

    # ========================================================
    # BILL INFORMATION
    # ========================================================

    now = get_india_datetime()

    bill_info = [
        [
            Paragraph(f"<b>Bill No:</b> {bill_no}", normal_style),
            Paragraph(f"<b>Date:</b> {now.strftime('%d/%m/%y')}", normal_style),
        ],
        [
            Paragraph(f"<b>Customer:</b> {customer_name}", normal_style),
            Paragraph(f"<b>Time:</b> {now.strftime('%H:%M:%S')}", normal_style),
        ],
        [
            Paragraph(f"<b>Customer Phone:</b> {customer_phone}", normal_style),
            "",
        ],
    ]

    info_table = Table(bill_info, colWidths=[38 * mm, 32 * mm])

    info_table.setStyle(
        TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ])
    )

    story.append(info_table)

    story.append(Spacer(1, 5))

    # ========================================================
    # PRODUCT TABLE
    # ========================================================

    product_data = [
        [
            Paragraph("<b>Item</b>", center_style),
            Paragraph("<b>Qty</b>", center_style),
            Paragraph("<b>Rate</b>", center_style),
            Paragraph("<b>Amt</b>", center_style),
        ]
    ]

    for product in products:
        amount = product["quantity"] * product["rate"]

        product_data.append([
            Paragraph(product["name"], normal_style),
            Paragraph(str(product["quantity"]), center_style),
            Paragraph(f"{product['rate']:,.2f}", right_style),
            Paragraph(f"{amount:,.2f}", right_style),
        ])

    product_table = Table(product_data, colWidths=[27 * mm, 10 * mm, 17 * mm, 21 * mm])

    product_table.setStyle(
        TableStyle([
            ("LINEABOVE", (0, 0), (-1, 0), 0.5, colors.black),
            ("LINEBELOW", (0, 0), (-1, 0), 0.5, colors.black),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 1),
            ("RIGHTPADDING", (0, 0), (-1, -1), 1),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ])
    )

    story.append(product_table)

    story.append(Spacer(1, 5))

    # ========================================================
    # TOTALS
    # ========================================================

    # Amount in words
    amount_in_words = number_to_words(grand_total)

    # Style for amount in words
    amount_words_style = ParagraphStyle(
        "AmountWords",
        parent=normal_style,
        fontName="Helvetica-Bold",
        fontSize=7.5,
        leading=10,
    )

    totals_data = [
        [
            Paragraph("Total Items", normal_style),
            Paragraph(str(total_items), bold_style),
        ],
        [
            Paragraph("Sub Total", normal_style),
            Paragraph(f"{subtotal:,.2f}", bold_style),
        ],
    ]

    # Add GST only if enabled
    if gst_rate > 0:
        totals_data.append([
            Paragraph(f"GST ({gst_rate:g}%)", normal_style),
            Paragraph(f"{gst_amount:,.2f}", normal_style),
        ])

    # ========================================================
    # GRAND TOTAL
    # ========================================================

    totals_data.append([
        Paragraph("GRAND TOTAL", total_style),
        Paragraph(f"{grand_total:,.2f}", total_style),
    ])

    # # ========================================================
    # # TOTALS TABLE
    # # ========================================================

    totals_table = Table(totals_data, colWidths=[40 * mm, 35 * mm])

    totals_table.setStyle(
        TableStyle([
            # Alignment
            ("ALIGN", (1, 0), (1, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            # No borders by default
            ("LINEABOVE", (0, 0), (-1, -1), 0, colors.white),
            ("LINEBELOW", (0, 0), (-1, -1), 0, colors.white),
            # Padding
            ("LEFTPADDING", (0, 0), (-1, -1), 1),
            ("RIGHTPADDING", (0, 0), (-1, -1), 1),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            # ================================================
            # LINE ABOVE TOTAL ITEMS
            # ================================================
            ("LINEABOVE", (0, 0), (-1, 0), 0.8, colors.black),
            # ================================================
            # LINE BELOW TOTAL ITEMS
            # ================================================
            ("LINEBELOW", (0, 0), (-1, 0), 0.8, colors.black),
            # ================================================
            # LINE BELOW SUB TOTAL
            # ================================================
            ("LINEBELOW", (0, 1), (-1, 1), 0.8, colors.black),
        ])
    )

    story.append(totals_table)

    # ========================================================
    # LINE BELOW GST
    # ========================================================

    if gst_rate > 0:
        # GST is row 2 when GST is enabled
        totals_table.setStyle(
            TableStyle([
                ("LINEBELOW", (0, 2), (-1, 2), 0.8, colors.black),
            ])
        )

    # ========================================================
    # LINE AFTER GRAND TOTAL
    # ========================================================

    story.append(Spacer(1, 3))

    story.append(
        HRFlowable(
            width="100%",
            thickness=0.8,
            color=colors.black,
            spaceBefore=1,
            spaceAfter=5,
        )
    )

    # ========================================================
    # AMOUNT IN WORDS
    # ========================================================

    story.append(
        Paragraph(f"<b>Amount in Words:</b><br/>{amount_in_words}", amount_words_style)
    )

    story.append(Spacer(1, 4))

    story.append(
        HRFlowable(
            width="100%",
            thickness=0.8,
            color=colors.black,
            spaceBefore=1,
            spaceAfter=5,
        )
    )

    story.append(Spacer(1, 3))

    # ========================================================
    # PAYMENT
    # ========================================================

    payment_table = Table(
        [
            [
                Paragraph("<b>Payment Mode</b>", normal_style),
                Paragraph(payment_method, normal_style),
            ]
        ],
        colWidths=[40 * mm, 35 * mm],
    )

    payment_table.setStyle(
        TableStyle([
            ("ALIGN", (1, 0), (1, 0), "RIGHT"),
            ("LEFTPADDING", (0, 0), (-1, -1), 1),
            ("RIGHTPADDING", (0, 0), (-1, -1), 1),
        ])
    )

    story.append(payment_table)

    story.append(Spacer(1, 5))

    # Line after Payment Mode
    story.append(
        HRFlowable(
            width="100%",
            thickness=0.8,
            color=colors.black,
            spaceBefore=1,
            spaceAfter=7,
        )
    )

    # ========================================================
    # FOOTER
    # ========================================================

    story.append(Paragraph("Follow Us on Instagram", center_style))

    # Instagram clickable link
    instagram_username = instagram.strip().replace("@", "")

    instagram_link = (
        f'<link href="https://www.instagram.com/{instagram_username}" '
        f'color="blue">@{instagram_username}</link>'
    )

    story.append(
        Paragraph(
            instagram_link,
            ParagraphStyle(
                "Instagram",
                parent=center_style,
                fontName="Helvetica-Bold",
                underline=True,
            ),
        )
    )

    story.append(Spacer(1, 8))

    story.append(
        Paragraph(
            "Thank You. Visit Us Again!",
            ParagraphStyle(
                "ThankYou",
                parent=center_style,
                fontName="Helvetica-Bold",
                fontSize=9,
            ),
        )
    )

    story.append(Spacer(1, 8))

    # ========================================================
    # CLICKABLE POWERED BY MAP
    # ========================================================

    mapos_link = "https://map-portfolio.netlify.app/"

    powered_by_map = f'<link href="{mapos_link}" color="black">Powered by MAPOS</link>'

    story.append(Paragraph(powered_by_map, center_style))

    doc.build(story)

    buffer.seek(0)

    return buffer.getvalue()


# ========================================================
# CREATE EXCEL FILE
# ========================================================


def create_excel_file(bill_records):

    columns = [
        "Bill No",
        "Date",
        "Time",
        "Customer Name",
        "Customer Number",
        "Item",
        "Quantity",
        "Rate",
        "Amount",
        "Sub Total",
        "GST %",
        "GST Amount",
        "Grand Total",
        "Payment Mode",
    ]

    df = pd.DataFrame(bill_records, columns=columns)

    excel_buffer = BytesIO()

    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Bills")

        worksheet = writer.sheets["Bills"]

        # Freeze header row
        worksheet.freeze_panes = "A2"

        # Auto-adjust column widths
        for column_cells in worksheet.columns:
            max_length = 0

            column_letter = column_cells[0].column_letter

            for cell in column_cells:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))

            worksheet.column_dimensions[column_letter].width = min(max_length + 2, 30)

    excel_buffer.seek(0)

    return excel_buffer.getvalue()


# ============================================================
# APP HEADER
# ============================================================

st.title("🧣 Shawl Billing App")
st.caption("Simple billing system for your shawl/stole business")


# ============================================================
# SIDEBAR - SHOP DETAILS
# ============================================================

st.sidebar.title("🏪 Shop Details")

shop_name = st.sidebar.text_input("Shop Name", value="Sajad Arts")

shop_location = st.sidebar.text_input(
    "Shop Location", value="Srinagar, Jammu & Kashmir"
)

shop_phone = st.sidebar.text_input("Shop Phone Number", value="8494001112")

instagram = st.sidebar.text_input("Instagram Username", value="@mohammadamaan32")

st.sidebar.divider()


# ============================================================
# SIDEBAR - CUSTOMER DETAILS
# ============================================================

st.sidebar.title("👤 Customer Details")

customer_name = st.sidebar.text_input("Customer Name")

customer_phone = st.sidebar.text_input("Customer Phone Number")

payment_method = st.sidebar.selectbox(
    "Payment Method", ["Cash", "UPI", "Card", "Bank Transfer", "Other"]
)

st.sidebar.divider()


# ============================================================
# BILL NUMBER
# ============================================================

st.sidebar.title("🧾 Bill Information")

bill_no = st.sidebar.number_input(
    "Bill Number",
    min_value=1,
    value=st.session_state.bill_no,
    step=1,
    key=f"bill_number_{st.session_state.bill_number_reset}",
)

# ============================================================
# PRODUCT ENTRY / EDIT PRODUCT
# ============================================================

st.header("Add Products")

# Check whether we are editing a product
if "editing_product" not in st.session_state:
    st.session_state.editing_product = None


if st.session_state.editing_product is not None:
    edit_index = st.session_state.editing_product
    edit_product = st.session_state.products[edit_index]

    st.info(f"✏️ Editing Product: {edit_product['name']}")

    # --------------------------------------------------------
    # PRODUCT OPTIONS
    # --------------------------------------------------------

    product_options = [
        "Hashidar",
        "Koundar",
        "Pointdar",
        "Paldar",
        "Jaildar",
        "Plain Colour",
        "Zaati",
        "Ari",
        "➕ Add New Product...",
    ]

    # If the existing product is not in the standard list,
    # add it temporarily so it can still be selected.
    current_product = edit_product["name"]

    if current_product not in product_options:
        product_options.insert(0, current_product)

    # --------------------------------------------------------
    # INPUT FIELDS
    # --------------------------------------------------------

    col1, col2, col3, col4 = st.columns([3, 2, 2.5, 1])

    # --------------------------------------------------------
    # PRODUCT NAME
    # --------------------------------------------------------

    with col1:
        selected_edit_product = st.selectbox(
            "Product Name",
            product_options,
            index=product_options.index(current_product),
            key="edit_product_select",
        )

        # If user chooses Add New Product
        if selected_edit_product == "➕ Add New Product...":
            product_name = st.text_input(
                "Enter New Product Name",
                placeholder="Type new product name",
                key="edit_new_product_name",
            )

        else:
            product_name = selected_edit_product

    # --------------------------------------------------------
    # QUANTITY
    # --------------------------------------------------------

    with col2:
        quantity = st.number_input(
            "Quantity",
            min_value=1,
            value=edit_product["quantity"],
            step=1,
            format="%d",
            key="edit_quantity",
        )

    # --------------------------------------------------------
    # RATE
    # --------------------------------------------------------

    with col3:
        rate = st.number_input(
            "Rate",
            min_value=0.0,
            value=float(edit_product["rate"]),
            step=50.0,
            key="edit_rate",
        )

    # --------------------------------------------------------
    # UPDATE BUTTON
    # --------------------------------------------------------

    with col4:
        st.write("")
        st.write("")

        update_product = st.button("💾 Update", use_container_width=True)

    # --------------------------------------------------------
    # CANCEL BUTTON
    # --------------------------------------------------------

    cancel_edit = st.button("Cancel Edit")

    # --------------------------------------------------------
    # UPDATE PRODUCT
    # --------------------------------------------------------

    if update_product:
        if product_name.strip() == "":
            st.error("Please enter the product name.")

        elif rate <= 0:
            st.error("Please enter a valid rate.")

        else:
            st.session_state.products[edit_index] = {
                "name": product_name,
                "quantity": quantity,
                "rate": rate,
            }

            st.session_state.editing_product = None

            st.success("✅ Product updated successfully!")

            st.rerun()

    # --------------------------------------------------------
    # CANCEL EDIT
    # --------------------------------------------------------

    if cancel_edit:
        st.session_state.editing_product = None

        st.rerun()

else:
    col1, col2, col3, col4 = st.columns([3, 2, 2.5, 1])

    with col1:
        product_options = [
            "Hashidar",
            "Koundar",
            "Pointdar",
            "Paldar",
            "Jaildar",
            "Plain Colour",
            "Zaati",
            "Ari",
            "➕ Add New Product...",
        ]

        selected_product = st.selectbox(
            "Product Name", product_options, key="selected_product"
        )

        if selected_product == "➕ Add New Product...":
            product_name = st.text_input(
                "Enter New Product Name",
                placeholder="Type new product name",
                key="new_product_name",
            )

        else:
            product_name = selected_product

    with col2:
        quantity = st.number_input(
            "Quantity",
            min_value=0,
            value=0,
            step=1,
            format="%d",
            key=f"new_quantity_{st.session_state.product_form_reset}",
        )

    with col3:
        rate = st.number_input(
            "Rate",
            min_value=0.0,
            value=0.0,
            step=50.0,
            key=f"new_rate_{st.session_state.product_form_reset}",
        )

    with col4:
        st.write("")
        st.write("")

        add_product = st.button("➕ Add", use_container_width=True)

    # --------------------------------------------------------
    # ADD PRODUCT
    # --------------------------------------------------------

    if add_product:
        if product_name.strip() == "":
            st.error("Please enter the product name.")

        elif quantity <= 0:
            st.error("Please enter a quantity greater than 0.")

        elif rate <= 0:
            st.error("Please enter a valid rate.")

        else:
            product = {"name": product_name, "quantity": quantity, "rate": rate}

            st.session_state.products.append(product)

            # Create fresh Quantity and Rate widgets
            st.session_state.product_form_reset += 1

            st.success(f"{product_name} added successfully!")

            st.rerun()

# ============================================================
# PRODUCT LIST
# ============================================================

st.header("Product List")

if len(st.session_state.products) == 0:
    st.info("No products added yet. Add products using the form above.")

else:
    # Table header
    header = st.columns([4, 1, 2, 2, 2])

    header[0].write("**Product**")
    header[1].write("**Qty**")
    header[2].write("**Rate**")
    header[3].write("**Amount**")
    header[4].write("**Action**")

    st.divider()

    for index, product in enumerate(st.session_state.products):
        amount = product["quantity"] * product["rate"]

        row = st.columns([4, 1, 2, 2, 2])

        row[0].write(product["name"])

        row[1].write(product["quantity"])

        row[2].write(f"{product['rate']:,.2f}")

        row[3].write(f"{amount:,.2f}")

        # ----------------------------------------------------
        # EDIT BUTTON
        # ----------------------------------------------------

        action_col1, action_col2 = row[4].columns(2)

        if action_col1.button("✏️", key=f"edit_{index}", help="Edit product"):
            st.session_state.editing_product = index

            st.rerun()

        # ----------------------------------------------------
        # DELETE BUTTON
        # ----------------------------------------------------

        if action_col2.button("❌", key=f"delete_{index}", help="Delete product"):
            st.session_state.products.pop(index)

            # If the deleted product was being edited
            if st.session_state.editing_product == index:
                st.session_state.editing_product = None

            st.rerun()


# ============================================================
# CALCULATIONS
# ============================================================

subtotal = 0
total_items = 0

for product in st.session_state.products:
    subtotal += product["quantity"] * product["rate"]

    total_items += product["quantity"]


# ============================================================
# GST
# ============================================================

st.divider()

st.header("Bill Calculation")

gst_enabled = st.checkbox("Add GST")

gst_rate = 0.0

if gst_enabled:
    gst_rate = st.number_input(
        "GST Rate (%)", min_value=0.0, max_value=100.0, value=5.0, step=1.0
    )

gst_amount = subtotal * gst_rate / 100

grand_total = subtotal + gst_amount

amount_in_words = number_to_words(grand_total)


# ============================================================
# BILL SUMMARY
# ============================================================

summary_col1, summary_col2 = st.columns(2)

with summary_col1:
    st.metric("Total Items", total_items)

with summary_col2:
    st.metric("Grand Total", f"₹{grand_total:,.2f}")


st.write("### Bill Summary")

summary_data = {
    "Subtotal": f"₹{subtotal:,.2f}",
}

if gst_enabled:
    summary_data[f"GST ({gst_rate:g}%)"] = f"₹{gst_amount:,.2f}"

summary_data["Grand Total"] = f"₹{grand_total:,.2f}"

for label, value in summary_data.items():
    col1, col2 = st.columns(2)

    col1.write(f"**{label}**")
    col2.write(f"**{value}**")


# ============================================================
# GENERATE BILL
# ============================================================

st.divider()

st.header("Generate Bill")

generate_bill = st.button("🧾 Generate Bill", type="primary", use_container_width=True)


if generate_bill:
    # Validate customer
    if customer_name.strip() == "":
        st.error("Please enter the customer name.")

    elif customer_phone.strip() == "":
        st.error("Please enter the customer phone number.")

    elif len(st.session_state.products) == 0:
        st.error("Please add at least one product.")

    else:
        pdf_bytes = generate_pdf(
            shop_name=shop_name,
            shop_location=shop_location,
            shop_phone=shop_phone,
            bill_no=bill_no,
            customer_name=customer_name,
            customer_phone=customer_phone,
            payment_method=payment_method,
            products=st.session_state.products,
            subtotal=subtotal,
            gst_rate=gst_rate,
            gst_amount=gst_amount,
            grand_total=grand_total,
            total_items=total_items,
            instagram=instagram,
        )

        # Store PDF in session
        st.session_state.pdf_bytes = pdf_bytes
        st.session_state.generated_bill_no = bill_no
        # ====================================================
        # SAVE BILL DATA
        # ====================================================

        now = get_india_datetime()

        bill_date = now.strftime("%d/%m/%y")
        bill_time = now.strftime("%H:%M:%S")

        for product in st.session_state.products:
            amount = product["quantity"] * product["rate"]

            st.session_state.bill_records.append({
                "Bill No": bill_no,
                "Date": bill_date,
                "Time": bill_time,
                "Customer Name": customer_name,
                "Customer Number": customer_phone,
                "Item": product["name"],
                "Quantity": product["quantity"],
                "Rate": product["rate"],
                "Amount": amount,
                "Sub Total": subtotal,
                "GST %": gst_rate,
                "GST Amount": gst_amount,
                "Grand Total": grand_total,
                "Payment Mode": payment_method,
            })

        st.success("✅ Bill generated successfully!")


# ============================================================
# DOWNLOAD + WHATSAPP
# ============================================================

if "pdf_bytes" in st.session_state:
    st.divider()

    st.subheader("Bill Ready")

    filename = (
        f"Bill_{st.session_state.generated_bill_no}_"
        f"{customer_name.replace(' ', '_')}.pdf"
    )

    st.download_button(
        label="⬇️ Download Bill PDF",
        data=st.session_state.pdf_bytes,
        file_name=filename,
        mime="application/pdf",
        use_container_width=True,
    )

    # ========================================================
    # WHATSAPP
    # ========================================================

    whatsapp_number = "".join(
        character for character in customer_phone if character.isdigit()
    )

    # If Indian customer number entered as 10 digits
    if len(whatsapp_number) == 10:
        whatsapp_number = "91" + whatsapp_number

    # ========================================================
    # SHOP WHATSAPP NUMBER
    # ========================================================

    shop_whatsapp_number = "".join(
        character for character in shop_phone if character.isdigit()
    )

    # If Indian shop number entered as 10 digits
    if len(shop_whatsapp_number) == 10:
        shop_whatsapp_number = "91" + shop_whatsapp_number

    # ========================================================
    # CLICKABLE LINKS
    # ========================================================

    # Shop → Google Maps
    maps_url = "https://www.google.com/maps/search/?api=1&query=" + quote(shop_location)

    # Instagram
    instagram_username = instagram.strip().replace("@", "")

    instagram_url = f"https://www.instagram.com/{instagram_username}"

    # MAPOS website
    mapos_url = "https://map-portfolio.netlify.app/"

    # Shop WhatsApp
    shop_whatsapp_url = f"https://wa.me/{shop_whatsapp_number}"

    # ========================================================
    # WHATSAPP MESSAGE
    # ========================================================

    whatsapp_message = f"""Hello {customer_name},

    Thank you for shopping with {shop_name}.

    Bill No: {bill_no}
    Total Items: {total_items}
    Grand Total: ₹{grand_total:,.2f}

    Please find your bill attached.

    Shop Location:
    {maps_url}

    Shop WhatsApp:
    {shop_whatsapp_url}

    Instagram:
    {instagram_url}

    Powered by MAPOS
    {mapos_url}

    Thank you. Visit Us Again!"""

    # ========================================================
    # ENCODE MESSAGE ONCE
    # ========================================================

    whatsapp_url = f"https://wa.me/{whatsapp_number}?text={quote(whatsapp_message)}"

    # ========================================================
    # SEND BUTTON
    # ========================================================

    st.markdown(
        f"""
        <a href="{whatsapp_url}" target="_blank">
            <button style="
                width: 100%;
                padding: 12px;
                background-color: #25D366;
                color: white;
                border: none;
                border-radius: 8px;
                font-size: 16px;
                cursor: pointer;
            ">
                📱 Send Bill on WhatsApp
            </button>
        </a>
        """,
        unsafe_allow_html=True,
    )

st.info(
    "WhatsApp will open with the customer's number and "
    "message already prepared. Attach the downloaded PDF "
    "and press Send."
)

# ========================================================
# DOWNLOAD EXCEL
# ========================================================

if len(st.session_state.bill_records) > 0:
    excel_bytes = create_excel_file(st.session_state.bill_records)

    st.download_button(
        label="📊 Download Bills Excel",
        data=excel_bytes,
        file_name="Sajad_Arts_Bills.xlsx",
        mime=("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        use_container_width=True,
    )
# ============================================================
# NEW BILL
# ============================================================

st.divider()

if st.button("🔄 Start New Bill", use_container_width=True):
    # Clear current products
    st.session_state.products = []

    # Remove previous PDF
    if "pdf_bytes" in st.session_state:
        del st.session_state.pdf_bytes

    # Increase bill number
    st.session_state.bill_no += 1

    # Force Bill Number widget to refresh
    st.session_state.bill_number_reset += 1

    # Reset product entry form
    st.session_state.product_form_reset += 1

    st.rerun()
