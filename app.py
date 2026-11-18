import streamlit as st
import pdfplumber
import pandas as pd
import zipfile
import io
import re
import os
from datetime import datetime

# ============================================================
# CONFIG
# ============================================================

st.set_page_config(
    page_title="Invoice PDF to Excel",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Invoice PDF → Excel")
st.caption(
    "Upload ZIP berisi invoice PDF. Sistem membaca seluruh halaman "
    "dan mengambil SEMUA nomor invoice yang unik."
)

# ============================================================
# MASTER BRAND
# ============================================================

BRANDS = [
    "Hanasui",
    "NCO",
    "Fyne",
    "Cave",
    "Yummys",
    "BHST",
    "Essens",
    "BMYT",
    "Hexxa",
    "NDR",
    "Jamu Jago",
    "Teratu",
    "Ultimate U",
    "Bumi Tauhid",
    "Safwan",
    "Attaqwa",
    "Aqshor",
    "Philips",
    "Focallure",
    "Pinkflash",
]

# ============================================================
# MASTER WAREHOUSE
# ============================================================

WAREHOUSES = [
    "Medan",
    "Padang",
    "Palembang",
    "Yogyakarta",
    "Surabaya",
    "Malang",
    "Makassar",
    "Banjarmasin",
    "Bekasi",
    "Jakarta",
]


# ============================================================
# HELPER
# ============================================================

def clean_text(text):
    if not text:
        return ""

    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)

    return text.strip()


# ============================================================
# NORMALIZE NOMINAL
# ============================================================

def normalize_number(value):

    if value is None:
        return None

    value = str(value).strip()

    # Hilangkan Rp
    value = re.sub(
        r"(?i)\brp\.?\s*",
        "",
        value
    ).strip()

    if not value:
        return None

    # Format Indonesia:
    # 151.740,00
    if "," in value and "." in value:

        if value.rfind(",") > value.rfind("."):

            value = value.replace(".", "")
            value = value.replace(",", ".")

        else:

            value = value.replace(",", "")

    # Contoh:
    # 151.740,00
    elif "," in value:

        decimal_part = value.rsplit(",", 1)[-1]

        if len(decimal_part) <= 2:

            value = value.replace(".", "")
            value = value.replace(",", ".")

        else:

            value = value.replace(",", "")

    # Contoh:
    # 151.740
    elif "." in value:

        value = value.replace(".", "")

    value = re.sub(
        r"[^\d.\-]",
        "",
        value
    )

    try:

        number = float(value)

        if number.is_integer():
            return int(number)

        return number

    except:

        return None


# ============================================================
# DETECT BRAND
# ============================================================

def detect_brand(text, filename=""):

    combined = (
        f"{text}\n{filename}"
    ).lower()

    # Brand terpanjang diperiksa lebih dahulu
    sorted_brands = sorted(
        BRANDS,
        key=len,
        reverse=True
    )

    for brand in sorted_brands:

        if brand.lower() in combined:

            return brand

    return ""


# ============================================================
# DETECT SEMUA NOMOR INVOICE
# ============================================================

def detect_invoice_numbers(text):

    """
    Mengambil SEMUA nomor invoice yang terdapat
    pada halaman PDF.

    Contoh:
    SI.2026.11.00265
    SI.2026.11.00160
    """

    patterns = [

        # Format utama invoice
        r"\bSI\.\d{4}\.\d{2}\.\d+\b",

        # Format INV
        r"\bINV[./-]?[A-Z0-9./-]+\b",

        # Invoice dengan label
        r"(?:invoice\s*(?:no|number|#)|"
        r"no\.?\s*invoice|"
        r"nomor\s*invoice)"
        r"\s*[:#-]?\s*([A-Z0-9./-]+)",
    ]

    found = []

    for pattern in patterns:

        matches = re.finditer(
            pattern,
            text,
            flags=re.IGNORECASE
        )

        for match in matches:

            if match.lastindex:

                value = match.group(1).strip()

            else:

                value = match.group(0).strip()

            if value and value not in found:

                found.append(value)

    return found


# ============================================================
# DETECT INVOICE PERTAMA
# ============================================================

def detect_invoice_number(text):

    numbers = detect_invoice_numbers(text)

    if numbers:

        return numbers[0]

    return ""


# ============================================================
# DETECT DATE
# ============================================================

def detect_invoice_date(text):

    patterns = [

        r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{4})\b",

        r"\b("
        r"\d{1,2}\s+"
        r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
        r"\s+\d{4}"
        r")\b",
    ]

    for pattern in patterns:

        matches = re.finditer(
            pattern,
            text,
            flags=re.IGNORECASE
        )

        for match in matches:

            value = match.group(1)

            for fmt in [
                "%d/%m/%Y",
                "%d-%m-%Y",
                "%d %b %Y",
            ]:

                try:

                    return datetime.strptime(
                        value,
                        fmt
                    ).strftime("%d/%m/%Y")

                except:

                    pass

    return ""


# ============================================================
# DETECT STORE
# ============================================================

def detect_store(text):

    patterns = [

        r"(?:customer|pelanggan|customer name|nama pelanggan)"
        r"\s*[:\-]\s*(.+)",

        r"(?:bill\s*to|ship\s*to)"
        r"\s*[:\-]\s*(.+)",

        r"(?:store|toko|outlet)"
        r"\s*[:\-]\s*(.+)",
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE
        )

        if match:

            result = match.group(1)

            result = result.split("\n")[0].strip()

            if len(result) > 1:

                return result

    # Fallback
    blacklist = [
        "invoice",
        "tanggal",
        "date",
        "total",
        "subtotal",
        "alamat",
        "address",
        "phone",
        "telp",
        "produk",
        "qty",
        "harga",
        "warehouse",
        "nomor",
    ]

    lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]

    for line in lines[:40]:

        lower = line.lower()

        if any(
            word in lower
            for word in blacklist
        ):
            continue

        if (
            4 <= len(line) <= 80
            and not re.fullmatch(
                r"[\d\s.,:/\-]+",
                line
            )
        ):

            return line

    return ""


# ============================================================
# DETECT TOTAL
# ============================================================

def detect_total(text):

    """
    Mengambil TOTAL invoice.

    Prioritas:
    1. Grand Total
    2. Total Invoice
    3. Total

    Jika invoice terdiri dari beberapa halaman,
    sistem mengambil TOTAL terakhir yang ditemukan.
    """

    lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]

    priority_matches = []
    normal_matches = []

    for i, line in enumerate(lines):

        has_total = re.search(
            r"\btotal\b",
            line,
            flags=re.IGNORECASE
        )

        if not has_total:
            continue

        candidates = re.findall(
            r"(?:Rp\s*)?[\d.,]+",
            line,
            flags=re.IGNORECASE
        )

        values = [
            normalize_number(x)
            for x in candidates
        ]

        values = [
            x for x in values
            if x is not None
        ]

        if values:

            item = (
                i,
                values[-1]
            )

            if re.search(
                r"\b(grand\s+total|total\s+invoice)\b",
                line,
                flags=re.IGNORECASE
            ):

                priority_matches.append(item)

            else:

                normal_matches.append(item)

            continue

        # Kalau angka berada di baris berikutnya
        for next_line in lines[i + 1:i + 3]:

            candidates = re.findall(
                r"(?:Rp\s*)?[\d.,]+",
                next_line,
                flags=re.IGNORECASE
            )

            values = [
                normalize_number(x)
                for x in candidates
            ]

            values = [
                x for x in values
                if x is not None
            ]

            if values:

                item = (
                    i,
                    values[-1]
                )

                if re.search(
                    r"\b(grand\s+total|total\s+invoice)\b",
                    line,
                    flags=re.IGNORECASE
                ):

                    priority_matches.append(item)

                else:

                    normal_matches.append(item)

                break

    # Ambil TOTAL terakhir
    if priority_matches:

        return priority_matches[-1][1]

    if normal_matches:

        return normal_matches[-1][1]

    # Fallback
    rp_values = re.findall(
        r"(?:Rp\s*)[\d.,]+",
        text,
        flags=re.IGNORECASE
    )

    values = [
        normalize_number(x)
        for x in rp_values
    ]

    values = [
        x for x in values
        if x is not None
    ]

    if values:

        return max(values)

    return None


# ============================================================
# EXTRACT SEMUA HALAMAN PDF
# ============================================================

def extract_pdf_pages(
    pdf_bytes,
    filename
):

    pages = []

    try:

        with pdfplumber.open(
            io.BytesIO(pdf_bytes)
        ) as pdf:

            for page_no, page in enumerate(
                pdf.pages,
                start=1
            ):

                text = page.extract_text(
                    x_tolerance=2,
                    y_tolerance=2
                ) or ""

                text = clean_text(text)

                pages.append({

                    "page": page_no,

                    "text": text,

                    "invoices":
                        detect_invoice_numbers(text),

                })

    except Exception as e:

        return None, str(e)

    return pages, ""


# ============================================================
# GROUP INVOICE
# ============================================================

def build_invoice_groups(pages):

    """
    SATU NOMOR INVOICE = SATU BARIS.

    Kalau invoice yang sama mempunyai beberapa halaman,
    semua halaman digabung.

    Contoh:

    PDF:
    Halaman 1 -> SI.001
    Halaman 2 -> SI.001
    Halaman 3 -> SI.002
    Halaman 4 -> SI.002
    Halaman 5 -> SI.003

    Hasil:

    SI.001
    SI.002
    SI.003
    """

    groups = {}

    order = []

    current_invoice = None

    unnamed_counter = 0

    for page in pages:

        page_text = page["text"]

        invoice_numbers = page["invoices"]

        # ====================================================
        # HALAMAN MEMILIKI NOMOR INVOICE
        # ====================================================

        if invoice_numbers:

            for invoice_no in invoice_numbers:

                if invoice_no not in groups:

                    groups[invoice_no] = {

                        "invoice":
                            invoice_no,

                        "pages": [],

                        "page_numbers": [],
                    }

                    order.append(
                        invoice_no
                    )

                groups[
                    invoice_no
                ]["pages"].append(
                    page_text
                )

                groups[
                    invoice_no
                ]["page_numbers"].append(
                    page["page"]
                )

            current_invoice = (
                invoice_numbers[-1]
            )

        # ====================================================
        # HALAMAN LANJUTAN
        # ====================================================

        else:

            if (
                current_invoice
                and
                current_invoice in groups
            ):

                groups[
                    current_invoice
                ]["pages"].append(
                    page_text
                )

                groups[
                    current_invoice
                ]["page_numbers"].append(
                    page["page"]
                )

            elif page_text:

                unnamed_counter += 1

                key = (
                    f"__NO_INVOICE_"
                    f"{unnamed_counter}"
                )

                groups[key] = {

                    "invoice": "",

                    "pages": [
                        page_text
                    ],

                    "page_numbers": [
                        page["page"]
                    ],
                }

                order.append(key)

                current_invoice = key

    return [
        groups[key]
        for key in order
    ]


# ============================================================
# PROCESS ZIP
# ============================================================

def process_zip(
    uploaded_file
):

    results = []

    debug_info = []

    # Dipakai agar nomor invoice yang sama
    # tidak muncul dua kali walaupun ada di
    # dua PDF berbeda.
    seen_invoice_numbers = set()

    try:

        zip_bytes = uploaded_file.read()

        with zipfile.ZipFile(
            io.BytesIO(zip_bytes),
            "r"
        ) as z:

            pdf_files = [

                name

                for name in z.namelist()

                if (
                    name.lower().endswith(".pdf")
                    and
                    not name.endswith("/")
                )
            ]

            if not pdf_files:

                st.error(
                    "ZIP tidak berisi file PDF."
                )

                return (
                    pd.DataFrame(),
                    pd.DataFrame()
                )

            progress = st.progress(0)

            for index, pdf_name in enumerate(
                pdf_files
            ):

                try:

                    pdf_bytes = z.read(
                        pdf_name
                    )

                    pages, error = (
                        extract_pdf_pages(
                            pdf_bytes,
                            os.path.basename(
                                pdf_name
                            )
                        )
                    )

                    if pages is None:

                        debug_info.append({

                            "File":
                                pdf_name,

                            "Status":
                                "Gagal",

                            "Halaman":
                                0,

                            "Invoice ditemukan":
                                0,

                            "Invoice unik ditambahkan":
                                0,

                            "Error":
                                error,
                        })

                        progress.progress(
                            (index + 1)
                            /
                            len(pdf_files)
                        )

                        continue

                    # ====================================================
                    # INI BAGIAN PALING PENTING
                    # ====================================================

                    groups = (
                        build_invoice_groups(
                            pages
                        )
                    )

                    added_this_pdf = 0

                    invoices_in_pdf = []

                    for group in groups:

                        invoice_no = (
                            group["invoice"]
                        )

                        combined_text = (
                            clean_text(
                                "\n".join(
                                    group["pages"]
                                )
                            )
                        )

                        # ====================================================
                        # CEK DUPLIKAT INVOICE
                        # ====================================================

                        if invoice_no:

                            normalized_key = re.sub(
                                r"\s+",
                                "",
                                invoice_no
                            ).upper()

                            if (
                                normalized_key
                                in seen_invoice_numbers
                            ):

                                continue

                            seen_invoice_numbers.add(
                                normalized_key
                            )

                        # ====================================================
                        # BUAT DATA INVOICE
                        # ====================================================

                        row = {

                            "No":
                                len(results) + 1,

                            "Tanggal Invoice":
                                detect_invoice_date(
                                    combined_text
                                ),

                            "Nama Store":
                                detect_store(
                                    combined_text
                                ),

                            "Brand":
                                detect_brand(
                                    combined_text,
                                    pdf_name
                                ),

                            "Nomor Invoice":
                                invoice_no,

                            "Nominal":
                                detect_total(
                                    combined_text
                                ),

                            "Warehouse":
                                "",

                            "Nomor Faktur Pajak":
                                "",
                        }

                        results.append(
                            row
                        )

                        added_this_pdf += 1

                        invoices_in_pdf.append(
                            invoice_no
                            if invoice_no
                            else
                            "(tanpa nomor)"
                        )

                except Exception as e:

                    debug_info.append({

                        "File":
                            pdf_name,

                        "Status":
                            "Gagal",

                        "Halaman":
                            0,

                        "Invoice ditemukan":
                            0,

                        "Invoice unik ditambahkan":
                            0,

                        "Error":
                            str(e),
                    })

                else:

                    debug_info.append({

                        "File":
                            pdf_name,

                        "Status":
                            "OK",

                        "Halaman":
                            len(pages),

                        "Invoice ditemukan":
                            len(groups),

                        "Invoice unik ditambahkan":
                            added_this_pdf,

                        "Nomor Invoice":
                            ", ".join(
                                invoices_in_pdf
                            ),
                    })

                progress.progress(
                    (index + 1)
                    /
                    len(pdf_files)
                )

    except zipfile.BadZipFile:

        st.error(
            "File yang diupload bukan ZIP yang valid."
        )

        return (
            pd.DataFrame(),
            pd.DataFrame()
        )

    return (
        pd.DataFrame(results),
        pd.DataFrame(debug_info)
    )


# ============================================================
# UPLOAD ZIP
# ============================================================

uploaded_zip = st.file_uploader(
    "Upload file ZIP berisi invoice PDF",
    type=["zip"]
)


if uploaded_zip:

    if (
        "invoice_data"
        not in st.session_state
    ):

        st.session_state.invoice_data = (
            pd.DataFrame()
        )

    if (
        "debug_data"
        not in st.session_state
    ):

        st.session_state.debug_data = (
            pd.DataFrame()
        )

    # ========================================================
    # PROSES
    # ========================================================

    if st.button(
        "🔍 Proses ZIP",
        type="primary"
    ):

        with st.spinner(
            "Membaca semua PDF, semua halaman, "
            "dan semua invoice unik..."
        ):

            df, debug_df = process_zip(
                uploaded_zip
            )

            st.session_state.invoice_data = (
                df
            )

            st.session_state.debug_data = (
                debug_df
            )

    # ========================================================
    # PREVIEW
    # ========================================================

    df = st.session_state.get(
        "invoice_data",
        pd.DataFrame()
    )

    if not df.empty:

        st.success(
            f"Berhasil membaca "
            f"{len(df)} invoice unik "
            f"dari seluruh PDF."
        )

        st.subheader(
            "Preview dan pilihan Warehouse"
        )

        st.info(
            "Pilih Warehouse pada setiap baris. "
            "Nomor Faktur Pajak tetap kosong."
        )

        edited_df = st.data_editor(

            df,

            use_container_width=True,

            hide_index=True,

            num_rows="fixed",

            column_config={

                "No":
                    st.column_config.NumberColumn(
                        "No",
                        disabled=True
                    ),

                "Tanggal Invoice":
                    st.column_config.TextColumn(
                        "Tanggal Invoice",
                        disabled=True
                    ),

                "Nama Store":
                    st.column_config.TextColumn(
                        "Nama Store",
                        disabled=False
                    ),

                "Brand":
                    st.column_config.SelectboxColumn(
                        "Brand",
                        options=BRANDS,
                        required=False
                    ),

                "Nomor Invoice":
                    st.column_config.TextColumn(
                        "Nomor Invoice",
                        disabled=True
                    ),

                "Nominal":
                    st.column_config.NumberColumn(
                        "Nominal",
                        format="Rp %d",
                        disabled=False
                    ),

                "Warehouse":
                    st.column_config.SelectboxColumn(
                        "Warehouse",
                        options=WAREHOUSES,
                        required=False
                    ),

                "Nomor Faktur Pajak":
                    st.column_config.TextColumn(
                        "Nomor Faktur Pajak",
                        required=False
                    ),
            },

            key="invoice_editor"
        )

        st.session_state.invoice_data = (
            edited_df
        )

        # ====================================================
        # EXPORT EXCEL
        # ====================================================

        st.divider()

        st.subheader(
            "📊 Export Excel"
        )

        output = io.BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            edited_df.to_excel(
                writer,
                index=False,
                sheet_name="Invoice"
            )

            worksheet = writer.sheets[
                "Invoice"
            ]

            # Header
            for cell in worksheet[1]:

                cell.font = cell.font.copy(
                    bold=True
                )

            # Lebar kolom
            widths = {

                "A": 8,

                "B": 18,

                "C": 35,

                "D": 20,

                "E": 25,

                "F": 18,

                "G": 20,

                "H": 25,
            }

            for column, width in widths.items():

                worksheet.column_dimensions[
                    column
                ].width = width

            # Format nominal
            for cell in worksheet["F"][1:]:

                cell.number_format = '#,##0'

        output.seek(0)

        st.download_button(

            label="⬇️ Download Excel",

            data=output.getvalue(),

            file_name="hasil_invoice.xlsx",

            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),

            type="primary"
        )

    # ========================================================
    # DEBUG
    # ========================================================

    debug_df = st.session_state.get(
        "debug_data",
        pd.DataFrame()
    )

    if not debug_df.empty:

        with st.expander(
            "🔎 Detail pembacaan PDF"
        ):

            st.dataframe(
                debug_df,
                use_container_width=True,
                hide_index=True
            )
