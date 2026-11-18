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
    "Upload ZIP berisi invoice PDF. Sistem akan membaca seluruh halaman "
    "dan menggabungkan hasilnya menjadi satu file Excel."
)

# ============================================================
# MASTER DATA
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
]

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


def normalize_number(value):
    """
    Mengubah angka invoice menjadi numeric.
    Contoh:
    Rp 3.003.880,00 -> 3003880
    3,003,880.00    -> 3003880
    """

    if value is None:
        return None

    value = str(value).strip()

    # Hapus Rp dan karakter non angka kecuali separator
    value = re.sub(r"(?i)rp", "", value)
    value = value.strip()

    if not value:
        return None

    # Format Indonesia: 151.740,00
    if "," in value and "." in value:
        if value.rfind(",") > value.rfind("."):
            value = value.replace(".", "")
            value = value.replace(",", ".")
        else:
            value = value.replace(",", "")

    # 151.740,00
    elif "," in value:
        decimal_part = value.split(",")[-1]

        if len(decimal_part) <= 2:
            value = value.replace(".", "")
            value = value.replace(",", ".")
        else:
            value = value.replace(",", "")

    # 151.740
    elif "." in value:
        parts = value.split(".")

        if len(parts[-1]) == 2:
            value = value.replace(".", "")
        else:
            value = value.replace(".", "")

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
    """
    Mencari brand berdasarkan isi PDF.
    Jika tidak ditemukan, coba dari nama file.
    """

    combined = f"{text}\n{filename}".lower()

    # Prioritaskan brand yang lebih spesifik
    sorted_brands = sorted(
        BRANDS,
        key=lambda x: len(x),
        reverse=True
    )

    for brand in sorted_brands:
        if brand.lower() in combined:
            return brand

    return ""


# ============================================================
# DETECT INVOICE NUMBER
# ============================================================

def detect_invoice_number(text):
    patterns = [
        r"(?:invoice\s*(?:no|number|#)?|no\.?\s*invoice|nomor\s*invoice)"
        r"\s*[:#\-]?\s*([A-Z0-9./\-]+)",

        r"\b(SI\.[0-9]{4}\.[0-9]{2}\.[0-9]+)\b",

        r"\b(INV[./\-]?[A-Z0-9./\-]+)\b",
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE
        )

        if match:
            return match.group(1).strip()

    return ""


# ============================================================
# DETECT DATE
# ============================================================

def detect_invoice_date(text):
    patterns = [
        # 17/11/2026
        r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{4})\b",

        # 17-11-2026
        r"\b(\d{1,2}-\d{1,2}-\d{4})\b",

        # 17 Nov 2026
        r"\b(\d{1,2}\s+"
        r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
        r"\s+\d{4})\b",
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE
        )

        if match:
            value = match.group(1)

            # dd/mm/yyyy
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
    """
    Mencari nama toko berdasarkan label yang umum ditemukan
    pada invoice.
    """

    patterns = [
        r"(?:customer|pelanggan|customer name|nama pelanggan)"
        r"\s*[:\-]\s*(.+)",

        r"(?:bill\s*to|ship\s*to)"
        r"\s*[:\-]\s*(.+)",

        r"(?:store|toko|outlet)"
        r"\s*[:\-]\s*(.+)",
    ]

    lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE
        )

        if match:
            result = match.group(1).strip()

            # Ambil hanya sampai akhir baris
            result = result.split("\n")[0].strip()

            if len(result) > 1:
                return result

    # Fallback:
    # cari baris yang mengandung nama toko umum
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
    ]

    for line in lines[:30]:

        lower = line.lower()

        if any(word in lower for word in blacklist):
            continue

        if len(line) >= 4 and len(line) <= 80:
            # Hindari baris yang hanya angka
            if not re.fullmatch(r"[\d\s.,:/\-]+", line):
                return line

    return ""


# ============================================================
# DETECT TOTAL
# ============================================================

def detect_total(text):
    """
    Mengambil nilai TOTAL invoice.

    Prioritas:
    1. TOTAL
    2. GRAND TOTAL
    3. TOTAL INVOICE
    """

    lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]

    # Cari dari baris yang mengandung TOTAL
    for i, line in enumerate(lines):

        if re.search(
            r"\b(grand\s+total|total\s+invoice|total)\b",
            line,
            flags=re.IGNORECASE
        ):

            # Cari angka di baris yang sama
            numbers = re.findall(
                r"(?:Rp\s*)?[\d.,]+",
                line,
                flags=re.IGNORECASE
            )

            if numbers:
                candidates = [
                    normalize_number(x)
                    for x in numbers
                ]

                candidates = [
                    x for x in candidates
                    if x is not None
                ]

                if candidates:
                    return candidates[-1]

            # Kalau angka ada di baris berikutnya
            for next_line in lines[i + 1:i + 3]:

                numbers = re.findall(
                    r"(?:Rp\s*)?[\d.,]+",
                    next_line,
                    flags=re.IGNORECASE
                )

                if numbers:

                    candidates = [
                        normalize_number(x)
                        for x in numbers
                    ]

                    candidates = [
                        x for x in candidates
                        if x is not None
                    ]

                    if candidates:
                        return candidates[-1]

    # Fallback: cari angka terbesar yang diawali Rp
    rp_values = re.findall(
        r"(?:Rp\s*)[\d.,]+",
        text,
        flags=re.IGNORECASE
    )

    values = []

    for value in rp_values:
        number = normalize_number(value)

        if number is not None:
            values.append(number)

    if values:
        return max(values)

    return None


# ============================================================
# EXTRACT ONE PDF
# ============================================================

def extract_pdf(pdf_bytes, filename):
    """
    Membaca SELURUH halaman PDF.

    Tidak hanya halaman pertama.
    """

    all_text = []
    page_count = 0

    try:

        with pdfplumber.open(
            io.BytesIO(pdf_bytes)
        ) as pdf:

            page_count = len(pdf.pages)

            for page in pdf.pages:

                text = page.extract_text(
                    x_tolerance=2,
                    y_tolerance=2
                )

                if text:
                    all_text.append(text)

    except Exception as e:

        return {
            "success": False,
            "error": str(e),
            "pages": 0
        }

    combined_text = "\n".join(all_text)
    combined_text = clean_text(combined_text)

    return {
        "success": True,
        "pages": page_count,
        "text": combined_text,
        "filename": filename
    }


# ============================================================
# PROCESS ZIP
# ============================================================

def process_zip(uploaded_file):

    results = []
    debug_info = []

    try:

        zip_bytes = uploaded_file.read()

        with zipfile.ZipFile(
            io.BytesIO(zip_bytes),
            "r"
        ) as z:

            pdf_files = [
                name
                for name in z.namelist()
                if name.lower().endswith(".pdf")
                and not name.endswith("/")
            ]

            if not pdf_files:
                st.error(
                    "ZIP tidak berisi file PDF."
                )
                return pd.DataFrame()

            progress = st.progress(0)

            for index, pdf_name in enumerate(pdf_files):

                try:

                    pdf_bytes = z.read(pdf_name)

                    extracted = extract_pdf(
                        pdf_bytes,
                        os.path.basename(pdf_name)
                    )

                    if not extracted["success"]:

                        debug_info.append({
                            "File": pdf_name,
                            "Status": "Gagal",
                            "Halaman": 0,
                            "Error": extracted["error"]
                        })

                        continue

                    text = extracted["text"]

                    brand = detect_brand(
                        text,
                        pdf_name
                    )

                    invoice_number = detect_invoice_number(
                        text
                    )

                    invoice_date = detect_invoice_date(
                        text
                    )

                    store = detect_store(
                        text
                    )

                    total = detect_total(
                        text
                    )

                    results.append({
                        "No": len(results) + 1,
                        "Tanggal Invoice": invoice_date,
                        "Nama Store": store,
                        "Brand": brand,
                        "Nomor Invoice": invoice_number,
                        "Nominal": total,
                        "Warehouse": "",
                        "Nomor Faktur Pajak": ""
                    })

                    debug_info.append({
                        "File": pdf_name,
                        "Status": "OK",
                        "Halaman": extracted["pages"],
                        "Brand": brand,
                        "Store": store,
                        "Invoice": invoice_number,
                        "Total": total
                    })

                except Exception as e:

                    debug_info.append({
                        "File": pdf_name,
                        "Status": "Gagal",
                        "Halaman": 0,
                        "Error": str(e)
                    })

                progress.progress(
                    (index + 1) / len(pdf_files)
                )

    except zipfile.BadZipFile:

        st.error(
            "File yang diupload bukan ZIP yang valid."
        )

        return pd.DataFrame()

    return pd.DataFrame(results), pd.DataFrame(debug_info)


# ============================================================
# UPLOAD
# ============================================================

uploaded_zip = st.file_uploader(
    "Upload file ZIP berisi invoice PDF",
    type=["zip"]
)

if uploaded_zip:

    if "invoice_data" not in st.session_state:

        st.session_state.invoice_data = pd.DataFrame()

    if st.button(
        "🔍 Proses ZIP",
        type="primary"
    ):

        with st.spinner(
            "Membaca seluruh PDF dan seluruh halaman..."
        ):

            result = process_zip(
                uploaded_zip
            )

            if isinstance(result, tuple):

                df, debug_df = result

                st.session_state.invoice_data = df
                st.session_state.debug_data = debug_df

    # ========================================================
    # PREVIEW
    # ========================================================

    df = st.session_state.get(
        "invoice_data",
        pd.DataFrame()
    )

    if not df.empty:

        st.success(
            f"Berhasil membaca {len(df)} invoice."
        )

        st.subheader(
            "Preview dan pilihan Warehouse"
        )

        st.info(
            "Pilih Warehouse pada setiap baris. "
            "Nomor Faktur Pajak dibiarkan kosong."
        )

        # ----------------------------------------------------
        # DATA EDITOR
        # ----------------------------------------------------

        edited_df = st.data_editor(
            df,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",

            column_config={

                "No": st.column_config.NumberColumn(
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

        st.session_state.invoice_data = edited_df

        # ====================================================
        # DOWNLOAD EXCEL
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

            workbook = writer.book
            worksheet = writer.sheets["Invoice"]

            # Format header
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
