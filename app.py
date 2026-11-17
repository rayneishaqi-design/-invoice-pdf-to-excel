import io
import re
import zipfile
from pathlib import Path
from datetime import datetime

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter


st.set_page_config(
    page_title="Invoice PDF to Excel",
    page_icon="📊",
    layout="wide"
)

WAREHOUSES = [
    "Medan", "Padang", "Palembang", "Yogyakarta", "Surabaya",
    "Malang", "Makassar", "Banjarmasin", "Bekasi", "Jakarta"
]

OUTPUT_COLUMNS = [
    "No", "Tanggal Invoice", "Nama Store", "Brand",
    "Nomor Invoice", "Nominal", "Warehouse", "Nomor Faktur Pajak"
]

BRAND_RULES = [
    ("Focallure", ("FOCALLURE",)),
    ("Pinkflash", ("PINKFLASH", "PINK FLASH")),
    ("Cave", ("CAVE",)),
    ("Teratu", ("TERATU",)),
    ("Philips", ("PHILIPS",)),
]


def tidy(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def parse_amount(value):
    """Ubah angka format Indonesia, misalnya 151.740,00 menjadi integer 151740."""
    value = tidy(value).replace("Rp", "").replace(" ", "")
    value = value.replace(".", "").replace(",", ".")
    try:
        return int(round(float(value)))
    except (TypeError, ValueError):
        return None


def parse_date(value):
    value = tidy(value)
    month_map = {
        "jan": 1, "feb": 2, "mar": 3, "apr": 4, "mei": 5, "may": 5,
        "jun": 6, "jul": 7, "agu": 8, "aug": 8, "sep": 9, "okt": 10,
        "oct": 10, "nov": 11, "des": 12, "dec": 12
    }
    match = re.search(r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", value)
    if not match:
        return value
    day, month_name, year = match.groups()
    month = month_map.get(month_name[:3].lower())
    if not month:
        return value
    return datetime(int(year), month, int(day))


def detect_brand(line):
    upper = line.upper()
    for brand, keywords in BRAND_RULES:
        if any(keyword in upper for keyword in keywords):
            return brand
    return None


def extract_pdf(pdf_bytes, filename):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        pages_text = [page.extract_text() or "" for page in pdf.pages]
    text = "\n".join(pages_text)
    lines = [tidy(line) for line in text.splitlines() if tidy(line)]

    invoice_match = re.search(
        r"\bSI[.-]\d{4}[.-]\d{2}[.-]\d{5}\b", text, re.I
    )
    invoice_number = invoice_match.group(0).replace("-", ".") if invoice_match else Path(filename).stem

    date_match = re.search(
        r"\b(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|Mei|May|Jun|Jul|Agu|Aug|Sep|Okt|Oct|Nov|Des|Dec)[a-z]*\s+\d{4})\b",
        text, re.I
    )
    invoice_date = parse_date(date_match.group(1)) if date_match else ""

    # Nama pelanggan berada pada baris setelah baris "Kepada : Tanggal Nomor".
    store = ""
    for i, line in enumerate(lines):
        if line.lower().startswith("kepada"):
            for candidate in lines[i + 1:i + 4]:
                if re.search(r"\bSI[.-]\d{4}[.-]\d{2}[.-]\d{5}\b", candidate, re.I):
                    continue
                if re.search(r"syarat pembayaran|tanggal jatuh tempo|net\s*\d+", candidate, re.I):
                    continue
                # Baris alamat biasanya diawali Jl./Jalan, bukan nama pelanggan.
                if re.match(r"^(Jl\.?|Jalan|Jln\.?)\b", candidate, re.I):
                    continue
                if candidate.lower() not in ("tanggal nomor",):
                    store = candidate
                    break
            break

    # Umumnya nama toko adalah bagian setelah tanda hubung pada "Bp. Nama - Nama Toko".
    if " - " in store:
        store = store.split(" - ", 1)[1].strip()

    # TOTAL invoice adalah angka pada baris yang labelnya tepat "Total".
    total_match = re.search(r"(?im)^\s*Total\s+([\d.,]+)\s*$", text)
    invoice_total = parse_amount(total_match.group(1)) if total_match else None

    # Baca jumlah per brand dari kolom Total Harga pada baris item.
    # Jika PDF hanya memiliki satu brand, gunakan angka TOTAL invoice sebagai nominal.
    brand_amounts = {}
    for line in lines:
        brand = detect_brand(line)
        if not brand:
            continue

        # Baris produk invoice berakhir dengan nilai Total Harga.
        # Ambil angka terakhir yang berdiri sendiri pada baris produk.
        amounts = re.findall(r"(?<![\w])(\d{1,3}(?:\.\d{3})*(?:,\d{2})?)(?![\w])", line)
        if not amounts:
            continue
        line_amount = parse_amount(amounts[-1])
        if line_amount is not None:
            brand_amounts[brand] = brand_amounts.get(brand, 0) + line_amount

    if len(brand_amounts) == 1:
        brand, _ = next(iter(brand_amounts.items()))
        brand_amounts = {brand: invoice_total if invoice_total is not None else next(iter(brand_amounts.values()))}
    elif not brand_amounts:
        brand_amounts = {"Perlu diperiksa": invoice_total}

    records = []
    for brand, amount in brand_amounts.items():
        records.append({
            "Tanggal Invoice": invoice_date,
            "Nama Store": store,
            "Brand": brand,
            "Nomor Invoice": invoice_number,
            "Nominal": amount,
            "Warehouse": "",
            "Nomor Faktur Pajak": ""
        })

    return records, {
        "filename": filename,
        "invoice_total": invoice_total,
        "brand_count": len(brand_amounts),
        "store": store,
        "invoice_number": invoice_number
    }


def build_excel(dataframe):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe[OUTPUT_COLUMNS].to_excel(
            writer, index=False, sheet_name="Rekap Invoice"
        )
    output.seek(0)

    workbook = load_workbook(output)
    sheet = workbook["Rekap Invoice"]
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.sheet_view.showGridLines = False

    header_fill = PatternFill("solid", fgColor="17365D")
    header_font = Font(name="Arial", bold=True, color="FFFFFF")
    thin = Side(style="thin", color="D9E2F3")

    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    widths = [8, 18, 42, 18, 25, 20, 20, 25]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width

    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = Border(bottom=thin)
        row[1].number_format = "dd/mm/yyyy"
        row[5].number_format = '"Rp" #,##0'
        row[7].number_format = "@"

    if sheet.max_row >= 2:
        validation = DataValidation(
            type="list",
            formula1='"' + ",".join(WAREHOUSES) + '"',
            allow_blank=True
        )
        validation.error = "Pilih warehouse dari daftar."
        validation.errorTitle = "Warehouse tidak valid"
        validation.prompt = "Pilih warehouse untuk baris invoice ini."
        validation.promptTitle = "Pilih Warehouse"
        sheet.add_data_validation(validation)
        validation.add(f"G2:G{sheet.max_row}")

    final = io.BytesIO()
    workbook.save(final)
    final.seek(0)
    return final.getvalue()


st.title("📊 Invoice PDF to Excel")
st.write(
    "Aplikasi baru untuk menggabungkan beberapa invoice PDF dari satu file ZIP "
    "menjadi satu file Excel."
)

st.markdown(
    """
    **Cara kerja:** unggah ZIP → aplikasi mengekstrak PDF → membaca data invoice →
    pilih Warehouse → unduh Excel.
    """
)

st.info(
    "Nominal diambil dari angka pada baris **Total** invoice (contoh: 151.740,00). "
    "Nomor Faktur Pajak dibiarkan kosong. Warehouse dipilih manual."
)

uploaded_zip = st.file_uploader(
    "Upload file ZIP yang berisi PDF invoice",
    type=["zip"],
    help="ZIP boleh berisi banyak PDF invoice."
)

if uploaded_zip:
    try:
        with zipfile.ZipFile(uploaded_zip) as archive:
            pdf_files = [
                name for name in archive.namelist()
                if name.lower().endswith(".pdf")
                and not name.startswith("__MACOSX/")
            ]

            if not pdf_files:
                st.error("Tidak ada PDF yang ditemukan di dalam ZIP.")
                st.stop()

            all_records = []
            read_errors = []
            audit = []

            progress = st.progress(0, text="Membaca invoice...")
            for index, filename in enumerate(pdf_files):
                try:
                    records, info = extract_pdf(archive.read(filename), filename)
                    all_records.extend(records)
                    audit.append(info)
                except Exception as exc:
                    read_errors.append(f"{filename}: {exc}")
                progress.progress(
                    (index + 1) / len(pdf_files),
                    text=f"Membaca {index + 1} dari {len(pdf_files)} PDF"
                )

        if read_errors:
            st.warning("Ada PDF yang gagal dibaca:")
            for message in read_errors:
                st.write(f"- {message}")

        if not all_records:
            st.error("Data invoice belum berhasil dibaca. Periksa PDF yang diunggah.")
            st.stop()

        dataframe = pd.DataFrame(all_records)
        dataframe.insert(0, "No", range(1, len(dataframe) + 1))

        st.success(f"Berhasil membaca {len(pdf_files)} PDF dan menghasilkan {len(dataframe)} baris rekap.")

        with st.expander("Rincian pembacaan PDF"):
            audit_df = pd.DataFrame(audit)
            st.dataframe(audit_df, use_container_width=True, hide_index=True)

        st.subheader("Preview dan pilihan Warehouse")
        st.caption("Pilih Warehouse pada setiap baris. Nomor Faktur Pajak tetap kosong.")
        edited = st.data_editor(
            dataframe[OUTPUT_COLUMNS],
            use_container_width=True,
            hide_index=True,
            disabled=[column for column in OUTPUT_COLUMNS if column != "Warehouse"],
            column_config={
                "Warehouse": st.column_config.SelectboxColumn(
                    "Warehouse",
                    options=WAREHOUSES,
                    required=False
                ),
                "Nominal": st.column_config.NumberColumn(
                    "Nominal",
                    format="Rp %d"
                ),
                "Tanggal Invoice": st.column_config.DateColumn(
                    "Tanggal Invoice",
                    format="DD/MM/YYYY"
                )
            },
            key="invoice_editor"
        )

        if st.button("Buat File Excel", type="primary", use_container_width=True):
            excel_bytes = build_excel(edited)
            st.download_button(
                "⬇️ Download Rekap Invoice.xlsx",
                data=excel_bytes,
                file_name="Rekap_Invoice.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

    except zipfile.BadZipFile:
        st.error("File yang diunggah bukan ZIP yang valid.")

st.divider()
st.caption("Invoice PDF to Excel | Aplikasi terpisah dari sistem Surat Jalan OXAR GLOBAL")
