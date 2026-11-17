# Invoice PDF to Excel

Aplikasi Streamlit baru untuk membaca beberapa invoice PDF dari satu file ZIP dan menggabungkan datanya ke satu file Excel.

## Fitur
- Upload ZIP berisi banyak PDF invoice.
- Ekstraksi tanggal invoice, nama toko, brand, nomor invoice, dan nominal.
- Nominal menggunakan angka pada baris **Total** invoice.
- Pilihan Warehouse manual: Medan, Padang, Palembang, Yogyakarta, Surabaya, Malang, Makassar, Banjarmasin, Bekasi, Jakarta.
- Nomor Faktur Pajak dibiarkan kosong.
- Hasil akhir satu file Excel dengan dropdown Warehouse.

## Deploy ke GitHub dan Streamlit Community Cloud
1. Buat repository GitHub baru, misalnya `invoice-pdf-to-excel`.
2. Upload `app.py`, `requirements.txt`, dan `README.md` ke root repository.
3. Commit perubahan.
4. Buka https://share.streamlit.io/ lalu pilih **New app**.
5. Hubungkan repository baru tersebut.
6. Pilih branch `main` dan main file `app.py`.
7. Klik **Deploy**.

## Menjalankan di komputer
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Catatan
Aplikasi membaca teks PDF. PDF hasil scan/foto yang tidak memiliki text layer mungkin perlu OCR agar dapat dibaca dengan benar.
