import win32com.client as win32
import os
import re
import logging
import yaml

# Konfigurasi logging
logging.basicConfig(
    filename="script_log.txt",
    level=logging.ERROR,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# Fungsi untuk memastikan Excel ditutup


def close_excel_instances():
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Quit()
    except Exception as e:
        logging.error(f"Failed to close Excel instances: {e}")

# Fungsi untuk membaca file teks


def read_text_file(file_path):
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            return [line.strip() for line in f.readlines()]
    except Exception as e:
        logging.error(f"Kesalahan saat membaca file {file_path}: {e}")
        raise

# Fungsi untuk menulis file teks


def write_text_file(file_path, data):
    try:
        with open(file_path, "w", encoding="utf-8") as f:
            f.write("\n".join(data))
    except Exception as e:
        logging.error(f"Kesalahan saat menulis file {file_path}: {e}")
        raise

# Fungsi untuk membaca konfigurasi dari file YAML


def read_config(config_file):
    try:
        with open(config_file, "r", encoding="utf-8") as f:
            return yaml.safe_load(f)
    except Exception as e:
        logging.error(f"Kesalahan saat membaca file konfigurasi {config_file}: {e}")
        raise

# Fungsi untuk menghapus karakter tak terlihat di awal nilai


def clean_leading_whitespace(value):
    if value:
        return re.sub(r'^[\s\u00A0]+', '', value)
    return value

# Fungsi untuk mengonversi file .xls ke .xlsx


def convert_xls_to_xlsx(xls_file, xlsx_file):
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        workbook = excel.Workbooks.Open(xls_file)
        workbook.SaveAs(xlsx_file, FileFormat=51)  # FileFormat 51 = .xlsx
        workbook.Close()
        excel.Quit()
    except Exception as e:
        logging.error(f"Kesalahan saat mengonversi file: {xls_file}. Error: {e}")
        raise

# Fungsi untuk memperbarui kolom BL Akhir


def edit_bl_akhir(sheet, header_row, bl_akhir_value):
    row_count = sheet.UsedRange.Rows.Count
    for row in range(header_row + 1, row_count + 1):
        sheet.Cells(row, 11).Value = bl_akhir_value  # Kolom BL Akhir (kolom ke-11)

# Fungsi untuk memperbarui kolom LBR


def edit_lbr(sheet, header_row):
    row_count = sheet.UsedRange.Rows.Count
    for row in range(header_row + 1, row_count + 1):
        bl_awal_value = sheet.Cells(row, 10).Value  # Kolom BL Awal (kolom ke-10)
        if not bl_awal_value:
            lbr_value = 1
        else:
            comma_count = str(bl_awal_value).count(",")
            lbr_value = comma_count + 1
        sheet.Cells(row, 12).Value = f"({lbr_value}"  # Kolom LBR (kolom ke-12)

# Fungsi untuk memperbarui kolom BL Awal berdasarkan kategori LBR


def edit_bl_awal(sheet, header_row, lbr_categories):
    row_count = sheet.UsedRange.Rows.Count
    for row in range(header_row + 1, row_count + 1):
        lbr_value = sheet.Cells(row, 12).Value
        if lbr_value in lbr_categories:
            sheet.Cells(row, 10).Value = lbr_categories[lbr_value]

# Fungsi untuk menghapus baris dengan nilai RPTAG 0


def delete_rows_with_zero_rptag(sheet, header_row):
    row_count = sheet.UsedRange.Rows.Count
    for row in range(row_count, header_row, -1):  # Iterasi dari baris terakhir ke baris pertama
        rptag_value = sheet.Cells(row, 13).Value
        if rptag_value is None or rptag_value == 0:  # Jika nilai RPTAG 0 atau None
            sheet.Rows(row).Delete()  # Hapus baris

# Fungsi untuk memperbarui kolom RPBK


def edit_rpbk(sheet, header_row):
    row_count = sheet.UsedRange.Rows.Count
    for row in range(header_row + 1, row_count + 1):
        rpbk_value = sheet.Cells(row, 14).Value
        lbr_value = sheet.Cells(row, 12).Value
        if lbr_value:
            lbr_value = lbr_value.strip()
            if rpbk_value is None or not isinstance(rpbk_value, (int, float)):
                rpbk_value = 0
            if lbr_value == "(1":
                pass
            elif lbr_value == "(2":
                sheet.Cells(row, 14).Value = rpbk_value * 2
            elif lbr_value == "(3":
                sheet.Cells(row, 14).Value = rpbk_value * 5 if rpbk_value == 3000 else rpbk_value * 3

# Fungsi untuk memperbarui kolom RPTAG


def edit_rptag(sheet, header_row, rptag_config):
    row_count = sheet.UsedRange.Rows.Count
    file_name = os.path.basename(sheet.Parent.Name).lower()
    rptag_tambahan = rptag_config.get(file_name, 0)

    for row in range(header_row + 1, row_count + 1):
        rptag_value = sheet.Cells(row, 13).Value  # Kolom RPTAG (kolom ke-13)
        if rptag_value is not None:
            sheet.Cells(row, 13).Value = rptag_value + rptag_tambahan

# Fungsi untuk memproses Folder1


def process_folder1(folder1, output_directory):
    for file_name in os.listdir(folder1):
        if file_name.endswith(".xls"):
            xls_file_path = os.path.join(folder1, file_name)
            xlsx_file_path = os.path.join(output_directory, f"{os.path.splitext(file_name)[0]}.xlsx")
            convert_xls_to_xlsx(xls_file_path, xlsx_file_path)

            excel = win32.gencache.EnsureDispatch('Excel.Application')
            workbook = excel.Workbooks.Open(xlsx_file_path)
            sheet = workbook.Worksheets(1)

            row_count = sheet.UsedRange.Rows.Count
            header_row = next((i for i in range(1, row_count + 1) if sheet.Cells(i, 1).Value == "NO"), None)
            if header_row is None:
                raise ValueError(f"Header 'NO' tidak ditemukan di file {file_name}.")

            bln_tagihan_values = []
            tagihan_values = []

            for row in range(header_row + 1, row_count + 1):
                bln_tagihan = sheet.Cells(row, 4).Value  # Kolom BLN TAGIHAN
                tagihan = sheet.Cells(row, 5).Value  # Kolom TAGIHAN
                bln_tagihan_values.append(clean_leading_whitespace(str(bln_tagihan)) if bln_tagihan else "")
                tagihan_values.append(clean_leading_whitespace(str(tagihan)) if tagihan else "")

            bln_tagihan_txt = os.path.join(output_directory, f"{os.path.splitext(file_name)[0]}_BLNTAGIHAN.txt")
            tagihan_txt = os.path.join(output_directory, f"{os.path.splitext(file_name)[0]}_TAGIHAN.txt")

            write_text_file(bln_tagihan_txt, bln_tagihan_values)
            write_text_file(tagihan_txt, tagihan_values)

            workbook.Close()
            excel.Quit()

# Fungsi untuk memproses Folder2


def process_folder2(folder2, output_directory, config):
    bl_akhir_value = config['BL_AKHIR']
    lbr_categories = {
        "(1": config['LBR_1'],
        "(2": config['LBR_2'],
        "(3": config['LBR_3']
    }
    rptag_config = {k.lower(): v for k, v in config['RPTAG'].items()}

    for file_name in os.listdir(folder2):
        if file_name.endswith(".xlsx"):
            file_path = os.path.join(folder2, file_name)
            output_file = os.path.join(output_directory, f"{os.path.splitext(file_name)[0]}_cleaned.xlsx")

            bln_tagihan_txt = os.path.join(output_directory, f"{os.path.splitext(file_name)[0]}_BLNTAGIHAN.txt")
            tagihan_txt = os.path.join(output_directory, f"{os.path.splitext(file_name)[0]}_TAGIHAN.txt")

            bln_tagihan_data = read_text_file(bln_tagihan_txt)
            tagihan_data = read_text_file(tagihan_txt)

            excel = win32.gencache.EnsureDispatch('Excel.Application')
            workbook = excel.Workbooks.Open(file_path)
            sheet = workbook.Worksheets(1)

            row_count = sheet.UsedRange.Rows.Count
            header_row = next((i for i in range(1, row_count + 1) if sheet.Cells(i, 1).Value == "NO"), None)
            if header_row is None:
                raise ValueError(f"Header 'NO' tidak ditemukan di file {file_name}.")

            for i, value in enumerate(bln_tagihan_data):
                sheet.Cells(header_row + 1 + i, 10).Value = value

            for i, value in enumerate(tagihan_data):
                clean_value = value.replace(".", "")
                sheet.Cells(header_row + 1 + i, 13).Value = float(clean_value)

            delete_rows_with_zero_rptag(sheet, header_row)
            edit_bl_akhir(sheet, header_row, bl_akhir_value)
            edit_lbr(sheet, header_row)
            edit_bl_awal(sheet, header_row, lbr_categories)
            edit_rpbk(sheet, header_row)
            edit_rptag(sheet, header_row, rptag_config)

            workbook.SaveAs(output_file, FileFormat=51)
            workbook.Close()
            excel.Quit()

# Fungsi untuk menggabungkan file hasil


def merge_cleaned_files(output_directory):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    merged_file = os.path.join(output_directory, "Merged_Data.xlsx")
    workbook_merged = excel.Workbooks.Add()
    sheet_merged = workbook_merged.Worksheets(1)

    current_row = 1
    for file_name in os.listdir(output_directory):
        if file_name.endswith("_cleaned.xlsx"):
            file_path = os.path.join(output_directory, file_name)
            workbook_source = excel.Workbooks.Open(file_path)
            sheet_source = workbook_source.Worksheets(1)

            row_count = sheet_source.UsedRange.Rows.Count
            col_count = sheet_source.UsedRange.Columns.Count

            data_range = sheet_source.Range(sheet_source.Cells(1, 1), sheet_source.Cells(row_count, col_count))
            target_range = sheet_merged.Cells(current_row, 1)
            data_range.Copy(target_range)

            current_row += row_count
            workbook_source.Close(False)

    workbook_merged.SaveAs(merged_file, FileFormat=51)
    workbook_merged.Close()
    excel.Quit()
    print(f"File gabungan berhasil dibuat: {merged_file}")


# Eksekusi Skrip
if __name__ == "__main__":
    try:
        source_directory = r"C:\Users\Administrator\Desktop"
        folder1 = os.path.join(source_directory, "Folder1")
        folder2 = os.path.join(source_directory, "Folder2")
        output_directory = os.path.join(source_directory, "out")

        config_file = os.path.join(source_directory, "input_config.yaml")
        config = read_config(config_file)

        if not os.path.exists(output_directory):
            os.makedirs(output_directory)

        process_folder1(folder1, output_directory)
        process_folder2(folder2, output_directory, config)
        merge_cleaned_files(output_directory)
        print("Proses selesai. Silakan cek folder output.")
    except Exception as e:
        logging.critical(f"Kesalahan dalam eksekusi skrip: {e}")
        raise
