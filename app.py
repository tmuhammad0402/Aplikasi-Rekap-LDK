import streamlit as st
import imaplib
import email
from email.header import decode_header
import os
import shutil
import json
import re
import glob
import pandas as pd
import numpy as np
import zipfile
import copy
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

# Konfigurasi Halaman Streamlit
st.set_page_config(page_title="Aplikasi Rekap LDK", layout="wide")

# ============================================================
# PROGRAM 1: YAHOO IMAP ATTACHMENT DOWNLOADER
# ============================================================

def clean_filename(filename):
    if not filename:
        return "untitled_attachment"
    return re.sub(r'[\\/*?:"<>|\n\r]', "", filename).strip()

def decode_mime_header(header_value):
    if not header_value:
        return ""
    decoded_parts = decode_header(header_value)
    result = ""
    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            try:
                result += part.decode(encoding or "utf-8", errors="ignore")
            except:
                result += part.decode("latin1", errors="ignore")
        else:
            result += str(part)
    return result

def do_imap_search(mail_conn, query_string):
    words = query_string.split()
    search_criteria = " ".join([f'SUBJECT "{word}"' for word in words])
    status, messages = mail_conn.search(None, search_criteria)
    if status == "OK" and messages[0]:
        return messages[0].split()
    return []

def extract_date_key(text):
    pattern = r'(?i)(.*?)\s+(\d{1,2})\s+([a-z]+)\s+(\d{4})'
    match = re.search(pattern, text)
    if match:
        return f"{int(match.group(2))}_{match.group(3).strip().upper()}_{match.group(4).strip()}"
    return None

def merge_downloaded_excel(download_dir):
    excel_files = glob.glob(os.path.join(download_dir, "*.xlsx"))
    valid_files = [f for f in excel_files if "(GABUNGAN)" not in f]
    
    if len(valid_files) < 2:
        return None

    data_dfs = []
    header_df = None
    lots_col_idx = 6 
    bs_col_idx = None 
    
    days = []
    months = set()
    years = set()
    prefixes = []
    summary_rows = []

    pattern = r'(?i)(.*?)\s+(\d{1,2})\s+([a-z]+)\s+(\d{4})'

    for file_path in valid_files:
        file_name = os.path.basename(file_path)
        match = re.search(pattern, file_name)
        if match:
            prefixes.append(match.group(1).strip())
            day_val = int(match.group(2))
            days.append(day_val)
            month_val = match.group(3).strip().upper()
            months.add(month_val)
            year_val = match.group(4).strip()
            years.add(year_val)
            date_str = f"{day_val:02d}-{month_val[:3].capitalize()}-{year_val[-2:]}"
        else:
            day_val = 0
            date_str = "Unknown Date"

        try:
            xls = pd.ExcelFile(file_path)
            bursa_sheet = [sheet for sheet in xls.sheet_names if "BURSA" in sheet.upper()]
            
            if bursa_sheet:
                df = pd.read_excel(file_path, sheet_name=bursa_sheet[0], header=None)
                curr_header_idx = 0
                for i in range(min(15, len(df))):
                    if df.iloc[i].astype(str).str.contains('Trade No', case=False).any():
                        curr_header_idx = i
                        break
                        
                if header_df is None:
                    header_df = df.iloc[:curr_header_idx+1].copy()
                    header_row = df.iloc[curr_header_idx]
                    for idx, val in header_row.items():
                        v_str = str(val).lower().strip()
                        if v_str == 'lots':
                            lots_col_idx = idx
                        elif v_str in ['b/s', 'b / s', 'type', 'side', 'action', 'beli/jual', 'jual/beli']:
                            bs_col_idx = idx
                            
                data = df.iloc[curr_header_idx+1:].copy()
                data = data[data[0].notna()]
                data = data[~data[0].astype(str).str.upper().str.contains('TOTAL')]
                
                temp_data = data.copy()
                temp_data[lots_col_idx] = pd.to_numeric(temp_data[lots_col_idx], errors='coerce').fillna(0)
                
                total_trades = len(temp_data)
                total_lots_day = temp_data[lots_col_idx].sum()
                buy_lots = 0
                sell_lots = 0
                
                if bs_col_idx is not None:
                    temp_data['bs_clean'] = temp_data[bs_col_idx].astype(str).str.upper().str.strip()
                    buy_mask = temp_data['bs_clean'].str.startswith('B')
                    sell_mask = temp_data['bs_clean'].str.startswith('S')
                    buy_lots = temp_data.loc[buy_mask, lots_col_idx].sum()
                    sell_lots = temp_data.loc[sell_mask, lots_col_idx].sum()
                
                summary_rows.append({
                    'DayInt': day_val, 'Date': date_str, 'TRADE': total_trades,
                    'SELL': round(sell_lots, 3), 'BUY': round(buy_lots, 3), 'TOTAL': round(total_lots_day, 3)
                })
                data_dfs.append(data)
        except Exception as e:
            pass

    if not data_dfs:
        return None

    try:
        merged_data = pd.concat(data_dfs, ignore_index=True)
        merged_data[lots_col_idx] = pd.to_numeric(merged_data[lots_col_idx], errors='coerce')
        total_lots = round(merged_data[lots_col_idx].sum(), 3) 
        
        row_data = [np.nan] * len(merged_data.columns)
        row_data[0] = 'TOTAL LOT'
        row_data[lots_col_idx] = total_lots
        total_row = pd.DataFrame([row_data])
        
        final_df = pd.concat([header_df, merged_data, total_row], ignore_index=True)
        
        if days and prefixes:
            min_day = min(days)
            max_day = max(days)
            best_prefix = max(prefixes, key=len)
            month = list(months)[0] if months else "BULAN"
            year = list(years)[0] if years else "TAHUN"
            month_mapping = {
                'JANUARY': 'JANUARI', 'FEBRUARY': 'FEBRUARI', 'MARCH': 'MARET',
                'MAY': 'MEI', 'JULY': 'JULI', 'AUGUST': 'AGUSTUS', 
                'OCTOBER': 'OKTOBER', 'DECEMBER': 'DESEMBER'
            }
            month_norm = month_mapping.get(month, month)

            if min_day == max_day:
                merged_filename = f"{best_prefix} {min_day:02d} {month_norm} {year} (GABUNGAN).xlsx"
            else:
                merged_filename = f"{best_prefix} {min_day:02d} - {max_day:02d} {month_norm} {year} (GABUNGAN).xlsx"
        else:
            merged_filename = "REKAP TRANSAKSI GABUNGAN TOTAL.xlsx"
            
        merged_filepath = os.path.join(download_dir, merged_filename)

        with pd.ExcelWriter(merged_filepath, engine='openpyxl') as writer:
            final_df.to_excel(writer, sheet_name='BURSA', header=False, index=False)
            
            if summary_rows:
                wb = writer.book
                ws = wb.create_sheet(title="REKAP PER HARI")
                
                blue_fill = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
                light_blue_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                orange_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
                green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                white_font = Font(color="FFFFFF", bold=True)
                black_bold = Font(color="000000", bold=True)
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                thick_green_border = Border(
                    left=Side(style='medium', color="00B050"), right=Side(style='medium', color="00B050"),
                    top=Side(style='medium', color="00B050"), bottom=Side(style='medium', color="00B050")
                )
                center_align = Alignment(horizontal="center", vertical="center")
                
                current_row = 2
                summary_rows = sorted(summary_rows, key=lambda x: x['DayInt'])
                
                for row_data in summary_rows:
                    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row+1, end_column=1)
                    c_trade = ws.cell(row=current_row, column=1, value="TRADE")
                    c_trade.fill, c_trade.font, c_trade.alignment, c_trade.border = blue_fill, white_font, center_align, thin_border
                    ws.cell(row=current_row+1, column=1).border = thin_border
                    
                    c_td = ws.cell(row=current_row, column=2, value="Trade date :")
                    c_td.fill, c_td.font, c_td.alignment, c_td.border = blue_fill, white_font, center_align, thin_border
                    
                    ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=4)
                    c_date = ws.cell(row=current_row, column=3, value=row_data['Date'])
                    c_date.fill, c_date.font, c_date.alignment, c_date.border = blue_fill, white_font, center_align, thin_border
                    ws.cell(row=current_row, column=4).border = thin_border
                    
                    c_sell = ws.cell(row=current_row+1, column=2, value="SELL")
                    c_sell.fill, c_sell.font, c_sell.alignment, c_sell.border = light_blue_fill, black_bold, center_align, thin_border
                    
                    c_buy = ws.cell(row=current_row+1, column=3, value="BUY")
                    c_buy.fill, c_buy.font, c_buy.alignment, c_buy.border = orange_fill, black_bold, center_align, thin_border
                    
                    c_tot = ws.cell(row=current_row+1, column=4, value="BUY + SELL")
                    c_tot.fill, c_tot.font, c_tot.alignment, c_tot.border = green_fill, black_bold, center_align, thick_green_border
                    
                    c_tv = ws.cell(row=current_row+2, column=1, value=row_data['TRADE'])
                    c_tv.fill, c_tv.font, c_tv.alignment, c_tv.border = light_blue_fill, black_bold, center_align, thin_border
                    
                    c_sv = ws.cell(row=current_row+2, column=2, value=row_data['SELL'])
                    c_sv.fill, c_sv.alignment, c_sv.border, c_sv.number_format = light_blue_fill, center_align, thin_border, '#,##0.00'
                    
                    c_bv = ws.cell(row=current_row+2, column=3, value=row_data['BUY'])
                    c_bv.fill, c_bv.alignment, c_bv.border, c_bv.number_format = light_blue_fill, center_align, thin_border, '#,##0.00'
                    
                    c_totv = ws.cell(row=current_row+2, column=4, value=row_data['TOTAL'])
                    c_totv.fill, c_totv.font, c_totv.alignment, c_totv.border, c_totv.number_format = light_blue_fill, black_bold, center_align, thick_green_border, '#,##0.00'
                    current_row += 4
                    
                ws.column_dimensions['A'].width = 15
                ws.column_dimensions['B'].width = 20
                ws.column_dimensions['C'].width = 20
                ws.column_dimensions['D'].width = 20

        return merged_filename
    except Exception as e:
        return None

# ============================================================
# PROGRAM GENERATOR TEMPLATE EXCEL OTOMATIS (BESERTA RUMUS)
# ============================================================
def buat_template_ldk():
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side
    
    wb = Workbook()
    font_bold = Font(bold=True)
    align_center = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        top=Side(border_style="thin", color="000000"),
        left=Side(border_style="thin", color="000000"),
        right=Side(border_style="thin", color="000000"),
        bottom=Side(border_style="thin", color="000000")
    )
    
    def apply_style(ws, row, col, bold=False, center=True, border=True):
        cell = ws.cell(row=row, column=col)
        if bold: cell.font = font_bold
        if center: cell.alignment = align_center
        if border: cell.border = thin_border
        return cell

    # ================= SHEET VOL =================
    ws = wb.active
    ws.title = "Vol"
    ws['A1'] = "Form 7"
    ws['A2'] = "Rekapitulasi Volume Transaksi (Lot)"
    ws['A3'] = "PT. CENTRAL CAPITAL  FUTURES"
    ws['A4'] = 2026
    
    headers_vol = {
        'A6': "No.", 'B6': "Lokasi", 'C6': "Bulan", 
        'C7': "Kontrak Berjangka", 'E7': "Kontrak Derivatif Lainnya", 'G7': "Jumlah",
        'C8': "Pertambangan", 'D8': "Pertanian Perkebunan", 'E8': "SPA", 'F8': "PALN"
    }
    for pos, val in headers_vol.items(): ws[pos] = val
    
    ws.merge_cells('A6:A8'); ws.merge_cells('B6:B8'); ws.merge_cells('C6:G6')
    ws.merge_cells('C7:D7'); ws.merge_cells('E7:F7'); ws.merge_cells('G7:G8')
    
    for r in range(6, 9):
        for c in range(1, 8): apply_style(ws, r, c, bold=True)
        
    data_vol = [
        [1, "PUSAT", 0, 0, 0, 0, "=SUM(D9:F9)"],
        [2, "BANJARMASIN", 0, 0, 0, 0, "=SUM(D10:F10)"]
    ]
    for r_idx, row_data in enumerate(data_vol, start=9):
        for c_idx, val in enumerate(row_data, start=1):
            cell = apply_style(ws, r_idx, c_idx, center=(c_idx == 1 or c_idx >= 3))
            cell.value = val

    ws['A12'] = "TOTAL"
    ws.merge_cells('A12:B12')
    apply_style(ws, 12, 1, bold=True)
    apply_style(ws, 12, 2, bold=True)
    
    for col, letter in enumerate(['C', 'D', 'E', 'F', 'G'], start=3):
        c = apply_style(ws, 12, col, bold=True)
        c.value = f"=SUM({letter}9:{letter}10)"
        
    ws['A14'] = "Catatan"
    ws['A15'] = "*)Isi lokasi dengan nama kantor cabang, pastikan nama yang terdapat pula di tab domisili"
    ws['A16'] = "*)pastikan dalam satu baris semua colom terisi"
    
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 15
    ws.column_dimensions['G'].width = 15

    # ================= SHEET KOM =================
    wk = wb.create_sheet("Kom")
    wk['A1'] = "Form 8"
    wk['A2'] = "Rekapitulasi Komisi Transaksi"
    wk['A3'] = "PT. CENTRAL CAPITAL  FUTURES"
    wk['A4'] = 2026
    
    headers_kom = {
        'A6': "No.", 'B6': "Jenis Kontrak", 'D6': "Bulan",
        'D7': "Vol (Lot)", 'E7': "Komisi (per Lot)", 'F7': "Total Komisi"
    }
    for pos, val in headers_kom.items(): wk[pos] = val
    
    wk.merge_cells('A6:A7'); wk.merge_cells('B6:C7'); wk.merge_cells('D6:F6')
    for r in range(6, 8):
        for c in range(1, 7): apply_style(wk, r, c, bold=True)
        
    wk['A8'], wk['B8'] = 1, "Kontrak Berjangka"
    wk['D8'], wk['E8'], wk['F8'] = "=SUM(D9:D9)", "=SUM(E9:E9)", "=SUM(F9:F9)"
    wk.merge_cells('B8:C8')
    for c in range(1, 7): apply_style(wk, 8, c, bold=True)
    
    row9_data = ["-", "KGE", 0, 0, "=D9*E9"]
    for c_idx, val in enumerate(row9_data, start=2):
        apply_style(wk, 9, c_idx).value = val
    apply_style(wk, 9, 1)

    wk['A10'], wk['B10'] = 2, "Kontrak Derivatif Lainnya"
    wk['D10'], wk['E10'], wk['F10'] = "=D11+D75", "=E11+E75", "=F11+F75"
    wk.merge_cells('B10:C10')
    for c in range(1, 7): apply_style(wk, 10, c, bold=True)
    
    wk['B11'], wk['C11'] = "-", "SPA"
    wk['D11'], wk['E11'], wk['F11'] = "=SUM(D12:D74)", "=SUM(E12:E74)", "=SUM(F12:F74)"
    for c in range(1, 7): apply_style(wk, 11, c, bold=True)
    
    commodities = [
        "JPK50_BBJ", "HKJ50_BBJ", "KRK50_BBJ", "EU1010_BBJ", "UJ1010_BBJ", "GU1010_BBJ", 
        "UC1010_BBJ", "AU1010_BBJ", "UK1010_BBJ", "NU1010_BBJ", "EJ1H10_BBJ", "GJ1H10_BBJ", 
        "GA1H10", "CJ1H10_BBJ", "AJ1H10_BBJ", "EG1H10_BBJ", "EN1010_BBJ", "EC1010_BBJ", 
        "GC1010_BBJ", "AN1010_BBJ", "XAG10_BBJ", "XUL10_BBJ", "CLSK10_BBJ", "EU1212_BBJ", 
        "UJ1212_BBJ", "GU1212_BBJ", "UC1212_BBJ", "AU1212_BBJ", "UK1212_BBJ", "NU1212_BBJ", 
        "EJ1H12_BBJ", "GJ1H12_BBJ", "GA1H12", "CJ1H12_BBJ", "AJ1H12_BBJ", "EG1H12_BBJ", 
        "EN1212_BBJ", "EC1212_BBJ", "GC1212_BBJ", "AN1212_BBJ", "XAG12_BBJ", "XUL12_BBJ", 
        "CLSK12_BBJ", "EU1414_BBJ", "UJ1414_BBJ", "GU1414_BBJ", "UC1414_BBJ", "AU1414_BBJ", 
        "UK1414_BBJ", "NU1414_BBJ", "EJ1H14_BBJ", "GJ1H14_BBJ", "GA1H14", "CJ1H14_BBJ", 
        "AJ1H14_BBJ", "EG1H14_BBJ", "EN1414_BBJ", "EC1414_BBJ", "GC1414_BBJ", "AN1414_BBJ", 
        "XAG14_BBJ", "XUL14_BBJ", "CLSK14_BBJ"
    ]
    
    r_idx = 12
    for comm in commodities:
        apply_style(wk, r_idx, 1)
        apply_style(wk, r_idx, 2).value = "-"
        apply_style(wk, r_idx, 3).value = comm
        apply_style(wk, r_idx, 4).value = 0
        apply_style(wk, r_idx, 5).value = 20
        apply_style(wk, r_idx, 6).value = f"=D{r_idx}*E{r_idx}"
        r_idx += 1
        
    apply_style(wk, r_idx, 1)
    apply_style(wk, r_idx, 2).value = "-"
    apply_style(wk, r_idx, 3).value = "PALN"
    apply_style(wk, r_idx, 4).value = "=(0)*2"
    apply_style(wk, r_idx, 5).value = f"=E{r_idx+1}"
    apply_style(wk, r_idx, 6).value = f"=F{r_idx+1}"
    
    apply_style(wk, r_idx+1, 1)
    apply_style(wk, r_idx+1, 2).value = "-"
    apply_style(wk, r_idx+1, 3).value = ""
    apply_style(wk, r_idx+1, 4).value = "=(0)*2"
    apply_style(wk, r_idx+1, 5).value = 0
    apply_style(wk, r_idx+1, 6).value = f"=D{r_idx+1}*E{r_idx+1}"
    
    apply_style(wk, r_idx+2, 1)
    apply_style(wk, r_idx+2, 2)
    apply_style(wk, r_idx+2, 3, bold=True).value = "TOTAL"
    apply_style(wk, r_idx+2, 4, bold=True).value = "=D8+D10"
    apply_style(wk, r_idx+2, 5, bold=True).value = "=E8+E10"
    apply_style(wk, r_idx+2, 6, bold=True).value = "=F8+F10"
    
    r_cat = r_idx + 4
    wk.cell(row=r_cat, column=2).value = "Catatan :"
    wk.cell(row=r_cat+1, column=2).value = "Kode kontrak harus diisi jika tidak diisi maka data tidakakan dibaca"
    wk.cell(row=r_cat+2, column=2).value = "Data harus berurut jika terselip baris yang kosong maka baris berikunya tidak dibaca"
    
    wk.column_dimensions['B'].width = 5
    wk.column_dimensions['C'].width = 25
    wk.column_dimensions['D'].width = 15
    wk.column_dimensions['E'].width = 18
    wk.column_dimensions['F'].width = 18
    
    return wb

# ============================================================
# PROGRAM 2: REKAP LOT LDK
# ============================================================

ALAMAT_BANJARMASIN = "DED-STY-AYN-BJM-CCF"
PREFIX_KLIEN       = "CC"
KOMISI_DEFAULT     = 20
UTAMAKAN_REVISI    = True
COL_COMMODITY, COL_LOT, COL_K, COL_N = 3, 6, 10, 13

def normalisasi_kode(v):
    s = str(v).replace('\u00a0', ' ').strip()
    if re.fullmatch(r'\d+\.0+', s):
        s = s.split('.')[0]
    return s

def ekstrak_akun(row):
    kandidat = [str(row[COL_K]).strip(), str(row[COL_N]).strip()]
    for v in kandidat:
        if v.upper().startswith(PREFIX_KLIEN):
            mm = re.search(r'(\d{5,})', v)
            if mm:
                return mm.group(1)
    for v in kandidat:
        mm = re.search(r'(\d{5,})', v)
        if mm:
            return mm.group(1)
    return None

def cari_baris(ws, teks, kolom=3):
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(r, kolom).value).strip().upper() == teks.upper():
            return r
    return None

def sisipkan_komoditas(ws, kode, lot):
    r_spa = cari_baris(ws, 'SPA')
    mm = re.match(r'=SUM\(D(\d+):D(\d+)\)', str(ws.cell(r_spa, 4).value))
    if not mm:
        return None
    awal, akhir = int(mm.group(1)), int(mm.group(2))
    baru = akhir + 1
    ws.insert_rows(baru)
    for c in range(1, 7):
        ws.cell(baru, c)._style = copy.copy(ws.cell(baru - 1, c)._style)
    ws.cell(baru, 2).value = '-'
    ws.cell(baru, 3).value = kode
    ws.cell(baru, 4).value = float(lot)
    ws.cell(baru, 5).value = KOMISI_DEFAULT
    ws.cell(baru, 6).value = f'=D{baru}*E{baru}'
    for c, h in ((4, 'D'), (5, 'E'), (6, 'F')):
        ws.cell(r_spa, c).value = f'=SUM({h}{awal}:{h}{baru})'
    r_paln  = cari_baris(ws, 'PALN')
    r_lain  = cari_baris(ws, 'Kontrak Derivatif Lainnya', kolom=2)
    r_total = cari_baris(ws, 'TOTAL')
    
    if r_paln:
        ws.cell(r_paln, 5).value = f'=E{r_paln + 1}'
        ws.cell(r_paln, 6).value = f'=F{r_paln + 1}'
        ws.cell(r_paln + 1, 6).value = f'=D{r_paln + 1}*E{r_paln + 1}'
        if r_lain:
            for c, h in ((4, 'D'), (5, 'E'), (6, 'F')):
                ws.cell(r_lain, c).value = f'={h}{r_spa}+{h}{r_paln}'
    if r_total and r_lain:
        for c, h in ((4, 'D'), (5, 'E'), (6, 'F')):
            ws.cell(r_total, c).value = f'={h}8+{h}{r_lain}'
    return baru

# ============================================================
# GUI STREAMLIT (ATAS & BAWAH)
# ============================================================

st.title("📊 Aplikasi Rekap Transaksi & LDK")
st.markdown("---")

# ============================================================
# BAGIAN ATAS: UNDUH & GABUNG
# ============================================================
st.header("Yahoo IMAP Attachment Downloader")
st.write("Fitur untuk mencari, mengunduh, dan menggabungkan lampiran Excel (Rekap Bursa) dari email Yahoo secara otomatis.")

col1, col2 = st.columns(2)
with col1:
    email_acc = st.text_input("Akun Email Yahoo", value="dealingccf_bbj@yahoo.com")
    app_pass = "krrrgdmbdoxorfbv"
with col2:
    subject_input = st.text_input("Masukkan Subjek (pisahkan koma jika banyak)", value="REKAP TRANSAKSI Shift III JANUARI 2026")
    imap_server = st.text_input("Server IMAP", value="imap.mail.yahoo.com")

if st.button("Mulai Proses Unduh & Gabung", type="primary"):
    if not subject_input:
        st.warning("Pencarian dibatalkan karena subjek kosong.")
    else:
        DOWNLOAD_DIR = "lampiran_email_temp"
        if os.path.exists(DOWNLOAD_DIR):
            shutil.rmtree(DOWNLOAD_DIR)
        os.makedirs(DOWNLOAD_DIR, exist_ok=True)

        queries = [q.strip() for q in subject_input.split(",") if q.strip()]
        safe_input_subject = clean_filename(queries[0])[:50] or "Hasil_Lampiran"
        dynamic_zip_filename = f"{safe_input_subject}.zip"

        try:
            with st.spinner("Menghubungkan ke server Yahoo IMAP..."):
                mail = imaplib.IMAP4_SSL(imap_server)
                mail.login(email_acc, app_pass)
                
            mail.select('"INBOX"')
            shift3_ids = set()
            shift2_ids = set()

            st.info("Mencari email yang cocok dengan kriteria...")
            for query in queries:
                e_ids = do_imap_search(mail, query)
                shift3_ids.update(e_ids)

                if "Shift III" in query:
                    fallback_query = query.replace("Shift III", "Shift II")
                    fb_ids = do_imap_search(mail, fallback_query)
                    shift2_ids.update(fb_ids)

            shift2_ids = shift2_ids - shift3_ids
            total_emails = len(shift3_ids) + len(shift2_ids)

            if total_emails == 0:
                st.error("Tidak ada satupun email yang ditemukan dari semua kriteria pencarian.")
            else:
                global download_count
                download_count = 0
                downloaded_dates = set()
                
                progress_bar = st.progress(0)
                status_text = st.empty()

                def process_emails(email_ids, is_fallback=False):
                    global download_count
                    for idx, eid in enumerate(list(email_ids)):
                        res, msg_data = mail.fetch(eid, "(RFC822)")
                        if res == "OK":
                            for response_part in msg_data:
                                if isinstance(response_part, tuple):
                                    msg = email.message_from_bytes(response_part[1])
                                    raw_subject = msg.get("Subject", "Tanpa_Subjek")
                                    email_subject = decode_mime_header(raw_subject)
                                    safe_subject = clean_filename(email_subject)

                                    if msg.is_multipart():
                                        for part in msg.walk():
                                            if part.get_content_maintype() == 'multipart': continue
                                            if part.get('Content-Disposition') is None: continue

                                            original_filename = part.get_filename()
                                            if original_filename:
                                                original_filename = decode_mime_header(original_filename)
                                                _, ext = os.path.splitext(original_filename)
                                                filename = clean_filename(f"{safe_subject}{ext}")
                                                date_key = extract_date_key(filename)
                                                
                                                if date_key:
                                                    if is_fallback and date_key in downloaded_dates: continue
                                                    if not is_fallback: downloaded_dates.add(date_key)

                                                filepath = os.path.join(DOWNLOAD_DIR, filename)
                                                base, ext = os.path.splitext(filename)
                                                counter = 1
                                                while os.path.exists(filepath):
                                                    filepath = os.path.join(DOWNLOAD_DIR, f"{base}_{counter}{ext}")
                                                    counter += 1

                                                with open(filepath, "wb") as f:
                                                    f.write(part.get_payload(decode=True))
                                                download_count += 1
                        progress_bar.progress((idx + 1) / total_emails)

                process_emails(shift3_ids, is_fallback=False)
                process_emails(shift2_ids, is_fallback=True)
                
                st.info("Memeriksa ketersediaan file 'Revisi'...")
                all_downloaded = glob.glob(os.path.join(DOWNLOAD_DIR, "*.*"))
                files_by_date = {}
                for fpath in all_downloaded:
                    fname = os.path.basename(fpath)
                    d_key = extract_date_key(fname)
                    ext = os.path.splitext(fname)[1].lower()
                    if d_key:
                        group_key = f"{d_key}_{ext}"
                        if group_key not in files_by_date: files_by_date[group_key] = []
                        files_by_date[group_key].append(fpath)

                for group_key, flist in files_by_date.items():
                    if len(flist) > 1:
                        revisi_files = [f for f in flist if 'revisi' in os.path.basename(f).lower()]
                        if revisi_files:
                            revisi_files.sort()
                            kept_file = revisi_files[-1]
                            for f in flist:
                                if f != kept_file:
                                    try:
                                        os.remove(f)
                                        download_count -= 1
                                    except: pass

                if download_count > 0:
                    st.success(f"Berhasil mendapatkan {download_count} file lampiran bersih.")
                    merge_downloaded_excel(DOWNLOAD_DIR)
                    
                    shutil.make_archive(dynamic_zip_filename.replace('.zip', ''), 'zip', DOWNLOAD_DIR)
                    with open(dynamic_zip_filename, "rb") as f:
                        st.download_button("📥 Unduh Hasil ZIP", data=f, file_name=dynamic_zip_filename, mime="application/zip")
                else:
                    st.warning("Email ditemukan, tetapi tidak ada file lampiran di dalamnya (atau telah terfilter).")

        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses email: {e}")
        finally:
            try:
                mail.close()
                mail.logout()
            except:
                pass

st.markdown("---")

# ============================================================
# BAGIAN BAWAH: EKSTRAK & REKAP (TANPA UPLOAD TEMPLATE & FIX BURSA)
# ============================================================
st.header("Rekap Lot LDK")
st.write("Fitur memproses file ZIP Rekap Transaksi, pemetaan List ACC, dan injeksi data otomatis (Rumus template dipertahankan).")

col_a, col_b = st.columns(2)
with col_a:
    zip_file_upload = st.file_uploader("1. Upload ZIP Rekap Transaksi", type=["zip"])
with col_b:
    acc_file_upload = st.file_uploader("2. Upload List ACC (.xlsx)", type=["xlsx"])

if st.button("Mulai Proses Ekstrak & Rekap", type="primary"):
    if not (zip_file_upload and acc_file_upload):
        st.error("Harap unggah file ZIP Rekap dan List ACC terlebih dahulu.")
    else:
        EXTRACT_DIR = "rekap_extracted_temp"
        if os.path.exists(EXTRACT_DIR):
            shutil.rmtree(EXTRACT_DIR)
        os.makedirs(EXTRACT_DIR, exist_ok=True)
        
        # Simpan file upload ke server sementara
        with open("temp_rekap.zip", "wb") as f: f.write(zip_file_upload.getbuffer())
        with open("temp_acc.xlsx", "wb") as f: f.write(acc_file_upload.getbuffer())

        try:
            with zipfile.ZipFile("temp_rekap.zip", 'r') as z:
                z.extractall(EXTRACT_DIR)

            file_rekaps = []
            for root, dirs, fl in os.walk(EXTRACT_DIR):
                if '__MACOSX' in root: continue
                for f in fl:
                    if f.endswith(('.xlsx', '.xls')) and not f.startswith(('~', '.')):
                        if 'GABUNGAN' in f.upper():
                            continue
                        file_rekaps.append(os.path.join(root, f))
            file_rekaps.sort()

            if UTAMAKAN_REVISI:
                POLA_REVISI = r'\b(REV|REVISI|REVISED|PERBAIKAN|FIX|UPDATE)\b'
                def is_revisi(p): return re.search(POLA_REVISI, os.path.basename(p).upper()) is not None
                def kunci(p):
                    n = re.sub(r'\.(XLSX|XLS)$', '', os.path.basename(p).upper())
                    n = re.sub(POLA_REVISI, ' ', n)
                    return re.sub(r'\s+', ' ', n).strip()

                ada_revisi = {kunci(p) for p in file_rekaps if is_revisi(p)}
                buang = [p for p in file_rekaps if not is_revisi(p) and kunci(p) in ada_revisi]
                file_rekaps = [p for p in file_rekaps if p not in buang]

            if not file_rekaps:
                st.error("Tidak ada file rekap yang terbaca dari ZIP.")
            else:
                nama_bulan, tahun = "Bulan", None
                for bln in ['JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI', 'JULI', 'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER']:
                    if bln in os.path.basename(file_rekaps[0]).upper():
                        nama_bulan = bln.capitalize()
                        break
                m = re.search(r'(20\d{2})', os.path.basename(file_rekaps[0]))
                if m: tahun = int(m.group(1))

                st.info(f"Memproses {len(file_rekaps)} file rekap | Periode: {nama_bulan} {tahun or ''}")

                # ---- PEMBACAAN DINAMIS SHEET BURSA ----
                data_dfs = []
                for f in file_rekaps:
                    try:
                        xls = pd.ExcelFile(f)
                        bursa_sheet = [s for s in xls.sheet_names if "BURSA" in s.upper()]
                        
                        if bursa_sheet:
                            df_temp = pd.read_excel(f, sheet_name=bursa_sheet[0], header=None)
                            data_dfs.append(df_temp)
                        else:
                            df_temp = pd.read_excel(f, sheet_name=0, header=None)
                            data_dfs.append(df_temp)
                    except Exception as e:
                        pass # Lewati jika file rusak / tidak bisa dibaca
                
                if not data_dfs:
                    st.error("Data tidak dapat dibaca dari file-file rekap.")
                else:
                    df = pd.concat(data_dfs, ignore_index=True)
                    df = df.sort_values(by=0, key=lambda x: x.astype(str)).reset_index(drop=True)
                    df['Commodity'] = df[COL_COMMODITY].astype(str).str.replace('\u00a0', ' ').str.strip()
                    df = df[df['Commodity'] != 'Commodity']
                    df['Lot'] = pd.to_numeric(df[COL_LOT], errors='coerce').fillna(0)
                    df['ACC'] = df.apply(ekstrak_akun, axis=1)
                    
                    total_mentah = round(float(df['Lot'].sum()), 4)

                    df_valid = df.dropna(subset=['ACC']).copy()
                    acc = pd.read_excel("temp_acc.xlsx", header=None)
                    acc['Login']   = acc[0].map(normalisasi_kode)
                    acc['Address'] = acc[5].map(lambda v: str(v).replace('\u00a0', ' ').strip())
                    dict_address = dict(zip(acc['Login'], acc['Address']))

                    df_valid['ACC']     = df_valid['ACC'].map(normalisasi_kode)
                    df_valid['Address'] = df_valid['ACC'].map(dict_address)

                    cocok = df_valid['Address'].notna().sum()
                    if cocok == 0:
                        st.error("FATAL: tidak satu pun akun cocok dengan List ACC. Cek kolom A (Login) dan kolom F (address) di List_ACC.")
                    else:
                        kond = df_valid['Address'] == ALAMAT_BANJARMASIN
                        lot_bjm   = round(float(df_valid.loc[kond, 'Lot'].sum()), 4)
                        lot_pusat = round(float(df_valid.loc[~kond, 'Lot'].sum()), 4)
                        
                        st.write(f"**Hasil Peta Akun:** PUSAT = {lot_pusat} | BANJARMASIN = {lot_bjm} | TOTAL = {round(lot_pusat + lot_bjm, 4)}")

                        df_kom = df_valid.groupby('Commodity')['Lot'].sum().reset_index()
                        df_kom['Lot'] = df_kom['Lot'].round(4)

                        file_output = f"Output_LDK_{nama_bulan}{tahun or ''}.xlsx"
                        
                        # -------------------------------------------------------------
                        # MENGGUNAKAN TEMPLATE OTOMATIS (Tanpa File Eksternal)
                        # -------------------------------------------------------------
                        wb = buat_template_ldk()
                        
                        ws = wb['Vol']
                        ws['C6'] = nama_bulan
                        if tahun: ws['A4'] = tahun
                        for row in ws.iter_rows():
                            lok = str(row[1].value).upper().replace(" ", "")
                            if 'PUSAT' in lok: row[4].value = lot_pusat
                            elif 'BANJARMASIN' in lok: row[4].value = lot_bjm

                        wk = wb['Kom']
                        wk['D6'] = nama_bulan
                        if tahun: wk['A4'] = tahun

                        for _, d in df_kom.iterrows():
                            kode = str(d['Commodity']).strip()
                            if kode in ("nan", ""): continue
                            lot = float(d['Lot'])

                            peta = {}
                            for r in range(2, wk.max_row + 1):
                                k = str(wk.cell(r, 3).value).strip()
                                if k and k.lower() != 'none': peta.setdefault(k.upper(), r)

                            baris = peta.get(kode.upper())
                            if baris is None:
                                for k_ldk, r in peta.items():
                                    if kode.upper() in k_ldk:
                                        baris = r
                                        break
                            if baris is None:
                                baris = sisipkan_komoditas(wk, kode, lot)
                            else:
                                wk.cell(baris, 4).value = lot

                        wb.save(file_output)

                        # VERIFIKASI (Mengecek file yang telah dibuat)
                        cek = load_workbook(file_output, data_only=True)
                        v, k = cek['Vol'], cek['Kom']
                        
                        r_pusat = next(r for r in range(1, v.max_row + 1) if 'PUSAT' in str(v.cell(r, 2).value).upper().replace(" ", ""))
                        r_bjm   = next(r for r in range(1, v.max_row + 1) if 'BANJARMASIN' in str(v.cell(r, 2).value).upper().replace(" ", ""))
                        
                        masalah = []
                        val_pusat = load_workbook(file_output)['Vol'].cell(r_pusat, 5).value
                        val_bjm = load_workbook(file_output)['Vol'].cell(r_bjm, 5).value

                        if val_pusat != lot_pusat: masalah.append(f"Vol PUSAT {val_pusat} != {lot_pusat}")
                        if val_bjm != lot_bjm: masalah.append(f"Vol BJM {val_bjm} != {lot_bjm}")

                        if masalah:
                            st.error("GAGAL VERIFIKASI INPUT:\n - " + "\n - ".join(masalah))
                        else:
                            st.success("✅ Semua data berhasil diproses dan diinjeksi ke template dengan aman!")
                            with open(file_output, "rb") as f:
                                st.download_button("📥 Unduh Hasil Excel", data=f, file_name=file_output, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses Rekap LDK: {e}")
