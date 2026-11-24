# -*- coding: utf-8 -*-
"""
Script to generate Panduan_Tata_Langkah_Pengerjaan_Kasus_A.docx
A practical, operational, step-by-step guidance document for executing Kasus A with OWASP Threat Dragon.
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, hex_color):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shd)

def set_cell_margins(cell, top=90, bottom=90, left=130, right=130):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}>'
                      f'<w:top w:w="{top}" w:type="dxa"/>'
                      f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
                      f'<w:left w:w="{left}" w:type="dxa"/>'
                      f'<w:right w:w="{right}" w:type="dxa"/>'
                      f'</w:tcMar>')
    tcPr.append(tcMar)

def set_table_borders(table, border_color="718096", inside_h_color="CBD5E0", inside_v_color="E2E8F0"):
    tblPr = table._tbl.tblPr
    for child in list(tblPr):
        if child.tag.endswith('tblBorders'):
            tblPr.remove(child)
    borders = parse_xml(f'<w:tblBorders {nsdecls("w")}>'
                        f'<w:top w:val="single" w:sz="6" w:space="0" w:color="{border_color}"/>'
                        f'<w:left w:val="single" w:sz="6" w:space="0" w:color="{border_color}"/>'
                        f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="{border_color}"/>'
                        f'<w:right w:val="single" w:sz="6" w:space="0" w:color="{border_color}"/>'
                        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="{inside_h_color}"/>'
                        f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="{inside_v_color}"/>'
                        f'</w:tblBorders>')
    tblPr.append(borders)

def format_table(table, col_widths, headers, data, align_cols=None, border_color="718096", total_dxa=9360):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    tblPr = table._tbl.tblPr
    for child in list(tblPr):
        if child.tag.endswith('tblW') or child.tag.endswith('tblLayout'):
            tblPr.remove(child)
    tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="{total_dxa}" w:type="dxa"/>')
    tblPr.append(tblW)
    tblLayout = parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>')
    tblPr.append(tblLayout)

    set_table_borders(table, border_color=border_color, inside_h_color="CBD5E0", inside_v_color="E2E8F0")

    sum_widths = sum(col_widths)
    col_dxas = [int((w / sum_widths) * total_dxa) for w in col_widths]
    col_dxas[-1] += total_dxa - sum(col_dxas)

    # Header row
    hdr_row = table.rows[0]
    trPr = hdr_row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

    hdr_cells = hdr_row.cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        set_cell_background(hdr_cells[i], "1A365D")
        set_cell_margins(hdr_cells[i], top=110, bottom=110, left=130, right=130)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(9.5)
            run.font.name = "Calibri"

    # Data rows
    for r_idx, row_data in enumerate(data):
        row = table.rows[r_idx + 1]
        r_trPr = row._tr.get_or_add_trPr()
        r_trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        row_cells = row.cells
        bg_col = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            row_cells[c_idx].text = str(val)
            set_cell_background(row_cells[c_idx], bg_col)
            set_cell_margins(row_cells[c_idx], top=80, bottom=80, left=120, right=120)
            p = row_cells[c_idx].paragraphs[0]
            if align_cols and c_idx in align_cols:
                p.alignment = align_cols[c_idx]
            else:
                if len(str(val)) > 35:
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for run in p.runs:
                run.font.size = Pt(9)
                run.font.name = "Calibri"
                run.font.color.rgb = RGBColor(45, 55, 72)

    # Set column and cell widths explicitly
    for row in table.rows:
        for i, w_dxa in enumerate(col_dxas):
            cell = row.cells[i]
            cell.width = Inches(col_widths[i])
            tcPr = cell._tc.get_or_add_tcPr()
            for child in list(tcPr):
                if child.tag.endswith('tcW'):
                    tcPr.remove(child)
            tcPr.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{w_dxa}" w:type="dxa"/>'))

def add_heading_1(doc, text):
    h = doc.add_paragraph()
    h.paragraph_format.space_before = Pt(15)
    h.paragraph_format.space_after = Pt(5)
    h.paragraph_format.keep_with_next = True
    run = h.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(13.5)
    run.font.bold = True
    run.font.color.rgb = RGBColor(26, 54, 93)
    return h

def add_heading_2(doc, text):
    h = doc.add_paragraph()
    h.paragraph_format.space_before = Pt(11)
    h.paragraph_format.space_after = Pt(3)
    h.paragraph_format.keep_with_next = True
    run = h.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(11.5)
    run.font.bold = True
    run.font.color.rgb = RGBColor(43, 108, 176)
    return h

def add_heading_3(doc, text):
    h = doc.add_paragraph()
    h.paragraph_format.space_before = Pt(8)
    h.paragraph_format.space_after = Pt(2)
    h.paragraph_format.keep_with_next = True
    run = h.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(10)
    run.font.bold = True
    run.font.color.rgb = RGBColor(45, 55, 72)
    return h

def add_body_paragraph(doc, text, bold_prefix="", space_after=5):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = "Calibri"
        r_pre.font.size = Pt(9.5)
        r_pre.font.bold = True
        r_pre.font.color.rgb = RGBColor(45, 55, 72)
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(9.5)
    run.font.color.rgb = RGBColor(45, 55, 72)
    return p

def add_bullet_point(doc, text, bold_prefix=""):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = "Calibri"
        r_pre.font.size = Pt(9.5)
        r_pre.font.bold = True
        r_pre.font.color.rgb = RGBColor(45, 55, 72)
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(9.5)
    run.font.color.rgb = RGBColor(45, 55, 72)
    return p

def add_callout(doc, text, title="PERHATIAN TEKNIS", total_dxa=9360):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    tblPr = table._tbl.tblPr
    for child in list(tblPr):
        if child.tag.endswith('tblW') or child.tag.endswith('tblLayout'):
            tblPr.remove(child)
    tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="{total_dxa}" w:type="dxa"/>')
    tblPr.append(tblW)
    tblLayout = parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>')
    tblPr.append(tblLayout)

    set_table_borders(table, border_color="D69E2E", inside_h_color="D69E2E", inside_v_color="D69E2E")

    for row in table.rows:
        r_trPr = row._tr.get_or_add_trPr()
        r_trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

    cell = table.cell(0, 0)
    set_cell_background(cell, "FEFCBF") # subtle warm amber
    set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
    cell.width = Inches(6.5)
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith('tcW'):
            tcPr.remove(child)
    tcPr.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{total_dxa}" w:type="dxa"/>'))

    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.15
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    rt = p.add_run(f"[{title}] ")
    rt.font.bold = True
    rt.font.size = Pt(9)
    rt.font.color.rgb = RGBColor(116, 66, 16)
    rb = p.add_run(text)
    rb.font.size = Pt(9)
    rb.font.color.rgb = RGBColor(116, 66, 16)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def build_panduan():
    doc = docx.Document()

    # Set page margins
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Header / Title
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(8)
    title_p.paragraph_format.space_after = Pt(3)
    run_t = title_p.add_run("PANDUAN TATA LANGKAH PENGERJAAN TEKNIS")
    run_t.font.name = "Calibri"
    run_t.font.size = Pt(17)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor(26, 54, 93)

    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_after = Pt(12)
    run_s = sub_p.add_run("Threat Modeling STRIDE Menggunakan OWASP Threat Dragon\nKasus A: Sistem Presensi Perkuliahan Berbasis QR Code (Kelompok 4 Orang)")
    run_s.font.name = "Calibri"
    run_s.font.size = Pt(11)
    run_s.font.color.rgb = RGBColor(74, 85, 104)

    # Info table
    info_table = doc.add_table(rows=6, cols=2)
    info_headers = ["Parameter", "Keterangan Operasional"]
    info_data = [
        ["Mata Kuliah & Semester", "Keamanan Siber (CAK3CAB3) — 2627-1"],
        ["Studi Kasus Analisis", "Kasus A: Sistem Presensi Perkuliahan Berbasis QR Code"],
        ["Anggota Tim Penyusun", "1. Fatir Gibran (103112430153)\n2. Tio Armani (103112430225)\n3. Chadafya Putra Zulfikar (103112430173)\n4. Tedy Permana Putra (103112430157)"],
        ["Platform / Perangkat", "OWASP Threat Dragon (Desktop atau Web: https://www.threatdragon.com/)"],
        ["Format Luaran Wajib", "1. Berkas Model (.json) | 2. Laporan Hasil Ekspor (.pdf) | 3. Laporan Analisis (.docx)"]
    ]
    format_table(info_table, [2.2, 4.3], info_headers, info_data)
    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 1. GAMBARAN UMUM
    add_heading_1(doc, "1. Gambaran Umum dan Ketentuan Pengumpulan")
    add_heading_2(doc, "1.1 Berkas Akhir yang Wajib Dikumpulkan")
    add_body_paragraph(doc, "Setiap kelompok wajib mengumpulkan tiga berkas utama sesuai instruksi dosen:")
    add_bullet_point(doc, "File model keluaran OWASP Threat Dragon yang memuat diagram DFD lengkap, 4 trust boundary, dan seluruh 14 entri ancaman. File ini harus valid dan dapat dibuka kembali.", "1. File Model Threat Dragon (.json): ")
    add_bullet_point(doc, "Laporan yang diekspor langsung dari fitur Report di Threat Dragon dengan mencentang seluruh opsi (Threat model diagrams, Empty model elements, Element properties).", "2. Laporan Resmi Threat Dragon (.pdf): ")
    add_bullet_point(doc, "Dokumen laporan komprehensif hasil perakitan tim (menggunakan Laporan_Threat_Modeling_Kasus_A_Presensi_QR.docx) yang menyajikan analisis, metodologi, dan justifikasi lengkap.", "3. Laporan Pendukung Analisis (.docx): ")

    add_heading_2(doc, "1.2 Ketentuan dan Kriteria Penilaian Dosen")
    add_bullet_point(doc, "DFD harus merepresentasikan aliran data dari browser pengguna (mahasiswa, dosen, admin) ke Backend Presensi, interaksi ke Layanan Autentikasi Kampus (SSO), serta akses basis data akademik. Diagram dijaga pada level komponen utama (sekitar 5–8 elemen utama).", "Struktur DFD: ")
    add_bullet_point(doc, "Dosen mensyaratkan minimal 6 ancaman berbeda. Pada rancangan Kasus A ini, kelompok menyediakan 14 ancaman komprehensif yang mencakup seluruh kategori STRIDE (S, T, R, I, D, E) dengan justifikasi skenario, dampak, skor risiko, dan mitigasi spesifik.", "Jumlah & Kualitas Ancaman: ")
    add_bullet_point(doc, "Maksimal 3–5 asumsi teknis yang diberi label eksplisit 'asumsi'. Mekanisme kontrol yang tidak tertulis pada deskripsi sistem tidak boleh langsung dianggap pasti tidak ada.", "Batasan Asumsi: ")
    add_bullet_point(doc, "Gunakan data fiktif. Tidak diperkenankan memasukkan kredensial nyata, token otentikasi asli, atau informasi pribadi sensitif.", "Integritas Data: ")

    # 2. PEMBAGIAN PERAN
    add_heading_1(doc, "2. Pembagian Peran dan Tanggung Jawab Anggota (Jobdesk Tim)")
    add_body_paragraph(doc, "Guna memastikan efisiensi kerja tim dan mencegah konflik file, pembagian tanggung jawab ditetapkan secara presisi untuk 4 anggota:")

    role_table = doc.add_table(rows=5, cols=4)
    role_headers = ["Anggota & NIM", "Peran Utama", "Tanggung Jawab Spesifik (Jobdesk)", "Luaran (Deliverable)"]
    role_data = [
        ["Fatir Gibran\n(103112430153)", "Ketua Tim & Pemodel DFD (Lead Architect)", "Memimpin analisis ruang lingkup, menyepakati asumsi A1–A5, membuat model baru di Threat Dragon, menyusun 8 elemen DFD, menghubungkan 16 data flow, dan memasang 4 trust boundary.", "File KasusA_Presensi_QR_v1_DFD.json dan draf Bab 1 & 2 dokumen laporan."],
        ["Tio Armani\n(103112430225)", "Analis STRIDE: Spoofing & Tampering", "Menganalisis potensi pemalsuan identitas (titip absen, pencurian token, pembagian QR, tebakan token) dan perusakan data (manipulasi ID, race condition, SQL injection). Menginput ancaman T01 s/d T07 ke Threat Dragon.", "File KasusA_Presensi_QR_v2_ST.json dan draf tabel kartu ancaman T01–T07."],
        ["Chadafya Putra Zulfikar\n(103112430173)", "Analis STRIDE: R, I, D, E", "Menganalisis potensi penolakan tindakan (audit log), kebocoran data rekap/database, serangan DoS pada backend/SSO, dan eskalasi wewenang dosen/admin. Menginput ancaman T08 s/d T14 ke Threat Dragon.", "File KasusA_Presensi_QR_v3_Semua.json dan draf tabel kartu ancaman T08–T14."],
        ["Tedy Permana Putra\n(103112430157)", "Penilai Risiko & QA Editor Final", "Mereview skor risiko (L×D) seluruh ancaman, memastikan konsistensi keparahan (9 High, 5 Medium), memvalidasi kelengkapan diagram, mengekspor PDF Report dari tool, merakit dokumen Word final, dan menguji buka ulang berkas .json.", "File KasusA_KelompokXX_FINAL.json, PDF Report resmi Threat Dragon, dan Laporan Word final."]
    ]
    format_table(role_table, [1.5, 1.5, 2.2, 1.4], role_headers, role_data)

    # 3. MANAJEMEN ALUR KERJA DAN KONTROL VERSI
    add_heading_1(doc, "3. Manajemen Alur Kerja dan Kontrol Versi Berkas")
    add_heading_2(doc, "3.1 Timeline Alokasi Waktu (90 Menit)")
    
    time_table = doc.add_table(rows=7, cols=4)
    time_headers = ["Fase Waktu", "Fokus Aktivitas", "Penanggung Jawab", "Media / Alat"]
    time_data = [
        ["00 – 10 Menit", "Membaca deskripsi Kasus A, menyepakati alur bisnis dan 5 asumsi.", "Seluruh Anggota", "Grup Diskusi / Video Call"],
        ["10 – 35 Menit", "Membuat model dan menggambar DFD lengkap di Threat Dragon.", "Anggota 1", "OWASP Threat Dragon"],
        ["35 – 40 Menit", "Review bersama hasil DFD; simpan v1; serahkan file ke Anggota 2.", "Seluruh Anggota", "Google Drive / OneDrive"],
        ["40 – 60 Menit", "Input ancaman T01–T07 (Anggota 2), dilanjutkan T08–T14 (Anggota 3).", "Anggota 2 & 3", "OWASP Threat Dragon (Bergiliran)"],
        ["60 – 75 Menit", "Audit skor risiko, verifikasi tautan elemen, perbaikan salah ketik.", "Anggota 4 & Semua", "Threat Dragon & Draf Laporan"],
        ["75 – 90 Menit", "Ekspor PDF dari Threat Dragon, uji buka ulang file .json, finalisasi .docx.", "Anggota 4", "Threat Dragon & MS Word"]
    ]
    format_table(time_table, [1.2, 2.5, 1.4, 1.4], time_headers, time_data)

    add_heading_2(doc, "3.2 Skema Kontrol Versi Berkas JSON")
    add_callout(doc, "JANGAN PERNAH dua orang mengedit file model .json Threat Dragon secara bersamaan pada laptop berbeda! Threat Dragon menyimpan model dalam satu file JSON lokal tunggal. Menggabungkan dua file JSON secara manual sangat rentan merusak koordinat diagram dan ID elemen.", "ATURAN MUTLAK KOLABORASI")
    
    ver_table = doc.add_table(rows=5, cols=3)
    ver_headers = ["Nama Berkas Estafet", "Dikerjakan Oleh", "Muatan Isi Berkas"]
    ver_data = [
        ["KasusA_KelompokXX_v1_DFD.json", "Anggota 1", "Arsitektur DFD lengkap (8 elemen, 16 aliran data, 4 boundary) tanpa ancaman."],
        ["KasusA_KelompokXX_v2_ST.json", "Anggota 2", "Basis v1 + 7 entri ancaman kategori Spoofing dan Tampering (T01 s/d T07)."],
        ["KasusA_KelompokXX_v3_Semua.json", "Anggota 3", "Basis v2 + 7 entri ancaman kategori R, I, D, E (T08 s/d T14). Total 14 ancaman."],
        ["KasusA_KelompokXX_FINAL.json", "Anggota 4", "Hasil review final, koreksi typo, validasi status Open dan severity. Siap dikumpulkan."]
    ]
    format_table(ver_table, [2.3, 1.4, 2.8], ver_headers, ver_data)

    # 4. PANDUAN EKSEKUSI THREAT DRAGON
    add_heading_1(doc, "4. Panduan Operasional Teknis di OWASP Threat Dragon")
    
    add_heading_2(doc, "4.1 Langkah 1: Inisialisasi Model Baru (Anggota 1)")
    add_bullet_point(doc, "Buka aplikasi desktop OWASP Threat Dragon atau akses versi web pada https://www.threatdragon.com/ lalu pilih Local Storage.", "1. Buka Aplikasi: ")
    add_bullet_point(doc, "Klik tombol New Model. Isi kolom metadata sebagai berikut:\n• Title: STRIDE - Kasus A - Presensi QR\n• Owner: Tim Mahasiswa Kasus A\n• Reviewer: Dosen Praktikum Keamanan Siber (CAK3CAB3)\n• Description: Threat modeling sistem presensi perkuliahan berbasis QR code menggunakan kerangka kerja STRIDE.\n• Contributors: Fatir Gibran (103112430153), Tio Armani (103112430225), Chadafya Putra Zulfikar (103112430173), Tedy Permana Putra (103112430157)", "2. Isi Form Identitas: ")
    add_bullet_point(doc, "Klik tombol Save. Model dasar berhasil diinisialisasi.", "3. Simpan: ")

    add_heading_2(doc, "4.2 Langkah 2: Pembuatan Kanvas Diagram (Anggota 1)")
    add_bullet_point(doc, "Pada halaman model detail, klik tombol Add a new diagram (+).", "1. Tambah Diagram: ")
    add_bullet_point(doc, "Isi Title: DFD Sistem Presensi QR. Pilih Diagram Type: STRIDE. Berikan deskripsi singkat: DFD Level-1 aliran data presensi perkuliahan.", "2. Konfigurasi Diagram: ")
    add_bullet_point(doc, "Klik Save lalu klik thumbnail diagram untuk membuka editor kanvas.", "3. Buka Kanvas: ")

    add_heading_2(doc, "4.3 Langkah 3: Menempatkan 8 Elemen DFD (Anggota 1)")
    add_body_paragraph(doc, "Tarik elemen dari panel sebelah kiri ke kanvas dan beri nama sesuai tabel:")
    
    elem_guide = doc.add_table(rows=9, cols=3)
    elem_g_headers = ["Bentuk di Tool", "Nama Elemen Wajib", "Posisi Relatif pada Kanvas"]
    elem_g_data = [
        ["Actor (External Entity)", "Mahasiswa (Browser)", "Kiri tengah kanvas (Zona TB1)"],
        ["Actor (External Entity)", "Dosen Pengampu (Browser)", "Kiri atas kanvas (Zona TB1)"],
        ["Actor (External Entity)", "Admin Akademik (Browser)", "Kiri bawah kanvas (Zona TB1)"],
        ["Actor (External Entity)", "Layanan Autentikasi Kampus", "Paling bawah tengah (Zona TB-Auth)"],
        ["Process", "P1. Kelola Sesi & QR", "Tengah atas kanvas (Zona TB2)"],
        ["Process", "P2. Validasi & Catat Presensi", "Tengah tengah kanvas (Zona TB2)"],
        ["Process", "P3. Muat Rekap Kehadiran", "Tengah bawah kanvas (Zona TB2)"],
        ["Data Store", "Database Akademik", "Kanan tengah kanvas (Zona TB3)"]
    ]
    format_table(elem_guide, [1.8, 2.3, 2.4], elem_g_headers, elem_g_data)

    add_heading_2(doc, "4.4 Langkah 4: Menghubungkan 16 Aliran Data (Anggota 1)")
    add_body_paragraph(doc, "Tarik elemen Data Flow (panah) dari titik port elemen asal ke elemen tujuan, lalu klik panah untuk mengisi nama/label:")
    add_bullet_point(doc, "F1: Dosen → P1 ('Buka/Tutup Sesi (Token Dosen)') | F2: P1 → Dosen ('Token QR Sesi')")
    add_bullet_point(doc, "F3: Mahasiswa → P2 ('QR Token + Token Mahasiswa') | F4: P2 → Mahasiswa ('Status Presensi')")
    add_bullet_point(doc, "F5: Klien → Layanan Autentikasi ('Kredensial') | F6: Layanan Autentikasi → Klien ('Token SSO')")
    add_bullet_point(doc, "F7: Backend → Layanan Autentikasi ('Verifikasi Token') | F8: Layanan Autentikasi → Backend ('Hasil Verifikasi')")
    add_bullet_point(doc, "F9: P1 → Database ('Tulis Sesi & QR') | F10: Database → P1 ('Jadwal & Dosen')")
    add_bullet_point(doc, "F11: P2 → Database ('Validasi & Tulis Presensi') | F12: Database → P2 ('Status Sesi & Hasil')")
    add_bullet_point(doc, "F13: Dosen/Admin → P3 ('Permintaan Rekap') | F14: P3 → Database ('Query Rekap')")
    add_bullet_point(doc, "F15: Database → P3 ('Data Rekap') | F16: P3 → Dosen/Admin ('Respons Rekapitulasi')")

    add_heading_2(doc, "4.5 Langkah 5: Menggambar 4 Batas Kepercayaan (Anggota 1)")
    add_body_paragraph(doc, "Tarik elemen Trust Boundary Box (garis putus-putus) untuk membungkus komponen:")
    add_bullet_point(doc, "TB1: Perangkat Pengguna / Internet (Membungkus Mahasiswa, Dosen, Admin).")
    add_bullet_point(doc, "TB-Auth: Layanan Autentikasi Kampus (Membungkus elemen SSO).")
    add_bullet_point(doc, "TB2: Backend Presensi (Membungkus proses P1, P2, P3).")
    add_bullet_point(doc, "TB3: Zona Data Internal (Membungkus Database Akademik).")
    add_body_paragraph(doc, "Setelah selesai, klik Save. Unduh dan bagikan KasusA_KelompokXX_v1_DFD.json ke Drive bersama.")

    # 4.6 & 4.7 INPUT ANCAMAN (SALIN TEMPEL)
    add_heading_2(doc, "4.6 Langkah 6: Input Entri Ancaman T01 s/d T07 (Anggota 2)")
    add_body_paragraph(doc, "Buka file KasusA_KelompokXX_v1_DFD.json di Threat Dragon. Klik elemen yang ditentukan, pilih tombol New Threat (+), lalu isi kolom sesuai tabel berikut:")

    t_input_1 = doc.add_table(rows=8, cols=5)
    t_in1_headers = ["ID", "Klik Elemen", "Tipe STRIDE", "Severity", "Judul Ancaman & Ringkasan Mitigasi"]
    t_in1_data = [
        ["T01", "Mahasiswa", "Spoofing", "High", "Judul: Titip absen dengan akun mahasiswa lain\nMitigasi: MFA pada akun kampus, pengikatan sesi ke perangkat, deteksi multi-presensi."],
        ["T02", "F3 (Aliran Data)", "Spoofing", "High", "Judul: Token autentikasi mahasiswa dicuri dan dipakai ulang\nMitigasi: Token ber-TTL singkat, cookie HttpOnly & Secure, token revocation."],
        ["T03", "P2 (Validasi)", "Spoofing", "High", "Judul: QR dibagikan jarak jauh ke mahasiswa tidak hadir\nMitigasi: Dynamic rolling QR (15-30 detik berbasis TOTP), batasan Wi-Fi kelas."],
        ["T04", "P1 (Kelola Sesi)", "Spoofing", "Medium", "Judul: Token QR dapat ditebak atau diprediksi\nMitigasi: Generator CSPRNG minimal 128-bit acak, rate limiting per IP."],
        ["T05", "F3 (Aliran Data)", "Tampering", "High", "Judul: Manipulasi ID sesi atau ID mahasiswa pada presensi\nMitigasi: Ekstraksi identitas eksklusif dari token terverifikasi, validasi relasi server."],
        ["T06", "P2 (Validasi)", "Tampering", "Medium", "Judul: Presensi ganda atau race condition waktu klien\nMitigasi: Unique constraint (id_sesi, id_mhs) database, transaksi atomik, jam server."],
        ["T07", "F11 (Aliran Data)", "Tampering", "High", "Judul: SQL injection melalui parameter presensi\nMitigasi: Parameterized query (prepared statements), validasi tipe input, least privilege."]
    ]
    format_table(t_input_1, [0.5, 1.3, 1.0, 0.8, 2.9], t_in1_headers, t_in1_data)
    add_body_paragraph(doc, "Setelah T01–T07 terinput, klik Save. Simpan sebagai KasusA_KelompokXX_v2_ST.json dan serahkan ke Anggota 3.")

    add_heading_2(doc, "4.7 Langkah 7: Input Entri Ancaman T08 s/d T14 (Anggota 3)")
    add_body_paragraph(doc, "Buka file KasusA_KelompokXX_v2_ST.json. Klik elemen terkait dan tambahkan ancaman T08–T14:")

    t_input_2 = doc.add_table(rows=8, cols=5)
    t_in2_headers = ["ID", "Klik Elemen", "Tipe STRIDE", "Severity", "Judul Ancaman & Ringkasan Mitigasi"]
    t_in2_data = [
        ["T08", "P2 (Validasi)", "Repudiation", "Medium", "Judul: Pengubahan atau penyangkalan catatan hadir tanpa audit log\nMitigasi: Centralized tamper-evident audit log, penanda waktu NTP, log IP & aktor."],
        ["T09", "F16 (Aliran Data)", "Info Disclosure", "Medium", "Judul: Rekap memuat data pribadi berlebih\nMitigasi: Minimalisasi data (data minimization DTO), pembatasan akses data berhak."],
        ["T10", "Database", "Info Disclosure", "Medium", "Judul: Akses langsung atau kebocoran cadangan basis data\nMitigasi: Enkripsi data saat istirahat (TDE), enkripsi backup, isolasi bastion host."],
        ["T11", "P2 (Validasi)", "Denial of Service", "High", "Judul: Lonjakan permintaan presensi serentak\nMitigasi: Message queue asynchronous, caching in-memory Redis, rate limiting."],
        ["T12", "Layanan SSO", "Denial of Service", "High", "Judul: Layanan Autentikasi Kampus lumpuh\nMitigasi: Verifikasi token lokal via JWKS public key cache, SOP presensi darurat."],
        ["T13", "P1 (Kelola Sesi)", "Elevation of Priv", "High", "Judul: Dosen non-pengampu membuka sesi kelas\nMitigasi: Strict object-level ownership check di server, deny-by-default."],
        ["T14", "P3 (Muat Rekap)", "Elevation of Priv", "High", "Judul: Akses data rekapitulasi lintas kelas\nMitigasi: Pemisahan endpoint dosen vs admin, filter query berbasis wewenang."]
    ]
    format_table(t_input_2, [0.5, 1.3, 1.1, 0.8, 2.8], t_in2_headers, t_in2_data)
    add_body_paragraph(doc, "Setelah selesai, klik Save. Simpan sebagai KasusA_KelompokXX_v3_Semua.json dan serahkan ke Anggota 4.")

    add_heading_2(doc, "4.8 Langkah 8: Quality Assurance dan Ekspor PDF (Anggota 4)")
    add_body_paragraph(doc, "Anggota 4 melakukan peninjauan menyeluruh terhadap file model v3:")
    add_bullet_point(doc, "Pastikan total ancaman tepat berjumlah 14 entri, seluruhnya berstatus 'Open', dan memiliki nilai Severity yang tepat (9 High, 5 Medium).", "1. Verifikasi Jumlah: ")
    add_bullet_point(doc, "Periksa bahwa setiap ancaman memiliki deskripsi skenario berpola 'Jika... maka...', serta rekomendasi mitigasi spesifik.", "2. Verifikasi Deskripsi: ")
    add_bullet_point(doc, "Simpan model final dengan nama KasusA_KelompokXX_FINAL.json.", "3. Simpan Final: ")
    add_bullet_point(doc, "Klik tombol Report pada detail model. Pada opsi laporan, pastikan memberi centang (checklist) pada:\n• Threat model diagrams\n• Empty model elements\n• Element properties\n• Mitigated threats dan Out of scope elements (jika ada)", "4. Pengaturan Laporan PDF: ")
    add_bullet_point(doc, "Jika menggunakan versi Desktop, klik PDF Report. Jika menggunakan versi Web, klik Print lalu pilih Save as PDF. Simpan dengan nama Laporan_ThreatModel_KasusA_KelompokXX.pdf.", "5. Simpan PDF: ")
    add_bullet_point(doc, "Tutup aplikasi Threat Dragon, lalu buka kembali file KasusA_KelompokXX_FINAL.json untuk memastikan file tidak rusak/korup.", "6. Uji Buka Ulang (Re-open Test): ")

    # 5. PANDUAN TANGKAPAN LAYAR (SCREENSHOT / CAPTURE)
    add_heading_1(doc, "5. Panduan Tangkapan Layar (Screenshot / Capture) Bukti Praktikum")
    add_body_paragraph(doc, 
        "Sebagai bukti autentik pelaksanaan praktikum pemodelan ancaman di OWASP Threat Dragon, setiap anggota tim wajib mengambil tangkapan layar (screenshot) sesuai tanggung jawab masing-masing. Tangkapan layar ini dilampirkan pada Bab 8 Dokumen Laporan Utama.")

    cap_table = doc.add_table(rows=7, cols=5)
    cap_headers = ["Kode", "Tampilan yang Wajib Di-Capture", "Elemen Kunci yang Harus Terlihat", "Momen Pengambilan", "Penanggung Jawab"]
    cap_data = [
        ["CAP-01", "Halaman Detail Model (Threat Model Details)", "Title model, Owner (nama tim), Reviewer (dosen), Description, dan daftar 4 Contributors beserta NIM.", "Setelah Langkah 1 (Inisialisasi Model)", "Fatir Gibran"],
        ["CAP-02", "Kanvas Utuh Diagram DFD Level-1", "Kanvas utuh memperlihatkan 8 elemen sistem, 16 aliran data berlabel jelas, dan 4 Trust Boundary berwarna.", "Setelah Langkah 5 (Penyusunan DFD Selesai)", "Fatir Gibran"],
        ["CAP-03", "Modal/Panel Rincian Ancaman (Spoofing/Tampering)", "Jendela dialog pengisian ancaman T01, T03, atau T05: Title, STRIDE Type, Severity (High/Med), Status Open, Skenario, dan Mitigasi.", "Saat pengisian T01–T07 di Langkah 6", "Tio Armani"],
        ["CAP-04", "Modal/Panel Rincian Ancaman (R, I, D, E)", "Jendela dialog pengisian ancaman T08, T11, atau T13: Title, STRIDE Type, Severity, Status Open, Skenario, dan Mitigasi.", "Saat pengisian T08–T14 di Langkah 7", "Chadafya Putra Z."],
        ["CAP-05", "Pengaturan Opsi Ekspor Laporan (Report Options)", "Halaman Report di Threat Dragon yang memperlihatkan centang aktif pada 'Threat model diagrams', 'Empty model elements', dan 'Element properties'.", "Sebelum klik Print/PDF Report di Langkah 8", "Tedy Permana Putra"],
        ["CAP-06", "Pratinjau / Hasil Ekspor Laporan PDF Tool", "Tampilan halaman pertama file PDF hasil generate tool yang memuat ringkasan ancaman dan visual diagram DFD.", "Setelah berhasil mengekspor PDF di Langkah 8", "Tedy Permana Putra"]
    ]
    format_table(cap_table, [0.8, 1.8, 2.1, 1.1, 1.1], cap_headers, cap_data)

    add_body_paragraph(doc, "Petunjuk Teknis Pengambilan Tangkapan Layar:")
    add_bullet_point(doc, "Gunakan format PNG berkualitas tinggi dengan rasio tampilan penuh (jangan dipotong hingga teks menjadi tidak terbaca).", "Resolusi: ")
    add_bullet_point(doc, "Simpan berkas capture dengan penamaan standar: Capture_01_Model_Detail.png, Capture_02_DFD_Kanvas.png, dst.", "Penamaan Berkas: ")
    add_bullet_point(doc, "Tempelkan gambar pada kotak lampiran yang telah disediakan pada Bab 8 Laporan_Threat_Modeling_Kasus_A_Presensi_QR.docx.", "Penempatan: ")

    # 6. PEMECAHAN MASALAH
    add_heading_1(doc, "6. Panduan Pemecahan Masalah (Troubleshooting)")
    add_body_paragraph(doc, "Berikut panduan solusi jika kelompok mengalami kendala teknis:")

    tb_sol = doc.add_table(rows=6, cols=3)
    tb_sol_headers = ["Gejala Masalah", "Penyebab Umum", "Solusi Penanganan Praktis"]
    tb_sol_data = [
        ["File .json tidak bisa dibuka di Threat Dragon", "File korup, format JSON rusak, atau pada macOS file berawalan '._' terunggah ke Drive.", "Buka file JSON menggunakan teks editor (VS Code/Notepad) untuk cek validitas sintaks. Pastikan memilih file utama, bukan file bayangan '._' macOS."],
        ["Garis panah Data Flow lepas saat digeser", "Ujung garis tidak menempel sempurna pada port elemen.", "Tarik ujung panah hingga lingkaran port elemen menyala biru/hijau, lalu lepaskan."],
        ["Entri ancaman tidak muncul di laporan PDF", "Checkbox 'Element properties' belum dicentang sebelum cetak.", "Buka kembali menu Report, centang seluruh opsi termasuk 'Element properties' dan 'Empty model elements', lalu cetak ulang."],
        ["Konflik versi file .json antaranggota", "Dua anggota mengedit model secara paralel di laptop berbeda.", "Tentukan file terbaru sebagai acuan, lalu ketik ulang entri ancaman yang tertinggal ke dalam satu file tunggal tersebut."],
        ["Elemen tidak bisa dipindahkan pada kanvas", "Elemen terkunci atau tertumpuk di bawah Trust Boundary Box.", "Klik kanan Trust Boundary Box, pilih 'Send to back' (zIndex paling rendah) agar elemen proses dan aktor dapat diklik dengan leluasa."]
    ]
    format_table(tb_sol, [1.8, 2.0, 2.7], tb_sol_headers, tb_sol_data)

    # 7. DAFTAR CEK AKHIR
    add_heading_1(doc, "7. Daftar Cek Akhir Sebelum Pengumpulan (Pre-Submission Checklist)")
    add_body_paragraph(doc, "Pastikan seluruh item berikut telah tercentang sebelum mengunggah tugas ke sistem akademik:")

    chk_table = doc.add_table(rows=11, cols=3)
    chk_headers = ["Status", "Poin Verifikasi Kualitas", "Standar Kelayakan"]
    align_chk = {0: WD_ALIGN_PARAGRAPH.CENTER}
    chk_data = [
        ["[  ]", "Identitas Threat Model Lengkap", "Memuat Title, Owner, Reviewer, Description, dan 4 Contributors."],
        ["[  ]", "Struktur DFD Lengkap & Proporsional", "8 Elemen utama, 16 Data Flow berlabel jelas, dan 4 Trust Boundary terpasang rapi."],
        ["[  ]", "Jumlah Ancaman Terpenuhi", "Terdapat 14 entri ancaman (melebihi syarat minimal 6 dosen)."],
        ["[  ]", "Cakupan Kerangka STRIDE Lengkap", "Memuat seluruh kategori: Spoofing, Tampering, Repudiation, Info Disc, DoS, EoP."],
        ["[  ]", "Format Penulisan Skenario Baku", "Mengikuti pola kausalitas: 'Jika [kondisi/aksi] pada [elemen], maka [dampak]...'"],
        ["[  ]", "Rekomendasi Mitigasi Spesifik", "Mitigasi berupa kontrol teknis perancangan, bukan sekadar kalimat umum."],
        ["[  ]", "Penilaian Risiko Objektif", "Skor dihitung dari Kemungkinan × Dampak dengan Severity sesuai rubrik (9 High, 5 Med)."],
        ["[  ]", "Status Seluruh Ancaman 'Open'", "Status tetap Open karena mitigasi masih berupa rancangan kontrol."],
        ["[  ]", "Uji Buka Ulang File .json Berhasil", "File JSON final sukses dibuka kembali di aplikasi Threat Dragon tanpa error."],
        ["[  ]", "Konsistensi Laporan Word & PDF", "Data pada Laporan Word (.docx) identik dengan hasil ekspor Report PDF."]
    ]
    format_table(chk_table, [0.8, 2.7, 3.0], chk_headers, chk_data, align_chk)

    output_path = "Panduan_Tata_Langkah_Pengerjaan_Kasus_A.docx"
    doc.save(output_path)
    print(f"File panduan langkah berhasil disimpan di {output_path}")

if __name__ == "__main__":
    build_panduan()
