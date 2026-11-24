# -*- coding: utf-8 -*-
"""
Script to generate Dokumen_Briefing_Presentasi_Kasus_A.docx
A comprehensive presentation briefing document and speaker script for the cybersecurity team.
Adheres strictly to EYD V, formal Indonesian academic standards, full-width bordered tables, and justified prose.
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

def add_callout(doc, text, title="TIPS PEMBICARA", total_dxa=9360):
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

    set_table_borders(table, border_color="2B6CB0", inside_h_color="2B6CB0", inside_v_color="2B6CB0")

    for row in table.rows:
        r_trPr = row._tr.get_or_add_trPr()
        r_trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

    cell = table.cell(0, 0)
    set_cell_background(cell, "EBF8FF")
    set_cell_margins(cell, top=90, bottom=90, left=130, right=130)
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
    rt.font.color.rgb = RGBColor(43, 108, 176)
    rb = p.add_run(text)
    rb.font.size = Pt(9)
    rb.font.color.rgb = RGBColor(45, 55, 72)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def build_briefing_doc():
    doc = docx.Document()

    # Page Margins (1.0 inch all sides)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Document Header Title
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(2)
    run_t = title_p.add_run("DOKUMEN BRIEFING & SKRIP PRESENTASI AKADEMIK")
    run_t.font.name = "Calibri"
    run_t.font.size = Pt(16)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor(26, 54, 93)

    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_after = Pt(12)
    run_s = sub_p.add_run("Panduan Operasional Pemaparan Tim, Alokasi Menit, Skrip Pembicara Slide 1–13, dan Bank Tanya Jawab Kritis Dosen\n"
                          "Mata Kuliah: Keamanan Siber (CAK3CAB3) • Studi Kasus A: Sistem Presensi Perkuliahan Berbasis QR Code")
    run_s.font.name = "Calibri"
    run_s.font.size = Pt(10)
    run_s.font.color.rgb = RGBColor(113, 128, 150)

    # ==========================
    # BAB 1
    # ==========================
    add_heading_1(doc, "1. Gambaran Umum Presentasi dan Manajemen Waktu")
    add_body_paragraph(doc,
        "Dokumen briefing ini disusun sebagai acuan operasional resmi bagi seluruh anggota kelompok dalam membawakan presentasi evaluasi threat modeling sistem presensi perkuliahan berbasis QR code. Presentasi ini didesain ringkas, padat data, dan berfokus pada ketepatan arsitektural di hadapan dosen pengampu.")
    
    add_heading_2(doc, "1.1 Roster Anggota Tim dan Pembagian Peran Presentasi")
    roster_table = doc.add_table(rows=5, cols=4)
    roster_headers = ["Nama Anggota", "NIM", "Peran Inti Proyek", "Bagian Presentasi"]
    roster_data = [
        ["Fatir Gibran", "103112430153", "Ketua Tim & Lead Architect", "Slide 1–6 (Fondasi & DFD), Slide 10 (Bukti DFD), Slide 12–13 (Kesimpulan & Q&A)"],
        ["Tio Armani", "103112430225", "Analis Spoofing & Tampering", "Slide 7 (Ancaman T01–T07: Pemalsuan & Manipulasi Token QR/Sesi)"],
        ["Chadafya Putra Zulfikar", "103112430173", "Analis R, I, D, E", "Slide 8 (Ancaman T08–T14: Penyangkalan, Kebocoran, DoS, Eskalasi Hak)"],
        ["Tedy Permana Putra", "103112430157", "Penilai Risiko & QA Tool", "Slide 9 (Matriks Risiko & Prioritas), Slide 11 (Bukti Validasi Tool & Laporan)"]
    ]
    format_table(roster_table, [1.5, 1.2, 1.8, 2.0], roster_headers, roster_data)

    add_heading_2(doc, "1.2 Alokasi Durasi dan Manajemen Waktu (Total: 15 Menit)")
    add_body_paragraph(doc,
        "Alokasi waktu standar presentasi praktikum adalah 15 menit, yang dibagi secara proporsional menjadi 10 menit pemaparan slide (rata-rata 45–50 detik per slide) dan 5 menit sesi tanya jawab kritis dengan dosen:")

    time_table = doc.add_table(rows=6, cols=4)
    time_headers = ["Segmen Waktu", "Alokasi Slide", "Pembicara Bertugas", "Fokus Bahasan"]
    time_data = [
        ["Menit 00:00 – 03:00", "Slide 1 – 6", "Fatir Gibran", "Pembuka, Ruang Lingkup Kasus A, DFD Level-1, Trust Boundary, & STRIDE"],
        ["Menit 03:00 – 05:00", "Slide 7", "Tio Armani", "Register Ancaman Spoofing (T01-T04) dan Tampering (T05-T07)"],
        ["Menit 05:00 – 07:00", "Slide 8", "Chadafya Putra Z.", "Register Ancaman Repudiation (T08), Info Disclosure (T09-T10), DoS (T11-T12), & EoP (T13-T14)"],
        ["Menit 07:00 – 08:30", "Slide 9 & 11", "Tedy Permana P.", "Matriks Risiko 3x3, Prioritas Remediasi P1-P3, dan Bukti Ekspor Laporan PDF"],
        ["Menit 08:30 – 10:00", "Slide 10 & 12-13", "Fatir Gibran", "Bukti Model Threat Dragon, Kesimpulan Rekayasa Arsitektur, & Membuka Q&A"]
    ]
    format_table(time_table, [1.3, 1.1, 1.5, 2.6], time_headers, time_data)

    add_callout(doc,
        "Jangan membaca teks slide secara kata-per-kata! Slide presentasi dirancang sebagai jangkar visual bagi audiens. Pembicara wajib menjelaskan konteks 'mengapa' ancaman tersebut terjadi dan 'bagaimana' solusi teknis arsitekturalnya.",
        "DISIPLIN WAKTU PRESENTASI")

    # ==========================
    # BAB 2
    # ==========================
    add_heading_1(doc, "2. Naskah Penjelasan Terperinci per Slide (Slide 1–13)")
    add_body_paragraph(doc,
        "Bagian ini memuat poin kunci visual, skrip bicara formal berbahasa Indonesia baku (EYD V), serta antisipasi pertanyaan dosen untuk setiap slide presentasi.")

    # Slide 1
    add_heading_2(doc, "Slide 1: Judul & Identitas Kelompok (Pembicara: Fatir Gibran)")
    add_bullet_point(doc, "Memperkenalkan judul tugas, mata kuliah Keamanan Siber (CAK3CAB3), dan keempat anggota kelompok beserta peran teknis masing-masing.", "Poin Kunci Visual: ")
    add_bullet_point(doc, "\"Selamat pagi/siang Bapak/Ibu dosen dan rekan-rekan sekalian. Hari ini kami dari kelompok pemodelan keamanan siber akan mempresentasikan hasil analisis Threat Modeling berbasis kerangka kerja STRIDE pada Studi Kasus A, yaitu Sistem Presensi Perkuliahan Mahasiswa Berbasis QR Code menggunakan perkakas OWASP Threat Dragon. Dalam proyek ini, saya Fatir Gibran bertindak sebagai Lead Architect, bersama rekan saya Tio Armani, Chadafya Putra Zulfikar, dan Tedy Permana Putra.\"", "Contoh Skrip Bicara: ")

    # Slide 2
    add_heading_2(doc, "Slide 2: Latar Belakang & Ruang Lingkup Sistem Presensi QR (Pembicara: Fatir Gibran)")
    add_bullet_point(doc, "Menjelaskan fenomena titip absen, jendela waktu QR 5 menit yang statis, dan keterbatasan deteksi lokasi fisik mahasiswa.", "Poin Kunci Visual: ")
    add_bullet_point(doc, "\"Latar belakang analisis ini berakar dari kerentanan operasional pada sistem presensi QR konvensional. Mekanisme saat ini mengizinkan kode QR aktif selama 5 menit penuh. Jendela waktu yang panjang ini membuka peluang eksploitasi nyata: mahasiswa di kelas memfoto kode QR dan menyebarkannya melalui grup perpesanan ke mahasiswa di luar kelas. Ditambah lagi, validasi lokasi hanya mengandalkan koordinat GPS mentah dari gawai klien yang sangat rentan dimanipulasi dengan mock location. Ruang lingkup kami mencakup analisis aliran data dari aplikasi mobile mahasiswa, portal dosen, gerbang API, hingga basis data akademik.\"", "Contoh Skrip Bicara: ")

    # Slide 3
    add_heading_2(doc, "Slide 3: Aturan Bisnis & 5 Asumsi Teknis Perancangan (Pembicara: Fatir Gibran)")
    add_bullet_point(doc, "Menegaskan 4 aturan bisnis utama (interval 5 menit, jaringan kampus, dependensi SSO, isolasi rekap) dan 5 asumsi batas analisis.", "Poin Kunci Visual: ")
    add_bullet_point(doc, "\"Untuk menjaga objektivitas analisis, kami menetapkan 5 asumsi teknis perancangan: pertama, gawai mahasiswa diasumsikan berada pada lingkungan yang tidak tepercaya (untrusted). Kedua, transmisi data diasumsikan belum menegakkan inspeksi mTLS penuh. Ketiga, server SSO kampus dikelola secara eksternal sehingga dianalisis sebagai ketergantungan pihak ketiga. Keempat, basis data akademik menyimpan catatan kehadiran bernilai tinggi. Dan kelima, proses bisnis membatasi presensi hanya sah saat jadwal kuliah aktif.\"", "Contoh Skrip Bicara: ")

    # Slide 4
    add_heading_2(doc, "Slide 4: Arsitektur Data Flow Diagram (DFD) Level-1 (Pembicara: Fatir Gibran)")
    add_bullet_point(doc, "Menampilkan diagram arsitektur utuh: 4 aktor (E1–E4), 3 proses (P1–P3), 1 datastore (DS1), dan 16 aliran data (F1–F16).", "Poin Kunci Visual: ")
    add_bullet_point(doc, "\"Diagram Aliran Data Level-1 ini memetakan seluruh interaksi sistem. Terdapat tiga proses komputasi utama: P1 Kelola Sesi & Pembangkitan QR oleh Dosen (E2), P2 Validasi & Pencatatan Presensi oleh Mahasiswa (E1), serta P3 Muat Rekap Kehadiran untuk Admin Akademik (E3). Aliran data F1 hingga F16 memperlihatkan pertukaran token autentikasi SSO (E4), token QR dinamis, parameter verifikasi, hingga penyimpanan status ke Database Akademik (DS1).\"", "Contoh Skrip Bicara: ")

    # Slide 5
    add_heading_2(doc, "Slide 5: Zona Keamanan & 4 Batas Kepercayaan (Trust Boundaries) (Pembicara: Fatir Gibran)")
    add_bullet_point(doc, "Menerangkan batas TB1 (Perangkat Pengguna/Internet), TB-Auth (SSO Kampus), TB2 (Backend Presensi Internal), dan TB3 (Zona Data Internal).", "Poin Kunci Visual: ")
    add_bullet_point(doc, "\"Fondasi analisis STRIDE bertumpu pada 4 Batas Kepercayaan atau Trust Boundaries. TB1 mewakili lingkungan internet publik tempat gawai mahasiswa dan peramban dosen berada. Ketika aliran data menyeberang dari TB1 ke TB2 (jaringan backend internal kampus), di sinilah perimeter keamanan pertama kali diuji. Kami juga memisahkan TB-Auth untuk layanan SSO kampus, serta TB3 sebagai zona penyimpanan data akademik terisolasi yang tidak boleh diakses langsung oleh publik.\"", "Contoh Skrip Bicara: ")

    # Slide 6
    add_heading_2(doc, "Slide 6: Metodologi Analisis STRIDE & Rubrik Risiko (Pembicara: Fatir Gibran)")
    add_bullet_point(doc, "Menerangkan 6 kategori STRIDE dan formula penilaian risiko (Kemungkinan 1–3 x Dampak 1–3 = Skor 1–9). Handover ke Tio.", "Poin Kunci Visual: ")
    add_bullet_point(doc, "\"Dalam mengidentifikasi ancaman, kami menerapkan metodologi STRIDE dengan rubrik matriks evaluasi 3x3 standar OWASP. Tingkat kemungkinan bernilai 1 hingga 3, dan dampak bernilai 1 hingga 3, menghasilkan skor risiko 1 sampai 9. Skor 6 hingga 9 dikategorikan Tinggi (High), dan skor 3 hingga 4 dikategorikan Sedang (Medium). Selanjutnya, rekan saya Tio Armani akan memaparkan temuan ancaman pada kategori Spoofing dan Tampering.\"", "Contoh Skrip Bicara: ")

    # Slide 7
    add_heading_2(doc, "Slide 7: Register Ancaman: Spoofing & Tampering (T01–T07) (Pembicara: Tio Armani)")
    add_bullet_point(doc, "Rincian ancaman T01–T04 (Spoofing) dan T05–T07 (Tampering), dengan penekanan pada ancaman berisiko tinggi seperti joki akun dan pemalsuan QR.", "Poin Kunci Visual: ")
    add_bullet_point(doc, "\"Terima kasih Fatir. Pada kategori Spoofing, ancaman paling kritis adalah T01 (Titip Absen Akun Lain) dan T03 (Penyebaran QR Jarak Jauh). Keduanya memiliki skor 6 (Tinggi). Jika mahasiswa membagikan tangkapan layar kode QR ke grup chat, mahasiswa di luar kampus dapat memindainya dan tercatat hadir secara curang. Rekomendasi mitigasi kami adalah menerapkan TOTP QR dinamis dengan masa kedaluwarsa pendek (30 detik), dipadukan dengan validasi BSSID Wi-Fi kampus serta pengikatan identitas perangkat keras (device fingerprinting). Pada kategori Tampering, ancaman T05 (Manipulasi Parameter ID Mahasiswa) dan T07 (SQL Injection pada Aliran F11) juga berisiko Tinggi. Mitigasinya adalah penandatanganan payload dengan JWT bertanda tangan kriptografis dan penegakan prepared statements serta Object-Relational Mapping (ORM) pada basis data.\"", "Contoh Skrip Bicara: ")

    # Slide 8
    add_heading_2(doc, "Slide 8: Register Ancaman: R, I, D, E (T08–T14) (Pembicara: Chadafya Putra Zulfikar)")
    add_bullet_point(doc, "Rincian ancaman T08 (Repudiation), T09–T10 (Info Disclosure), T11–T12 (Denial of Service), dan T13–T14 (Elevation of Privilege). Handover ke Tedy.", "Poin Kunci Visual: ")
    add_bullet_point(doc, "\"Terima kasih Tio. Saya Chadafya akan melanjutkan pada empat kategori STRIDE berikutnya. Pada kategori Repudiation, ancaman T08 mencakup pengubahan catatan hadir tanpa jejak audit (Skor 4). Kami merekomendasikan audit log tersinkronisasi WORM (Write Once Read Many). Pada Information Disclosure, ancaman T10 (Akses Tidak Sah Basis Data Akademik) berisiko Sedang (Skor 3), dimitigasi dengan enkripsi kolom sensitif (AES-256). Pada Denial of Service, ancaman T11 (Lonjakan Presensi Serentak pada Aliran F3) berisiko Tinggi (Skor 6) akibat potensi server hang di awal jam kuliah; solusinya adalah antrean asinkron Redis dan pembatasan laju request (rate limiting). Terakhir, pada Elevation of Privilege, ancaman T13 (Dosen Non-Pengampu Membuka Sesi) dan T14 (Akses Rekapitulasi Lintas Kelas) keduanya berisiko Tinggi (Skor 6), yang dimitigasi secara ketat melalui penegakan Role-Based Access Control (RBAC) pada backend P1 dan P3. Selanjutnya, rekan saya Tedy Permana Putra akan menjelaskan evaluasi risiko.\"", "Contoh Skrip Bicara: ")

    # Slide 9
    add_heading_2(doc, "Slide 9: Matriks Evaluasi Risiko & Prioritas Remediasi (Pembicara: Tedy Permana Putra)")
    add_bullet_point(doc, "Peta sebaran matriks 3x3: 9 High (merah) dan 5 Medium (oranye), serta urutan prioritas mitigasi P1 (Kritis), P2 (Signifikan), dan P3 (Operasional).", "Poin Kunci Visual: ")
    add_bullet_point(doc, "\"Terima kasih Chadafya. Berdasarkan evaluasi risiko matriks 3x3, dari total 14 ancaman yang diidentifikasi, terdapat 9 ancaman berkategori Tinggi (Skor 6) dan 5 ancaman berkategori Sedang (Skor 3–4). Tidak ada ancaman berstatus Rendah karena seluruhnya bersentuhan langsung dengan integritas akademik. Kami menyusun strategi remediasi ke dalam tiga tingkat prioritas: Prioritas 1 (P1) menangani keabsahan fisik kehadiran dan perlindungan database (T01, T03, T05, T07). Prioritas 2 (P2) menangani integritas otorisasi dan ketersediaan layanan presensi (T02, T11, T12, T13, T14). Dan Prioritas 3 (P3) memperkuat jejak audit dan kerahasiaan data mahasiswa (T04, T06, T08, T09, T10).\"", "Contoh Skrip Bicara: ")

    # Slide 10
    add_heading_2(doc, "Slide 10: Bukti Praktikum 1: Detail Model & DFD di Threat Dragon (Pembicara: Fatir Gibran)")
    add_bullet_point(doc, "Menampilkan tangkapan layar CAP-01 (Metadata model & 4 kontributor) dan CAP-02 (Kanvas visual DFD level-1 resmi tool).", "Poin Kunci Visual: ")
    add_bullet_point(doc, "\"Sebagai bukti pelaksanaan praktikum mandiri, Slide 10 menampilkan tangkapan layar langsung dari perangkat lunak OWASP Threat Dragon v2. Pada panel kiri (CAP-01), tercatat judul model STRIDE - Kasus A - Presensi QR, dengan owner dan kontributor resmi seluruh anggota tim lengkap beserta NIM. Pada panel kanan (CAP-02), terlihat kanvas visual DFD Level-1 yang kami bangun dengan 8 elemen, 16 aliran data berarah, dan 4 batas kepercayaan yang memisahkan zona publik dari backend kampus.\"", "Contoh Skrip Bicara: ")

    # Slide 11
    add_heading_2(doc, "Slide 11: Bukti Praktikum 2: Dokumentasi Ancaman & Laporan PDF (Pembicara: Tedy Permana Putra)")
    add_bullet_point(doc, "Menampilkan tangkapan layar CAP-03 (Entri Tio), CAP-04 (Entri Chadafya), CAP-05 (Centang opsi Report), dan CAP-06 (Ekspor PDF resmi 9 halaman).", "Poin Kunci Visual: ")
    add_bullet_point(doc, "\"Pada Slide 11, kami mendokumentasikan bukti kontribusi masing-masing analis dalam menginput ancaman ke dalam Threat Dragon. Terlihat form entri ancaman Spoofing dan Tampering oleh Tio Armani (CAP-03), entri kategori Repudiation hingga Elevation of Privilege oleh Chadafya (CAP-04), serta pengaturan fitur laporan oleh saya sendiri (CAP-05). Kami memastikan opsi 'Show element properties' dicentang agar seluruh deskripsi, dampak, dan mitigasi tercetak utuh pada berkas ekspor PDF resmi 9 halaman (CAP-06).\"", "Contoh Skrip Bicara: ")

    # Slide 12
    add_heading_2(doc, "Slide 12: Kesimpulan & Rekomendasi Keamanan Arsitektural (Pembicara: Fatir Gibran)")
    add_bullet_point(doc, "Menyimpulkan 3 pilar utama: model DFD komprehensif, dominasi risiko integritas presensi, dan status Open sebagai rekomendasi arsitektur pra-implementasi.", "Poin Kunci Visual: ")
    add_bullet_point(doc, "\"Sebagai kesimpulan, pemodelan ancaman pada Sistem Presensi QR membuktikan bahwa celah terbesar bukan terletak pada bug pemrograman, melainkan pada kelemahan desain arsitektur presensi yang terlalu permisif terhadap jendela waktu dan data lokasi klien. Rekomendasi teknis yang kami ajukan—yaitu TOTP dinamis, pengikatan BSSID jaringan kampus, inspeksi mTLS, dan penegakan RBAC ketat—menjadi cetak biru pertahanan sebelum sistem ini diimplementasikan ke lingkungan produksi. Seluruh ancaman berstatus Open karena perbaikan ini merupakan usulan kontrol perancangan.\"", "Contoh Skrip Bicara: ")

    # Slide 13
    add_heading_2(doc, "Slide 13: Sesi Tanya Jawab (Q&A) & Penutup (Seluruh Anggota Tim)")
    add_bullet_point(doc, "Menutup presentasi secara formal dan mengundang pertanyaan serta masukan dari dosen penguji.", "Poin Kunci Visual: ")
    add_bullet_point(doc, "\"Demikian presentasi laporan hasil analisis Threat Modeling dari kelompok kami. Terima kasih atas perhatian Bapak/Ibu dosen dan rekan-rekan sekalian. Kami membuka sesi tanya jawab untuk mendiskusikan temuan dan rekomendasi arsitektur ini lebih lanjut. Waktu dan tempat kami kembalikan kepada Bapak/Ibu dosen.\"", "Contoh Skrip Bicara: ")

    # ==========================
    # BAB 3
    # ==========================
    add_heading_1(doc, "3. Bank Pertanyaan Kritis Dosen & Panduan Jawaban (Defense Guide)")
    add_body_paragraph(doc,
        "Dosen penguji keamanan siber lazimnya menguji pemahaman konseptual metodologi threat modeling dan ketepatan teknis kontrol keamanan. Berikut adalah 5 prediksi pertanyaan kritis beserta strategi jawaban terbaik:")

    # Q1
    add_heading_2(doc, "Pertanyaan 1: Mengapa status seluruh 14 ancaman masih 'Open', bukan 'Mitigated'?")
    add_callout(doc,
        "Jawaban Kunci: Status ancaman dalam kerangka kerja OWASP Threat Dragon mencerminkan status implementasi aktual pada sistem. Karena tugas praktikum ini berada pada tahap analisis desain arsitektur (Threat Modeling Activity) tanpa melibatkan perubahan kode sumber aplikasi yang sedang berjalan, rekomendasi mitigasi yang kami rumuskan masih berstatus usulan perancangan teknis. Status 'Mitigated' baru sah diberikan apabila kontrol keamanan telah selesai dikoding, diuji coba (*security testing*), dan diverifikasi efektivitasnya.",
        "JAWABAN STANDAR AKADEMIK (DIJAWAB OLEH FATIR ATAU TEDY)")

    # Q2
    add_heading_2(doc, "Pertanyaan 2: Mengapa aktor Dosen (E2) dan Admin Akademik (E3) tidak memiliki entri ancaman langsung pada kanvas diagram?")
    add_callout(doc,
        "Jawaban Kunci: Dalam metodologi pemodelan DFD, entitas eksternal (Aktor) hanya bertindak sebagai sumber input atau penerima output di luar kendali langsung aplikasi. Kerentanan dan penyalahgunaan hak akses dosen maupun admin terjadi pada saat proses sistem mengeksekusi instruksi mereka. Oleh karena itu, potensi ancaman dosen non-pengampu membuka sesi kelas dipusatkan pada proses P1 (Kelola Sesi & Pembangkitan QR / Ancaman T13), dan potensi admin memanipulasi rekap presensi dipusatkan pada proses P3 (Muat Rekap Kehadiran / Ancaman T14). Hal ini mencegah redundansi dokumentasi sesuai panduan resmi OWASP Threat Dragon.",
        "JAWABAN STANDAR AKADEMIK (DIJAWAB OLEH FATIR ATAU CHADAFYA)")

    # Q3
    add_heading_2(doc, "Pertanyaan 3: Mengapa ancaman T01 (Titip Absen) dan T03 (Share QR) ditautkan ke proses P2 dan aliran data F3, bukan ke Batas Kepercayaan (Trust Boundary)?")
    add_callout(doc,
        "Jawaban Kunci: Secara arsitektur perangkat lunak Threat Dragon, objek 'Trust Boundary' berfungsi sebagai pembatas visual zona kepercayaan dan tidak memiliki wadah atribut untuk menampung entri ancaman. Ancaman harus selalu ditautkan pada elemen operasional yang memproses atau mentransmisikan data. Ancaman titip absen terjadi ketika aliran data permintaan presensi (F3) dikirim dari perangkat tidak tepercaya (TB1) menuju proses validasi backend (P2 pada TB2). Menautkan T01 dan T03 pada P2 dan F3 memberikan kejelasan konteks teknis di mana fungsi validasi identitas dan pemeriksaan token presensi harus diperkuat.",
        "JAWABAN STANDAR AKADEMIK (DIJAWAB OLEH TIO ARMANI)")

    # Q4
    add_heading_2(doc, "Pertanyaan 4: Bagaimana dynamic QR code 30 detik mencegah titip absen jika mahasiswa menggunakan panggilan video langsung (live screen sharing)?")
    add_callout(doc,
        "Jawaban Kunci: Dynamic QR code dengan algoritma TOTP 30 detik secara signifikan memangkas jendela waktu eksploitasi, namun kontrol tersebut tidak berdiri sendiri. Kami merumuskan mekanisme pertahanan berlapis (defense-in-depth). Selain token QR yang berubah setiap 30 detik, backend P2 memvalidasi BSSID (alamat fisik access point Wi-Fi kelas) tempat gawai mahasiswa terhubung, serta memverifikasi penanda perangkat keras (device fingerprinting). Sekalipun kode QR berhasil di-share secara real-time via video call, mahasiswa di luar kampus akan gagal melakukan presensi karena gawai mereka terhubung ke jaringan internet publik atau BTS seluler di luar jangkauan access point kelas terkait.",
        "JAWABAN STANDAR AKADEMIK (DIJAWAB OLEH TIO ATAU FATIR)")

    # Q5
    add_heading_2(doc, "Pertanyaan 5: Apa urgensi memisahkan Database Akademik (DS1) ke dalam Trust Boundary terpisah (TB3)?")
    add_callout(doc,
        "Jawaban Kunci: Pemisahan Database Akademik ke dalam TB3 (Zona Data Internal) menerapkan prinsip pemisahan tugas (separation of concerns) dan pertahanan berlapis (network segmentation). Jika peladen backend presensi pada TB2 berhasil dikompromikan oleh penyerang melalui celah web application, penyerang tidak secara otomatis memperoleh akses langsung ke basis data induk kampus. Komunikasi antara TB2 dan TB3 diwajibkan melalui jaringan privat terisolasi dengan pembatasan hak akses berbasis least privilege, serta penegakan firewall internal yang menolak seluruh koneksi langsung dari internet publik.",
        "JAWABAN STANDAR AKADEMIK (DIJAWAB OLEH CHADAFYA ATAU TEDY)")

    # ==========================
    # BAB 4
    # ==========================
    add_heading_1(doc, "4. Protokol Teknis, Transisi Antar-Pembicara, dan Etika Presentasi")
    add_body_paragraph(doc,
        "Keberhasilan presentasi kelompok sangat dipengaruhi oleh kelancaran perpindahan giliran bicara (*handover cues*) dan kesiapan berkas teknis pendukung.")

    add_heading_2(doc, "4.1 Kalimat Transisi Pembicara (Handover Cues)")
    add_bullet_point(doc, "\"Setelah kita memahami arsitektur DFD dan rubrik risiko, selanjutnya analisis mendalam untuk kategori Spoofing dan Tampering akan disampaikan oleh rekan saya, Tio Armani. Silakan Tio.\"", "Transisi dari Fatir ke Tio (Akhir Slide 6): ")
    add_bullet_point(doc, "\"Sekian pemaparan ancaman pemalsuan dan manipulasi. Untuk empat kategori berikutnya, yaitu Repudiation, Information Disclosure, Denial of Service, dan Elevation of Privilege, akan dijelaskan oleh rekan saya Chadafya Putra Zulfikar. Silakan Chadafya.\"", "Transisi dari Tio ke Chadafya (Akhir Slide 7): ")
    add_bullet_point(doc, "\"Terima kasih. Selanjutnya, evaluasi pemetaan risiko pada matriks 3x3 dan prioritas mitigasi sistem akan dipaparkan oleh rekan saya Tedy Permana Putra. Silakan Tedy.\"", "Transisi dari Chadafya ke Tedy (Akhir Slide 8): ")
    add_bullet_point(doc, "\"Setelah melihat matriks risiko, bukti praktikum langsung pada aplikasi Threat Dragon dan kesimpulan akhir akan dirangkum kembali oleh ketua tim kami, Fatir Gibran. Silakan Fatir.\"", "Transisi dari Tedy ke Fatir (Akhir Slide 11): ")

    add_heading_2(doc, "4.2 Kesiapan Berkas Pendukung di Laptop Presenter (File Checklist)")
    add_body_paragraph(doc,
        "Pastikan sebelum sesi presentasi dimulai, presenter telah membuka aplikasi dan tab berikut di latar belakang (*background tabs*) agar siap ditampilkan saat diminta dosen:")
    
    check_table = doc.add_table(rows=5, cols=3)
    check_headers = ["Nama Berkas / Aplikasi", "Format Berkas", "Fungsi saat Sesi Q&A"]
    check_data = [
        ["Presentasi_Threat_Modeling_Kasus_A_Presensi_QR.pptx", "PowerPoint 16:9", "Tampilan utama yang ditayangkan ke proyektor / layar kelas."],
        ["KasusA_KelompokXX_FINAL.json", "Model JSON Threat Dragon", "Dibuka langsung pada antarmuka web/desktop Threat Dragon jika dosen meminta live demo diagram."],
        ["Laporan_Threat_Modeling_Kasus_A_Presensi_QR.docx", "Word Document (Bab 1–8)", "Rujukan lengkap jika dosen menanyakan detail narasi skenario, dampak, atau mitigasi tertentu."],
        ["Laporan_ThreatModel_KasusA_Presensi_QR.pdf", "PDF Resmi Ekspor Tool", "Bukti keaslian ekspor laporan Threat Dragon v2 dengan 14 entri ancaman lengkap."]
    ]
    format_table(check_table, [2.5, 1.5, 2.5], check_headers, check_data)

    output_path = "Dokumen_Briefing_Presentasi_Kasus_A.docx"
    doc.save(output_path)
    print(f"File dokumen briefing berhasil disimpan di {output_path}")

if __name__ == "__main__":
    build_briefing_doc()
