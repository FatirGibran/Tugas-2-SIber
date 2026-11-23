# -*- coding: utf-8 -*-
"""
Script to generate Laporan_Threat_Modeling_Kasus_A_Presensi_QR.docx
Following high academic standards, clean layout, complete sections, and strict anti-slop rules.
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, hex_color):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}>'
                      f'<w:top w:w="{top}" w:type="dxa"/>'
                      f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
                      f'<w:left w:w="{left}" w:type="dxa"/>'
                      f'<w:right w:w="{right}" w:type="dxa"/>'
                      f'</w:tcMar>')
    tcPr.append(tcMar)

def set_table_borders(table, color="D3D3D3"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'<w:tblBorders {nsdecls("w")}>'
                        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
                        f'<w:left w:val="none"/>'
                        f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="{color}"/>'
                        f'<w:right w:val="none"/>'
                        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
                        f'<w:insideV w:val="none"/>'
                        f'</w:tblBorders>')
    tblPr.append(borders)

def format_table(table, col_widths, headers, data, align_cols=None):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, "B0BEC5")
    
    # Header row
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        set_cell_background(hdr_cells[i], "1A365D")
        set_cell_margins(hdr_cells[i], top=120, bottom=120, left=140, right=140)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(9.5)
            run.font.name = "Calibri"

    # Data rows
    for r_idx, row_data in enumerate(data):
        row_cells = table.rows[r_idx + 1].cells
        bg_col = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            row_cells[c_idx].text = str(val)
            set_cell_background(row_cells[c_idx], bg_col)
            set_cell_margins(row_cells[c_idx], top=90, bottom=90, left=130, right=130)
            p = row_cells[c_idx].paragraphs[0]
            if align_cols and c_idx in align_cols:
                p.alignment = align_cols[c_idx]
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for run in p.runs:
                run.font.size = Pt(9)
                run.font.name = "Calibri"
                run.font.color.rgb = RGBColor(45, 55, 72)

    # Set column widths
    for row in table.rows:
        for i, w in enumerate(col_widths):
            row.cells[i].width = Inches(w)

def add_heading_1(doc, text):
    h = doc.add_paragraph()
    h.paragraph_format.space_before = Pt(16)
    h.paragraph_format.space_after = Pt(6)
    h.paragraph_format.keep_with_next = True
    run = h.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = RGBColor(26, 54, 93) # #1A365D
    return h

def add_heading_2(doc, text):
    h = doc.add_paragraph()
    h.paragraph_format.space_before = Pt(12)
    h.paragraph_format.space_after = Pt(4)
    h.paragraph_format.keep_with_next = True
    run = h.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = RGBColor(43, 108, 176) # #2B6CB0
    return h

def add_heading_3(doc, text):
    h = doc.add_paragraph()
    h.paragraph_format.space_before = Pt(8)
    h.paragraph_format.space_after = Pt(3)
    h.paragraph_format.keep_with_next = True
    run = h.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(10.5)
    run.font.bold = True
    run.font.color.rgb = RGBColor(45, 55, 72)
    return h

def add_body_paragraph(doc, text, bold_prefix="", space_after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = "Calibri"
        r_pre.font.size = Pt(10)
        r_pre.font.bold = True
        r_pre.font.color.rgb = RGBColor(45, 55, 72)
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(45, 55, 72)
    return p

def add_bullet_point(doc, text, bold_prefix=""):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = "Calibri"
        r_pre.font.size = Pt(10)
        r_pre.font.bold = True
        r_pre.font.color.rgb = RGBColor(45, 55, 72)
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(45, 55, 72)
    return p

def add_threat_card(doc, tid, title, stride_type, element, likelihood, impact, score_text, severity, scenario, impact_desc, assumption, mitigation):
    # Container table for threat card
    table = doc.add_table(rows=6, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, "CBD5E0")
    
    # Header row
    hdr_cell = table.cell(0, 0)
    hdr_cell.merge(table.cell(0, 1))
    badge_bg = "C53030" if severity == "Tinggi (High)" else "DD6B20"
    set_cell_background(hdr_cell, "2D3748")
    set_cell_margins(hdr_cell, top=100, bottom=100, left=140, right=140)
    
    hp = hdr_cell.paragraphs[0]
    hrun1 = hp.add_run(f"[{tid}] {title}")
    hrun1.font.bold = True
    hrun1.font.size = Pt(10.5)
    hrun1.font.color.rgb = RGBColor(255, 255, 255)
    
    # Row 1: Metadata Grid
    r1_c1 = table.cell(1, 0)
    r1_c2 = table.cell(1, 1)
    set_cell_background(r1_c1, "F7FAFC")
    set_cell_background(r1_c2, "F7FAFC")
    set_cell_margins(r1_c1, 70, 70, 120, 120)
    set_cell_margins(r1_c2, 70, 70, 120, 120)
    
    p1 = r1_c1.paragraphs[0]
    p1.add_run("Kategori STRIDE: ").font.bold = True
    p1.add_run(stride_type)
    p1.runs[0].font.size = Pt(9.5)
    p1.runs[1].font.size = Pt(9.5)
    
    p2 = r1_c2.paragraphs[0]
    p2.add_run("Elemen Terkait: ").font.bold = True
    p2.add_run(element)
    p2.runs[0].font.size = Pt(9.5)
    p2.runs[1].font.size = Pt(9.5)

    # Row 2: Risk & Status
    r2_c1 = table.cell(2, 0)
    r2_c2 = table.cell(2, 1)
    set_cell_background(r2_c1, "F7FAFC")
    set_cell_background(r2_c2, "F7FAFC")
    set_cell_margins(r2_c1, 70, 70, 120, 120)
    set_cell_margins(r2_c2, 70, 70, 120, 120)

    p3 = r2_c1.paragraphs[0]
    p3.add_run("Tingkat Risiko: ").font.bold = True
    r_sev = p3.add_run(f"Kemungkinan {likelihood} × Dampak {impact} = {score_text} ({severity})")
    r_sev.font.bold = True
    r_sev.font.color.rgb = RGBColor(197, 48, 48) if severity == "Tinggi (High)" else RGBColor(221, 107, 32)
    p3.runs[0].font.size = Pt(9.5)
    r_sev.font.size = Pt(9.5)

    p4 = r2_c2.paragraphs[0]
    p4.add_run("Status Penanganan: ").font.bold = True
    p4.add_run("Open (Rekomendasi Kontrol)")
    p4.runs[0].font.size = Pt(9.5)
    p4.runs[1].font.size = Pt(9.5)

    # Row 3: Skenario
    r3 = table.cell(3, 0)
    r3.merge(table.cell(3, 1))
    set_cell_margins(r3, 80, 80, 120, 120)
    p_sc = r3.paragraphs[0]
    p_sc.add_run("Skenario Ancaman: ").font.bold = True
    p_sc.add_run(scenario)
    p_sc.runs[0].font.size = Pt(9.5)
    p_sc.runs[1].font.size = Pt(9.5)

    # Row 4: Dampak & Asumsi
    r4 = table.cell(4, 0)
    r4.merge(table.cell(4, 1))
    set_cell_margins(r4, 80, 80, 120, 120)
    p_da = r4.paragraphs[0]
    p_da.add_run("Dampak: ").font.bold = True
    p_da.add_run(f"{impact_desc}\n")
    p_da.add_run("Asumsi Terkait: ").font.bold = True
    p_da.add_run(assumption)
    for r in p_da.runs:
        r.font.size = Pt(9.5)

    # Row 5: Mitigasi
    r5 = table.cell(5, 0)
    r5.merge(table.cell(5, 1))
    set_cell_background(r5, "EDF2F7")
    set_cell_margins(r5, 90, 90, 120, 120)
    p_mit = r5.paragraphs[0]
    p_mit.add_run("Rekomendasi Mitigasi Teknis: ").font.bold = True
    p_mit.add_run(mitigation)
    p_mit.runs[0].font.size = Pt(9.5)
    p_mit.runs[0].font.color.rgb = RGBColor(26, 54, 93)
    p_mit.runs[1].font.size = Pt(9.5)

    # Set column widths
    for row in table.rows:
        row.cells[0].width = Inches(3.2)
        row.cells[1].width = Inches(3.3)

    # Add spacing after table
    sp_p = doc.add_paragraph()
    sp_p.paragraph_format.space_before = Pt(0)
    sp_p.paragraph_format.space_after = Pt(8)

def build_laporan():
    doc = docx.Document()

    # Set page margins to 1 inch
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # ==========================
    # COVER / HEADER
    # ==========================
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(10)
    title_p.paragraph_format.space_after = Pt(4)
    run_t = title_p.add_run("LAPORAN THREAT MODELING BERBASIS STRIDE")
    run_t.font.name = "Calibri"
    run_t.font.size = Pt(18)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor(26, 54, 93)

    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_after = Pt(14)
    run_s = sub_p.add_run("Studi Kasus A: Sistem Presensi Perkuliahan Berbasis QR Code\nMata Kuliah Keamanan Siber (CAK3CAB3) — Semester Ganjil 2026/2027")
    run_s.font.name = "Calibri"
    run_s.font.size = Pt(12)
    run_s.font.color.rgb = RGBColor(74, 85, 104)

    # Meta Table
    meta_table = doc.add_table(rows=6, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(meta_table, "CBD5E0")
    meta_headers = ["Atribut Dokumen", "Keterangan Identitas Proyek"]
    meta_data = [
        ["Identitas Tim / Kasus", "Studi Kasus A: Sistem Presensi Perkuliahan Berbasis QR Code"],
        ["Anggota Tim Penyusun", "1. Fatir Gibran (NIM: 103112430153) — Ketua Tim & Pemodel DFD\n2. Tio Armani (NIM: 103112430225) — Analis STRIDE: Spoofing & Tampering\n3. Chadafya Putra Zulfikar (NIM: 103112430173) — Analis STRIDE: R, I, D, E\n4. Tedy Permana Putra (NIM: 103112430157) — Penilai Risiko & QA Editor"],
        ["Reviewer / Dosen", "Dosen Praktikum Keamanan Siber (CAK3CAB3)"],
        ["Perangkat Pemodelan", "OWASP Threat Dragon v2.x (Desktop / Web Engine)"],
        ["Status Laporan", "Laporan Resmi Komprehensif — Terverifikasi Siap Kumpul"]
    ]
    format_table(meta_table, [2.2, 4.3], meta_headers, meta_data)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ==========================
    # 1. PENDAHULUAN
    # ==========================
    add_heading_1(doc, "1. Pendahuluan")
    
    add_heading_2(doc, "1.1 Tujuan dan Latar Belakang")
    add_body_paragraph(doc, 
        "Penerapan teknologi informasi pada sektor pendidikan tinggi menuntut kepastian integritas data dan ketersediaan layanan yang tinggi. Salah satu sistem operasional vital adalah sistem presensi perkuliahan. Institusi perguruan tinggi mengimplementasikan layanan presensi berbasis web guna mencatat kehadiran mahasiswa di ruang kelas secara efisien. Dalam alur operasionalnya, dosen pengampu membuka sesi perkuliahan, sistem membangkitkan QR code yang diproyeksikan pada layar kelas, mahasiswa melakukan autentikasi dengan akun kampus dan memindai QR code tersebut, selanjutnya Backend Presensi memvalidasi dan membukukan kehadiran ke Database Akademik.")
    add_body_paragraph(doc,
        "Aktivitas threat modeling ini bertujuan untuk mengidentifikasi potensi kerentanan, vektor serangan, dan kegagalan keamanan sejak tahap perancangan arsitektur sistem tanpa memerlukan akses terhadap kode sumber maupun aplikasi yang berjalan langsung (design-level security analysis). Mengacu pada metodologi STRIDE (Spoofing, Tampering, Repudiation, Information Disclosure, Denial of Service, dan Elevation of Privilege) serta berbantukan perangkat OWASP Threat Dragon, analisis ini menghasilkan inventarisasi ancaman komprehensif, evaluasi tingkat keparahan risiko secara objektif, serta rekomendasi kontrol mitigasi yang tepat guna.")

    add_heading_2(doc, "1.2 Ruang Lingkup Analisis")
    add_body_paragraph(doc, "Ruang lingkup analisis dibatasi secara tegas agar threat modeling terfokus pada arsitektur sistem presensi:")
    add_bullet_point(doc, "Mekanisme pembukaan dan penutupan sesi perkuliahan oleh dosen pengampu, pembangkitan dan penayangan token/kode QR dinamis, proses pemindaian QR oleh mahasiswa, validasi parameter sesi dan status keanggotaan mahasiswa di kelas, pencatatan transaksi kehadiran ke Database Akademik, mekanisme pemuatan rekap kehadiran (dosen dan admin), serta interaksi otorisasi dan pertukaran token dengan Layanan Autentikasi Kampus (SSO).", "Komponen dalam Lingkup (In-Scope): ")
    add_bullet_point(doc, "Proses pendaftaran mata kuliah/rencana studi semesteran, perubahan nilai akademik mahasiswa, keamanan infrastruktur fisik ruang kelas, penyedia cloud hosting/jaringan fisik kampus, pengujian penetrasi aktif (penetration testing), eksploitasi kode nyata, serta audit baris kode sumber aplikasi.", "Komponen di Luar Lingkup (Out-of-Scope): ")

    add_heading_2(doc, "1.3 Alur Bisnis Utama Sistem")
    add_body_paragraph(doc,
        "Alur bisnis inti sistem presensi perkuliahan berlangsung dalam tahapan sekuensial berikut:")
    add_bullet_point(doc, "Dosen pengampu melakukan autentikasi melalui Layanan Autentikasi Kampus (SSO) kemudian membuka sesi perkuliahan untuk kelas yang diampunya melalui antarmuka web.", "1. Pembukaan Sesi: ")
    add_bullet_point(doc, "Backend Presensi memverifikasi identitas dosen, memastikan hak pengampu kelas pada jadwal yang bersangkutan, membuat sesi presensi baru berdurasi aktif 5 menit, dan menerbitkan token QR yang diproyeksikan di layar kelas.", "2. Pembangkitan QR: ")
    add_bullet_point(doc, "Mahasiswa yang hadir secara fisik membuka browser, masuk menggunakan akun SSO kampus, dan memindai QR code. Browser mahasiswa mengirimkan identitas sesi, token QR, serta token autentikasi mahasiswa ke Backend Presensi.", "3. Pemindaian Presensi: ")
    add_bullet_point(doc, "Backend Presensi memvalidasi keabsahan token SSO mahasiswa, memastikan sesi presensi masih aktif (dalam jendela 5 menit), memverifikasi bahwa mahasiswa terdaftar secara sah pada kelas tersebut, serta memastikan mahasiswa belum tercatat pada sesi berjalan.", "4. Validasi & Pembukuan: ")
    add_bullet_point(doc, "Jika seluruh kriteria valid, kehadiran dicatat ke Database Akademik dan konfirmasi keberhasilan dikirim ke mahasiswa. Dosen dapat memantau daftar hadir secara langsung atau menutup sesi lebih awal. Admin akademik memiliki wewenang untuk meninjau rekapitulasi seluruh kelas.", "5. Pemantauan & Rekap: ")

    add_heading_2(doc, "1.4 Aturan Bisnis (Business Rules)")
    add_body_paragraph(doc, "Terdapat tiga aturan bisnis fundamental yang menjadi batas integritas operasional sistem:")
    add_bullet_point(doc, "Hanya dosen yang terdaftar resmi sebagai pengampu mata kuliah dan kelas terkait yang berhak membuka, memantau, dan menutup sesi presensi.", "Aturan 1 (Otorisasi Dosen): ")
    add_bullet_point(doc, "Satu akun mahasiswa hanya diperbolehkan memiliki maksimal satu catatan kehadiran yang sah untuk setiap sesi perkuliahan yang berlangsung.", "Aturan 2 (Keunikan Presensi): ")
    add_bullet_point(doc, "Sesi presensi memiliki batas masa berlaku ketat selama 5 menit sejak dibuka, namun dosen pengampu memiliki hak prerogatif untuk menutup sesi lebih cepat sewaktu-waktu.", "Aturan 3 (Batas Waktu Sesi): ")

    add_heading_2(doc, "1.5 Identifikasi Aset dan Pemangku Kepentingan")
    add_body_paragraph(doc, "Identifikasi aset sistem mencakup aspek kerahasiaan (Confidentiality), integritas (Integrity), ketersediaan (Availability), serta keaslian (Authenticity):")

    asset_table = doc.add_table(rows=8, cols=3)
    asset_headers = ["Nama Aset Sistem", "Pihak Berkepentingan", "Sifat Keamanan yang Dijaga"]
    asset_data = [
        ["Catatan Kehadiran Mahasiswa", "Mahasiswa, Dosen Pengampu, Admin Akademik, Program Studi", "Integritas (keaslian data), Non-repudiation (bukti presensi tidak dapat disangkal)"],
        ["Token / Kode QR Sesi", "Dosen Pengampu, Institusi", "Keaslian, Entropi tinggi, Masa berlaku terbatas (TTL 5 menit)"],
        ["Identitas & Token Autentikasi (JWT/SSO)", "Seluruh Pengguna (Mahasiswa, Dosen, Admin)", "Kerahasiaan, Integritas tanda tangan kriptografis, Anti-tampering"],
        ["Keanggotaan Kelas & Data Jadwal Kuliah", "Dosen Pengampu, Admin Akademik", "Integritas data relasi mahasiswa-kelas"],
        ["Rekapitulasi Kehadiran & Nilai Keaktifan", "Dosen Pengampu, Admin Akademik, Mahasiswa", "Kerahasiaan data pribadi mahasiswa, Integritas laporan"],
        ["Ketersediaan Layanan Presensi", "Mahasiswa, Dosen Pengampu", "Ketersediaan (Availability) prima selama jendela 5 menit"],
        ["Log Audit Transaksi Keamanan", "Admin Sistem, Tim Keamanan Institusi", "Integritas, Sifat append-only (anti-tampering), Non-repudiation"]
    ]
    format_table(asset_table, [2.0, 2.2, 2.3], asset_headers, asset_data)

    add_heading_2(doc, "1.6 Asumsi Teknis Perancangan")
    add_body_paragraph(doc, "Analisis threat modeling bertumpu pada lima asumsi teknis eksplisit berikut:")
    
    asump_table = doc.add_table(rows=6, cols=3)
    asump_headers = ["ID Asumsi", "Pernyataan Asumsi Teknis", "Rasional Keamanan"]
    asump_data = [
        ["A1 (asumsi)", "Seluruh komunikasi data klien-ke-server (browser ke Backend) dan server-ke-server (Backend ke SSO) diamankan dengan protokol HTTPS/TLS.", "Mereduksi risiko network sniffing pada saluran publik, sehingga analisis fokus pada titik akhir (endpoint) dan aplikasi."],
        ["A2 (asumsi)", "Layanan Autentikasi Kampus mengimplementasikan mekanisme Single Sign-On (SSO) standar yang menerbitkan token bertanda tangan dengan masa berlaku terbatas, dan backend memverifikasi tanda tangan tersebut secara konsisten.", "Identitas pengguna didasarkan pada verifikasi kriptografis token SSO terpercaya."],
        ["A3 (asumsi)", "Database Akademik berada dalam segmen jaringan privat internal yang terisolasi dan hanya dapat diakses langsung oleh Backend Presensi dan administrator basis data resmi.", "Mencegah akses jaringan langsung dari luar perimeter kampus tanpa otorisasi."],
        ["A4 (asumsi)", "Token QR memuat identitas sesi dan string acak, bersifat statis selama sesi 5 menit aktif karena sistem awal tidak merancang mekanisme dynamic refresh.", "Membuka potensi pembagian citra QR (remote sharing) ke pihak yang tidak hadir secara fisik."],
        ["A5 (asumsi)", "Mekanisme pencatatan log audit terperinci dan pembatasan laju permintaan (rate limiting) belum dijelaskan dalam spesifikasi awal sistem.", "Prinsip evaluasi: kontrol yang belum tercantum tidak serta-merta dianggap mustahil, namun belum dapat dianggap teruji dan terverifikasi."]
    ]
    format_table(asump_table, [1.1, 3.2, 2.2], asump_headers, asump_data)

    # ==========================
    # 2. MODEL DFD
    # ==========================
    add_heading_1(doc, "2. Model Data Flow Diagram (DFD)")
    add_heading_2(doc, "2.1 Diagram Aliran Data Level-1")
    add_body_paragraph(doc,
        "Pemodelan arsitektur sistem presensi digambarkan melalui Data Flow Diagram (DFD) Level-1 yang merepresentasikan aktor luar (external entities), proses pemrosesan backend, penyimpan data permanen (data store), aliran informasi, serta batas kepercayaan (trust boundaries).")

    if os.path.exists("dfd_kasus_a.png"):
        doc.add_picture("dfd_kasus_a.png", width=Inches(6.4))
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(4)
        p_cap.paragraph_format.space_after = Pt(10)
        rcap = p_cap.add_run("Gambar 1. Data Flow Diagram (DFD) Sistem Presensi Perkuliahan Berbasis QR Code")
        rcap.font.italic = True
        rcap.font.size = Pt(9)
        rcap.font.color.rgb = RGBColor(100, 116, 139)

    add_heading_2(doc, "2.2 Inventarisasi Elemen DFD")
    add_body_paragraph(doc, "Sistem tersusun atas 8 elemen fungsional utama:")
    
    elem_table = doc.add_table(rows=9, cols=3)
    elem_headers = ["ID Elemen", "Tipe Elemen", "Deskripsi Tanggung Jawab Fungsional"]
    elem_data = [
        ["E1", "Actor (Browser Klien)", "Mahasiswa: Melakukan login SSO, memindai QR code di ruang kelas, dan mengirimkan permintaan pencatatan presensi."],
        ["E2", "Actor (Browser Klien)", "Dosen Pengampu: Melakukan login SSO, membuka/menutup sesi perkuliahan, menayangkan kode QR di proyektor, dan memantau daftar hadir."],
        ["E3", "Actor (Browser Klien)", "Admin Akademik: Mengakses rekapitulasi data kehadiran seluruh kelas untuk kebutuhan monitoring dan evaluasi institusi."],
        ["E4", "Actor (External Service)", "Layanan Autentikasi Kampus (SSO/Identity Provider): Memverifikasi kredensial akun kampus, menerbitkan token autentikasi, dan memvalidasi klaim pengguna."],
        ["P1", "Process (Backend Presensi)", "Kelola Sesi & QR: Menangani pembukaan dan penutupan sesi presensi, membangkitkan string/token QR sesi berdurasi 5 menit, dan memvalidasi wewenang dosen pengampu."],
        ["P2", "Process (Backend Presensi)", "Validasi & Catat Presensi: Memverifikasi masa aktif sesi presensi, mencocokkan token QR, memvalidasi keanggotaan mahasiswa pada kelas, dan menjamin kepatuhan aturan satu catatan per mahasiswa."],
        ["P3", "Process (Backend Presensi)", "Muat Rekap Kehadiran: Mengambil data riwayat kehadiran dari basis data dan menyajikannya sesuai tingkat otorisasi peran pengguna (dosen vs admin)."],
        ["DS1", "Data Store", "Database Akademik: Penyimpanan persisten untuk data akun, jadwal kuliah, relasi mahasiswa-kelas, status sesi aktif, token QR, dan catatan log kehadiran."]
    ]
    format_table(elem_table, [0.9, 1.8, 3.8], elem_headers, elem_data)
    add_body_paragraph(doc, "Catatan: P1, P2, dan P3 merupakan pemisahan logis dari layanan Backend Presensi kampus.")

    add_heading_2(doc, "2.3 Inventarisasi Aliran Data (Data Flow)")
    add_body_paragraph(doc, "Terdapat 16 aliran data (F1 hingga F16) yang memindahkan informasi antar-elemen:")

    flow_table = doc.add_table(rows=17, cols=4)
    flow_headers = ["ID Aliran", "Arah Komunikasi", "Muatan Data / Parameter", "Lintasan Batas Kepercayaan"]
    flow_data = [
        ["F1", "E2 → P1", "Permintaan buka/tutup sesi (ID kelas, jadwal, token SSO dosen)", "TB1 → TB2 (Perangkat Dosen ke Backend)"],
        ["F2", "P1 → E2", "Data token QR sesi, string acak, dan status sesi perkuliahan", "TB2 → TB1 (Backend ke Perangkat Dosen)"],
        ["F3", "E1 → P2", "Payload presensi (ID sesi, token QR terdeteksi, token SSO mahasiswa)", "TB1 → TB2 (Perangkat Mahasiswa ke Backend)"],
        ["F4", "P2 → E1", "Status respons presensi (berhasil dicatat / ditolak beserta kode pesan)", "TB2 → TB1 (Backend ke Perangkat Mahasiswa)"],
        ["F5", "E1/E2/E3 → E4", "Kredensial login akun kampus (NIM/NIP, kata sandi, kode MFA)", "TB1 → TB-Auth (Klien ke Layanan SSO)"],
        ["F6", "E4 → E1/E2/E3", "Token autentikasi bertanda tangan (JWT/token sesi)", "TB-Auth → TB1 (Layanan SSO ke Klien)"],
        ["F7", "P1/P2/P3 → E4", "Permintaan introspeksi/verifikasi tanda tangan dan klaim token", "TB2 → TB-Auth (Backend ke Layanan SSO)"],
        ["F8", "E4 → P1/P2/P3", "Hasil validasi kriptografis: identitas terverifikasi dan peran pengguna", "TB-Auth → TB2 (Layanan SSO ke Backend)"],
        ["F9", "P1 → DS1", "Operasi insert/update status sesi baru, masa berlaku, dan token QR", "TB2 → TB3 (Backend ke Database Akademik)"],
        ["F10", "DS1 → P1", "Data verifikasi jadwal perkuliahan dan relasi dosen pengampu sah", "TB3 → TB2 (Database Akademik ke Backend)"],
        ["F11", "P2 → DS1", "Query pengecekan keanggotaan kelas dan transaksi pencatatan hadir", "TB2 → TB3 (Backend ke Database Akademik)"],
        ["F12", "DS1 → P2", "Konfirmasi status sesi aktif, status keanggotaan, dan hasil eksekusi", "TB3 → TB2 (Database Akademik ke Backend)"],
        ["F13", "E2/E3 → P3", "Permintaan penayangan rekap kehadiran (ID mata kuliah, token peran)", "TB1 → TB2 (Klien Dosen/Admin ke Backend)"],
        ["F14", "P3 → DS1", "Eksekusi query agregasi data rekapitulasi kehadiran mahasiswa", "TB2 → TB3 (Backend ke Database Akademik)"],
        ["F15", "DS1 → P3", "Hasil himpunan data kehadiran historis dan status perkuliahan", "TB3 → TB2 (Database Akademik ke Backend)"],
        ["F16", "P3 → E2/E3", "Paket data respons rekapitulasi kehadiran (daftar hadir & analitik)", "TB2 → TB1 (Backend ke Klien Dosen/Admin)"]
    ]
    format_table(flow_table, [0.8, 1.3, 2.7, 1.7], flow_headers, flow_data)

    add_heading_2(doc, "2.4 Batas Kepercayaan (Trust Boundaries)")
    add_body_paragraph(doc, "Empat batas kepercayaan didefinisikan guna memisahkan zona-zona dengan tingkat kepercayaan dan kontrol berbeda:")

    tb_table = doc.add_table(rows=5, cols=3)
    tb_headers = ["ID Batas", "Nama Batas Kepercayaan & Elemen", "Rasional Keamanan & Karakteristik"]
    tb_data = [
        ["TB1", "Perangkat Pengguna / Internet\n(Melingkupi E1, E2, E3)", "Zona tidak terpercaya (untrusted). Perangkat klien berada di luar kontrol institusi, rentan disusupi malware atau manipulasi parameter input oleh pengguna. Semua payload data yang masuk dari TB1 wajib divalidasi ketat di sisi server."],
        ["TB-Auth", "Layanan Autentikasi Kampus\n(Melingkupi E4)", "Zona identitas eksternal yang dipercaya untuk memverifikasi akun pengguna. Walaupun sangat terpercaya, dependensi ini menimbulkan titik kegagalan tunggal (single point of failure) terhadap ketersediaan sistem."],
        ["TB2", "Backend Presensi\n(Melingkupi P1, P2, P3)", "Zona pemrosesan aplikasi terpercaya (trusted application tier). Menjadi benteng utama tempat dilakukannya validasi otorisasi, pemeriksaan logika bisnis, dan evaluasi masa aktif sesi presensi."],
        ["TB3", "Zona Data Internal\n(Melingkupi DS1)", "Zona penyimpanan paling sensitif dan terisolasi ketat (high trust data tier). Hanya dapat diakses melalui koneksi terotentikasi dari Backend Presensi dan DBA resmi sesuai Asumsi A3."]
    ]
    format_table(tb_table, [1.0, 2.3, 3.2], tb_headers, tb_data)

    add_heading_2(doc, "2.5 Penelaahan Empat Pertanyaan Inti Threat Modeling")
    add_bullet_point(doc, "Sistem web presensi perkuliahan berbasis pemindaian QR code yang menghubungkan perangkat dosen, mahasiswa, layanan SSO, backend aplikasi, dan basis data akademik dengan aturan bisnis ketat (5 menit dan otorisasi pengampu).", "1. Apa yang sedang dibangun? ")
    add_bullet_point(doc, "Pengguna dapat memanipulasi parameter, membagikan QR code ke luar kelas secara jarak jauh, menitipkan kredensial akun, melakukan injeksi basis data, mengalami lonjakan permintaan (Denial of Service), atau mengakses data rekapitulasi tanpa hak otorisasi yang sah.", "2. Apa yang bisa salah? ")
    add_bullet_point(doc, "Menerapkan kontrol keamanan berbasis pertahanan berlapis (defense-in-depth): tokenisasi dinamis, validasi otorisasi berbasis peran dan kepemilikan objek (RBAC/ABAC), penggunaan parameter binding SQL, pembatasan laju (rate limiting), serta audit logging terenkripsi.", "3. Apa yang akan dilakukan? ")
    add_bullet_point(doc, "Evaluasi asumsi teknis menunjukkan bahwa seluruh 16 aliran data dan elemen DFD telah ditinjau secara metodis menggunakan matriks STRIDE, dengan 14 ancaman yang terdokumentasi secara spesifik dan memiliki mitigasi terukur.", "4. Apakah analisis telah memadai? ")

    # ==========================
    # 3. METODOLOGI ANALISIS
    # ==========================
    add_heading_1(doc, "3. Metodologi Analisis Ancaman")
    add_heading_2(doc, "3.1 Kerangka Kerja STRIDE dan Rubrik Penilaian Risiko")
    add_body_paragraph(doc,
        "Analisis ancaman mengadopsi taksonomi STRIDE yang dikembangkan oleh Microsoft. Kategori STRIDE digunakan untuk memicu identifikasi kegagalan desain:")

    stride_table = doc.add_table(rows=7, cols=3)
    stride_headers = ["Kategori STRIDE", "Definisi Konseptual", "Pertanyaan Pemicu Analisis"]
    stride_data = [
        ["Spoofing (S)", "Menyamarkan identitas sebagai pengguna atau entitas lain secara tidak sah.", "Dapatkah pihak luar atau mahasiswa menyamar sebagai mahasiswa lain atau dosen pengampu?"],
        ["Tampering (T)", "Mengubah data, transaksi, atau parameter tanpa otorisasi yang sah.", "Dapatkah muatan presensi atau token QR dimanipulasi di tengah jalan?"],
        ["Repudiation (R)", "Menyangkal tindakan yang telah dilakukan karena bukti audit tidak memadai.", "Dapatkah mahasiswa atau staf menyangkal kehadiran atau perubahan data presensi?"],
        ["Information Disclosure (I)", "Membocorkan informasi rahasia kepada pihak yang tidak berwenang.", "Dapatkah token QR atau data rekapitulasi kelas diakses oleh entitas yang tidak berhak?"],
        ["Denial of Service (D)", "Mengganggu ketersediaan layanan sehingga pengguna sah tidak dapat mengaksesnya.", "Dapatkah sistem presensi dilumpuhkan selama jendela kritis 5 menit?"],
        ["Elevation of Privilege (E)", "Memperoleh hak akses atau wewenang melebihi batas yang telah ditentukan.", "Dapatkah mahasiswa membuka sesi kuliah atau dosen melihat rekapitulasi kelas lain?"]
    ]
    format_table(stride_table, [1.5, 2.5, 2.5], stride_headers, stride_data)

    add_body_paragraph(doc, 
        "Penilaian risiko menerapkan rubrik ordinal sederhana 1–3 untuk parameter Kemungkinan (Likelihood - L) dan Dampak (Impact - D). Skor Risiko dihitung melalui perkalian: Skor = L × D.")
    
    rubrik_table = doc.add_table(rows=4, cols=3)
    rubrik_headers = ["Nilai", "Tingkat Kemungkinan (Likelihood - L)", "Tingkat Dampak (Impact - D)"]
    rubrik_data = [
        ["1", "Rendah: Kondisi serangan sulit diwujudkan atau membutuhkan akses khusus.", "Rendah: Dampak terbatas, data tidak bocor luas, mudah dipulihkan."],
        ["2", "Sedang: Masuk akal terjadi dalam skenario penggunaan perkuliahan reguler.", "Sedang: Mengganggu sebagian pengguna atau merusak integritas parsial."],
        ["3", "Tinggi: Relatif mudah dieksekusi mahasiswa atau sering terjadi tanpa hambatan teknis.", "Tinggi: Akses tidak sah skala besar, kebocoran data sensitif, integritas kehadiran rusak total, atau layanan lumpuh."]
    ]
    format_table(rubrik_table, [0.6, 2.9, 3.0], rubrik_headers, rubrik_data)
    add_body_paragraph(doc, "Interpretasi Skor Risiko: Skor 1–2 = Rendah (Low); Skor 3–4 = Sedang (Medium); Skor 6–9 = Tinggi (High).")

    add_heading_2(doc, "3.2 Tinjauan Matriks STRIDE per Elemen DFD")
    add_body_paragraph(doc, "Tabel berikut merangkum penelaahan STRIDE pada seluruh elemen DFD:")

    matrix_table = doc.add_table(rows=11, cols=8)
    matrix_headers = ["Elemen DFD", "S", "T", "R", "I", "D", "E", "ID Ancaman Terpetakan"]
    align_c = {1: WD_ALIGN_PARAGRAPH.CENTER, 2: WD_ALIGN_PARAGRAPH.CENTER, 3: WD_ALIGN_PARAGRAPH.CENTER,
               4: WD_ALIGN_PARAGRAPH.CENTER, 5: WD_ALIGN_PARAGRAPH.CENTER, 6: WD_ALIGN_PARAGRAPH.CENTER}
    matrix_data = [
        ["E1. Mahasiswa (Browser)", "✓", "-", "-", "-", "-", "-", "T01 (Titip absen akun)"],
        ["E2/E3. Dosen/Admin", "-", "-", "-", "-", "-", "-", "Ditinjau; risiko tercakup di P1 & P3"],
        ["E4. Layanan Autentikasi", "-", "-", "-", "-", "✓", "-", "T12 (Dependensi ketersediaan)"],
        ["P1. Kelola Sesi & QR", "✓", "-", "-", "-", "-", "✓", "T04 (Prediksi QR), T13 (Otorisasi sesi)"],
        ["P2. Validasi & Catat", "✓", "✓", "✓", "-", "✓", "-", "T03 (QR dibagikan), T06 (Race cond), T08 (Audit log), T11 (DoS lonjakan)"],
        ["P3. Muat Rekap", "-", "-", "-", "-", "-", "✓", "T14 (Akses rekap lintas kelas)"],
        ["DS1. Database Akademik", "-", "-", "-", "✓", "-", "-", "T10 (Pencurian cadangan/data store)"],
        ["F3. Mahasiswa ke P2", "✓", "✓", "-", "-", "-", "-", "T02 (Pencurian token), T05 (Manipulasi ID)"],
        ["F11. P2 ke Database", "-", "✓", "-", "-", "-", "-", "T07 (SQL Injection transaksi)"],
        ["F16. P3 ke Klien", "-", "-", "-", "✓", "-", "-", "T09 (Eksposur data pribadi berlebih)"]
    ]
    format_table(matrix_table, [1.8, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 2.3], matrix_headers, matrix_data, align_c)

    # ==========================
    # 4. REGISTER ANCAMAN
    # ==========================
    add_heading_1(doc, "4. Register Ancaman Keamanan (Threat Register)")
    add_heading_2(doc, "4.1 Ringkasan Register Ancaman")
    add_body_paragraph(doc, "Berikut rekapitulasi 14 ancaman keamanan yang diidentifikasi secara komprehensif:")

    reg_table = doc.add_table(rows=15, cols=8)
    reg_headers = ["ID", "Judul Ancaman Spesifik", "STRIDE", "Elemen", "L", "D", "Skor", "Tingkat Risiko"]
    reg_data = [
        ["T01", "Titip absen dengan akun mahasiswa lain", "Spoofing", "E1", "3", "2", "6", "Tinggi (High)"],
        ["T02", "Token autentikasi mahasiswa dicuri dan dipakai ulang", "Spoofing", "F3", "2", "3", "6", "Tinggi (High)"],
        ["T03", "QR dibagikan jarak jauh ke mahasiswa tidak hadir", "Spoofing", "P2", "3", "2", "6", "Tinggi (High)"],
        ["T04", "Token QR dapat ditebak atau diprediksi", "Spoofing", "P1", "2", "2", "4", "Sedang (Medium)"],
        ["T05", "Manipulasi ID sesi atau ID mahasiswa pada presensi", "Tampering", "F3", "2", "3", "6", "Tinggi (High)"],
        ["T06", "Presensi ganda atau manipulasi waktu klien (race condition)", "Tampering", "P2", "2", "2", "4", "Sedang (Medium)"],
        ["T07", "SQL injection melalui parameter presensi ke database", "Tampering", "F11", "2", "3", "6", "Tinggi (High)"],
        ["T08", "Pengubahan atau penyangkalan catatan hadir tanpa audit log", "Repudiation", "P2", "2", "2", "4", "Sedang (Medium)"],
        ["T09", "Rekap memuat data pribadi berlebih ke pihak tidak berhak", "Info Disclosure", "F16", "2", "2", "4", "Sedang (Medium)"],
        ["T10", "Akses langsung atau kebocoran cadangan basis data", "Info Disclosure", "DS1", "1", "3", "3", "Sedang (Medium)"],
        ["T11", "Lonjakan permintaan presensi serentak melumpuhkan backend", "Denial of Service", "P2", "3", "2", "6", "Tinggi (High)"],
        ["T12", "Layanan Autentikasi Kampus lumpuh menghentikan presensi", "Denial of Service", "E4", "2", "3", "6", "Tinggi (High)"],
        ["T13", "Dosen non-pengampu atau pihak lain membuka sesi kelas", "Elevation of Priv", "P1", "2", "3", "6", "Tinggi (High)"],
        ["T14", "Akses data rekapitulasi lintas kelas oleh peran tidak berhak", "Elevation of Priv", "P3", "2", "3", "6", "Tinggi (High)"]
    ]
    align_reg = {0: WD_ALIGN_PARAGRAPH.CENTER, 4: WD_ALIGN_PARAGRAPH.CENTER, 5: WD_ALIGN_PARAGRAPH.CENTER, 6: WD_ALIGN_PARAGRAPH.CENTER}
    format_table(reg_table, [0.5, 2.1, 1.0, 0.6, 0.3, 0.3, 0.4, 1.3], reg_headers, reg_data, align_reg)

    add_heading_2(doc, "4.2 Detail Analisis Setiap Entri Ancaman")
    add_body_paragraph(doc, "Berikut analisis mendalam untuk setiap entri ancaman (T01 hingga T14) sesuai format standar evaluasi keamanan:")

    # Cards for T01 - T14
    add_threat_card(doc, "T01", "Titip absen dengan akun mahasiswa lain", "Spoofing", "E1. Mahasiswa (Browser)",
        3, 2, "6", "Tinggi (High)",
        "Jika mahasiswa yang tidak hadir di kelas memberikan kredensial akun SSO miliknya secara sukarela kepada rekannya yang berada di ruang kelas, dan rekannya tersebut melakukan login serta memindai kode QR dari layar proyektor, maka sistem akan mencatat mahasiswa tersebut hadir secara sah.",
        "Integritas data kehadiran perkuliahan menjadi rusak, merugikan tata tertib akademik institusi, serta menciptakan ketidakadilan evaluasi antarmahasiswa.",
        "Asumsi A2: Layanan SSO menggunakan kredensial standar tanpa pengikatan identitas perangkat keras (device fingerprinting).",
        "Mewajibkan autentikasi multifaktor (MFA) berbasis aplikasi pada akun kampus; menerapkan mekanisme pengikatan sesi ke fingerprint perangkat browser; mendeteksi dan memblokir anomali ketika satu perangkat fisik mengirimkan presensi untuk beberapa akun mahasiswa berbeda dalam sesi yang sama; menyediakan notifikasi konfirmasi presensi real-time ke perangkat utama mahasiswa.")

    add_threat_card(doc, "T02", "Token autentikasi mahasiswa dicuri dan dipakai ulang", "Spoofing", "F3. Mahasiswa ke P2",
        2, 3, "6", "Tinggi (High)",
        "Jika penyerang berhasil mencuri token autentikasi (JWT/Bearer token) milik mahasiswa sah melalui session hijacking pada perangkat bersama (laboratorium/warnet) atau celah XSS, kemudian mengirimkan token tersebut pada payload F3, maka Backend Presensi akan mempercayai transaksi tersebut berasal dari korban.",
        "Terjadi pencatatan kehadiran palsu atas nama korban tanpa sepengetahuannya, serta potensi pembajakan profil akun mahasiswa.",
        "Asumsi A1: Jaringan menggunakan TLS, sehingga risiko intersepsi sniffing minim dan vektor pencurian terkonsentrasi pada manipulasi sisi klien (browser storage).",
        "Menetapkan masa aktif token (TTL) sangat singkat dengan pembatasan audience khusus presensi perkuliahan; menyimpan token dalam cookie beratribut HttpOnly, Secure, dan SameSite=Strict guna mencegah pembacaan melalui skrip JavaScript; mengimplementasikan mekanisme token revocation saat pengguna logout; mewajibkan kebijakan HSTS.")

    add_threat_card(doc, "T03", "QR dibagikan jarak jauh ke mahasiswa yang tidak hadir", "Spoofing", "P2. Validasi & Catat Presensi",
        3, 2, "6", "Tinggi (High)",
        "Jika mahasiswa yang berada di ruang kelas memotret kode QR yang ditampilkan di layar lalu menyebarkannya melalui aplikasi pesan instan kepada mahasiswa lain di luar kampus, maka mahasiswa di luar kelas dapat memindai foto tersebut selama batas waktu 5 menit masih aktif dan berhasil tercatat hadir.",
        "Keabsahan presensi fisik perkuliahan gagal total; tujuan verifikasi keberadaan fisik mahasiswa di ruang kelas tidak tercapai.",
        "Asumsi A4: Kode QR bersifat statis sepanjang durasi 5 menit sesi perkuliahan tanpa validasi geolokasi maupun verifikasi jaringan.",
        "Mengubah mekanisme penayangan QR statis menjadi QR dinamis (rolling QR code) yang diperbarui otomatis setiap 15–30 detik berbasis algoritma TOTP; membatasi subnet IP jaringan pemindai hanya pada Wi-Fi internal ruang kelas perkuliahan; mengintegrasikan deteksi anomali keterlambatan pemindaian.")

    add_threat_card(doc, "T04", "Token QR dapat ditebak atau diprediksi", "Spoofing", "P1. Kelola Sesi & QR",
        2, 2, "4", "Sedang (Medium)",
        "Jika algoritma pembangkit token QR pada proses P1 menggunakan generator bilangan pseudo-acak yang lemah (low entropy) atau berpola inkremental sekuensial, maka penyerang dapat menebak nilai token yang valid dan mengirimkan presensi tanpa perlu melihat atau memindai kode QR fisik.",
        "Potensi kecurangan presensi otomatis berskala luas menggunakan skrip injeksi token hasil tebakan.",
        "Asumsi teknis: Spesifikasi algoritma pembangkitan string acak pada rancangan awal sistem belum didefinisikan secara eksplisit.",
        "Wajib menggunakan Cryptographically Secure Pseudo-Random Number Generator (CSPRNG) dengan entropi minimal 128-bit untuk membangkitkan token; mengaitkan token secara kriptografis dengan ID sesi perkuliahan; membatasi percobaan pengiriman token salah (rate limiting dan auto-lock) per akun mahasiswa.")

    add_threat_card(doc, "T05", "Manipulasi ID sesi atau ID mahasiswa pada permintaan presensi", "Tampering", "F3. Mahasiswa ke P2",
        2, 3, "6", "Tinggi (High)",
        "Jika Backend Presensi (P2) menerima dan mempercayai parameter ID mahasiswa atau ID sesi yang dikirimkan pada body/query parameter permintaan tanpa memverifikasinya terhadap klaim identitas pada token autentikasi, maka pengguna dapat merekayasa parameter untuk mencatatkan kehadiran bagi mahasiswa lain atau kelas lain.",
        "Manipulasi catatan akademik secara ilegal; pelanggaran integritas batasan keanggotaan kelas perkuliahan.",
        "Asumsi teknis: Backend belum tentu mengekstrak identitas secara eksklusif dari klaim token yang telah divalidasi.",
        "Menetapkan aturan mutlak bahwa identitas pengguna (NIM/User ID) hanya diekstraksi dari klaim token autentikasi yang terverifikasi, bukan dari parameter body atau header klien; melakukan validasi relasi keanggotaan mahasiswa pada kelas di sisi server sebelum mengeksekusi operasi basis data; menerapkan schema validation ketat.")

    add_threat_card(doc, "T06", "Presensi ganda atau manipulasi waktu klien akibat race condition", "Tampering", "P2. Validasi & Catat Presensi",
        2, 2, "4", "Sedang (Medium)",
        "Jika mahasiswa mengirimkan sejumlah permintaan presensi secara paralel dalam waktu bersamaan (race condition) atau memanipulasi stempel waktu perangkat klien, dan backend tidak menggunakan transaksi atomik, maka dapat tercipta catatan ganda atau presensi diterima melampaui jendela 5 menit.",
        "Pelanggaran aturan bisnis Aturan 2 (maksimal satu catatan per mahasiswa per sesi) dan Aturan 3 (batas kedaluwarsa 5 menit); inkonsistensi data persisten.",
        "Asumsi teknis: Penegakan integritas data unik (unique constraint) per pasang sesi dan mahasiswa pada skema basis data belum terkonfirmasi.",
        "Menerapkan unique constraint gabungan (id_sesi, id_mahasiswa) pada tabel database; menggunakan transaksi basis data dengan tingkat isolasi yang memadai (atomic execution); mengevaluasi batas kedaluwarsa sesi secara ketat hanya menggunakan jam server tersinkronisasi NTP; mengimplementasikan idempotency key pada endpoint presensi.")

    add_threat_card(doc, "T07", "SQL injection melalui parameter presensi ke database", "Tampering", "F11. P2 ke Database",
        2, 3, "6", "Tinggi (High)",
        "Jika parameter presensi (seperti token QR atau ID sesi) digabungkan langsung ke dalam kueri SQL dinamis tanpa sanitasi dan parameter binding, maka penyerang dapat menyuntikkan perintah SQL berbahaya ke Database Akademik melalui aliran data F11.",
        "Manipulasi massal data akademik, bypass validasi presensi, hingga kebocoran seluruh data rahasia institusi.",
        "Asumsi teknis: Rincian implementasi layer persistensi data (ORM vs raw query) tidak dijelaskan dalam spesifikasi awal sistem.",
        "Mewajibkan penggunaan parameterized queries (prepared statements) atau Object-Relational Mapping (ORM) yang aman secara konsisten; menerapkan validasi tipe data dan ekspresi reguler (regex) ketat pada input; mengkonfigurasi hak akses database user dengan prinsip least privilege; memantau eksekusi kueri anomali.")

    add_threat_card(doc, "T08", "Pengubahan atau penyangkalan catatan hadir tanpa jejak audit log", "Repudiation", "P2. Validasi & Catat Presensi",
        2, 2, "4", "Sedang (Medium)",
        "Jika peristiwa pencatatan presensi, pembukaan sesi, penutupan manual, atau perubahan kehadiran tidak dicatat secara menyeluruh ke dalam jejak audit, maka mahasiswa dapat menyangkal kehadirannya (atau ketidakhadirannya), dan pihak berwenang dapat mengubah data presensi tanpa dapat ditelusuri riwayatnya.",
        "Sengketa kehadiran mahasiswa tidak dapat diselesaikan secara objektif; hilangnya akuntabilitas operasional saat dilakukan audit akademik.",
        "Asumsi A5: Mekanisme logging dan audit trail belum dispesifikasikan dalam perancangan awal sistem.",
        "Membangun mekanisme centralized audit logging dengan format tamper-evident (append-only) yang mencatat stempel waktu server (NTP), identitas aktor, ID sesi, alamat IP sumber, User-Agent, serta status aksi; menerapkan pemisahan tugas administratif dan proteksi integritas log.")

    add_threat_card(doc, "T09", "Rekap memuat data pribadi berlebih ke pihak tidak berhak", "Information Disclosure", "F16. P3 ke Klien",
        2, 2, "4", "Sedang (Medium)",
        "Jika layanan backend mengirimkan data agregasi yang memuat atribut pribadi sensitif berlebih (seperti NIK, nomor ponsel, informasi keuangan, atau rekapitulasi nilai) atau mengirimkannya tanpa validasi kecocokan identitas dosen pengampu, maka pihak yang tidak berkepentingan dapat mengekstrak data pribadi mahasiswa.",
        "Pelanggaran privasi data mahasiswa dan potensi pelanggaran terhadap regulasi pelindungan data pribadi institusi.",
        "Asumsi teknis: Spesifikasi penyaringan data respons (data transfer object - DTO) belum dirancang secara minimalis.",
        "Menerapkan prinsip minimalisasi data (data minimization) pada payload respons API rekapitulasi; memisahkan DTO publik, dosen, dan admin; memvalidasi kepemilikan kelas di sisi server sebelum mengembalikan data kehadiran.")

    add_threat_card(doc, "T10", "Akses langsung atau kebocoran cadangan basis data akademik", "Information Disclosure", "DS1. Database Akademik",
        1, 3, "3", "Sedang (Medium)",
        "Jika file cadangan (backup) basis data akademik disimpan tanpa enkripsi, atau kontrol akses pada segmen internal basis data mengalami miskonfigurasi, maka pihak internal yang tidak berwenang dapat membaca seluruh catatan kehadiran dan identitas pengguna kampus.",
        "Kebocoran massal data akademik institusional; pelanggaran kerahasiaan skala besar.",
        "Asumsi A3: Database Akademik berada dalam jaringan internal privat dan hanya diakses oleh Backend Presensi serta DBA resmi.",
        "Menerapkan enkripsi data saat istirahat (encryption at rest - TDE); membatasi akses administratif ke server database hanya melalui bastion host berautentikasi kunci SSH dan MFA; mengamankan dan mengenkripsi seluruh file backup basis data; audit berkala terhadap hak akses pengguna basis data.")

    add_threat_card(doc, "T11", "Lonjakan permintaan presensi serentak melumpuhkan backend", "Denial of Service", "P2. Validasi & Catat Presensi",
        3, 2, "6", "Tinggi (High)",
        "Karena sesi presensi perkuliahan hanya dibuka selama 5 menit, ratusan mahasiswa di berbagai kelas akan mengirimkan permintaan presensi secara bersamaan pada waktu yang sama. Jika backend tidak sanggup menangani lonjakan koneksi konkuren, layanan akan hang atau crash.",
        "Layanan presensi terhenti total; mahasiswa yang hadir secara fisik gagal mencatatkan kehadirannya tepat waktu.",
        "Asumsi teknis: Beban puncak lalu lintas jaringan terkonsentrasi secara ekstrem dalam jendela waktu singkat 5 menit.",
        "Menerapkan arsitektur antrean pesan (message queue) asynchronous untuk memproses pencatatan transaksi presensi ke database; memanfaatkan in-memory cache (Redis) untuk validasi status sesi dan token QR aktif tanpa membebani basis data relasional; mengonfigurasi rate limiting per IP/subnet; menyediakan mekanisme fallback perpanjangan sesi darurat oleh dosen.")

    add_threat_card(doc, "T12", "Layanan Autentikasi Kampus lumpuh menghentikan presensi", "Denial of Service", "E4. Layanan Autentikasi",
        2, 3, "6", "Tinggi (High)",
        "Jika Layanan Autentikasi Kampus (SSO) mengalami kegagalan sistem, pemeliharaan mendadak, atau serangan siber, maka seluruh alur login mahasiswa dan dosen akan terputus, dan backend tidak dapat melakukan validasi token secara daring (online introspeksi).",
        "Presensi seluruh perkuliahan di lingkungan kampus lumpuh total; proses akademik terganggu parah.",
        "Asumsi teknis: Backend Presensi diasumsikan sangat bergantung pada verifikasi token daring secara langsung ke server SSO.",
        "Menerapkan validasi token SSO secara asinkron dan lokal menggunakan verifikasi kunci publik (caching JSON Web Key Set - JWKS) tanpa perlu melakukan query HTTP ke SSO setiap kali presensi; mengimplementasikan pola arsitektur circuit breaker; menyediakan SOP presensi manual darurat berbasis berita acara kelas.")

    add_threat_card(doc, "T13", "Dosen non-pengampu atau pihak lain membuka sesi kelas", "Elevation of Privilege", "P1. Kelola Sesi & QR",
        2, 3, "6", "Tinggi (High)",
        "Jika proses P1 hanya memeriksa apakah aktor memiliki atribut peran 'dosen' tanpa mencocokkan identitas dosen tersebut terhadap daftar dosen pengampu resmi kelas terkait pada Database Akademik, maka dosen lain atau mahasiswa yang memanipulasi klaim peran dapat membuka atau menutup sesi kelas semaunya.",
        "Pelanggaran mendasar terhadap Aturan Bisnis 1; timbulnya sesi perkuliahan palsu dan kekacauan administrasi akademik.",
        "Asumsi teknis: Pengecekan otorisasi berbasis kepemilikan objek (object-level authorization) belum diterapkan secara ketat.",
        "Menegakkan kontrol akses berbasis atribut dan kepemilikan objek (ABAC / strict ownership check) di sisi server: setiap pemanggilan endpoint pembukaan sesi wajib memverifikasi relasi antara NIP dosen pada token dengan kolom dosen_pengampu pada jadwal kelas di basis data; menerapkan kebijakan tolak secara default (deny-by-default).")

    add_threat_card(doc, "T14", "Akses data rekapitulasi lintas kelas oleh peran tidak berhak", "Elevation of Privilege", "P3. Muat Rekap Kehadiran",
        2, 3, "6", "Tinggi (High)",
        "Jika proses P3 tidak membatasi pengambilan data rekapitulasi berdasarkan relasi kepemilikan kelas atau tidak memisahkan endpoint administratif secara aman, maka dosen dapat melihat catatan presensi kelas dosen lain, atau mahasiswa dapat memanggil endpoint rekapitulasi untuk melihat rekapitulasi seluruh kampus.",
        "Eksfiltrasi data kehadiran institusi secara tidak sah; pelanggaran privasi antar-dosen dan antar-mahasiswa.",
        "Asumsi teknis: Penanganan pemisahan hak akses endpoint pelaporan belum diimplementasikan dengan pengujian otorisasi yang ketat.",
        "Menerapkan pemisahan endpoint yang jelas antara rekapitulasi dosen pengampu dan rekapitulasi agregat administrator kampus; mengunci kueri basis data pada P3 agar membatasi scope data hanya untuk kelas yang diajar dosen tersebut; menerapkan pengujian otorisasi otomatis (authorization testing) pada pipeline CI/CD.")

    # ==========================
    # 5. MATRIKS RISIKO DAN PRIORITAS
    # ==========================
    add_heading_1(doc, "5. Matriks Risiko dan Prioritas Penanganan")
    add_heading_2(doc, "5.1 Matriks Evaluasi Risiko 3×3")
    add_body_paragraph(doc,
        "Pemetaan 14 ancaman ke dalam matriks evaluasi risiko 3×3 berdasarkan nilai Kemungkinan (Likelihood) dan Dampak (Impact) disajikan pada tabel berikut:")

    rm_table = doc.add_table(rows=4, cols=4)
    rm_headers = ["Tingkat Kemungkinan", "Dampak 1 (Rendah)", "Dampak 2 (Sedang)", "Dampak 3 (Tinggi)"]
    rm_data = [
        ["Kemungkinan 3 (Tinggi)", "—", "T01, T03, T11\n(Skor 6 — Tinggi)", "—"],
        ["Kemungkinan 2 (Sedang)", "—", "T04, T06, T08, T09\n(Skor 4 — Sedang)", "T02, T05, T07, T12, T13, T14\n(Skor 6 — Tinggi)"],
        ["Kemungkinan 1 (Rendah)", "—", "—", "T10\n(Skor 3 — Sedang)"]
    ]
    format_table(rm_table, [1.8, 1.4, 1.6, 1.7], rm_headers, rm_data)

    add_heading_2(doc, "5.2 Strategi dan Urutan Prioritas Remediasi")
    add_body_paragraph(doc,
        "Berdasarkan tingkat keparahan risiko dan kelayakan implementasi teknis, urutan prioritas penanganan ditetapkan sebagai berikut:")
    add_bullet_point(doc, "T13 (Otorisasi Sesi Dosen), T14 (Otorisasi Akses Rekap), T05 (Validasi Parameter Presensi), dan T07 (Pencegahan SQL Injection). Kelompok ini memerlukan penegakan kontrol validasi dan otorisasi sisi server yang fundamental dengan kompleksitas rendah namun berdampak krusial terhadap integritas sistem.", "Prioritas 1 — Integritas Otorisasi & Validasi Input (Skor 6 — Tinggi): ")
    add_bullet_point(doc, "T01 (Titip Absen Akun), T03 (Penyebaran Foto QR Jarak Jauh), dan T02 (Pencurian Token Klien). Menjadi risiko paling nyata dan kerap terjadi pada operasional sehari-hari perkuliahan. Penanganannya membutuhkan rolling QR dinamis (TOTP), pembatasan subnet kampus, dan pengikatan sesi ke perangkat.", "Prioritas 2 — Keaslian Presensi Fisik (Skor 6 — Tinggi): ")
    add_bullet_point(doc, "T11 (Ketahanan terhadap Lonjakan Beban Presensi 5 Menit) dan T12 (Dependensi Ketersediaan SSO Kampus). Berfokus pada pencegahan kelumpuhan layanan melalui caching lokal kunci publik JWKS dan arsitektur antrean.", "Prioritas 3 — Ketahanan Ketersediaan Layanan (Skor 6 — Tinggi): ")
    add_bullet_point(doc, "T08 (Penyediaan Audit Trail), T04 (Penguatan Entropi Token QR), T06 (Pencegahan Kondisi Balapan Database), T09 (Minimalisasi Data Rekapitulasi), dan T10 (Pengamanan Cadangan Database). Berperan sebagai kontrol pelengkap dan penguatan postur keamanan berkelanjutan.", "Prioritas 4 — Pertahanan Pendukung & Tata Kelola (Skor 3–4 — Sedang): ")

    # ==========================
    # 6. PEMETAAN KE THREAT DRAGON
    # ==========================
    add_heading_1(doc, "6. Pemetaan ke OWASP Threat Dragon")
    add_body_paragraph(doc,
        "Guna memastikan konsistensi antara dokumen laporan formal dengan berkas model Threat Dragon (.json), seluruh entri ancaman ditautkan secara presisi pada elemen DFD terkait:")

    map_table = doc.add_table(rows=10, cols=3)
    map_headers = ["Elemen pada Threat Dragon", "Daftar ID Ancaman", "Kategori STRIDE & Status"]
    map_data = [
        ["E1. Mahasiswa (Browser)", "T01", "Spoofing — Status: Open"],
        ["F3. Mahasiswa ke P2", "T02, T05", "Spoofing (T02), Tampering (T05) — Status: Open"],
        ["P1. Kelola Sesi & QR", "T04, T13", "Spoofing (T04), Elevation of Privilege (T13) — Status: Open"],
        ["P2. Validasi & Catat Presensi", "T03, T06, T08, T11", "Spoofing (T03), Tampering (T06), Repudiation (T08), DoS (T11) — Status: Open"],
        ["F11. P2 ke Database", "T07", "Tampering — Status: Open"],
        ["F16. P3 ke Dosen/Admin", "T09", "Information Disclosure — Status: Open"],
        ["DS1. Database Akademik", "T10", "Information Disclosure — Status: Open"],
        ["E4. Layanan Autentikasi Kampus", "T12", "Denial of Service — Status: Open"],
        ["P3. Muat Rekap Kehadiran", "T14", "Elevation of Privilege — Status: Open"]
    ]
    format_table(map_table, [2.3, 1.8, 2.4], map_headers, map_data)

    # ==========================
    # 7. PEMBAHASAN DAN KESIMPULAN
    # ==========================
    add_heading_1(doc, "7. Pembahasan Reflektif dan Kesimpulan")
    add_heading_2(doc, "7.1 Pembahasan Reflektif")
    add_body_paragraph(doc,
        "Analisis threat modeling menghasilkan sejumlah temuan teknis reflektif yang penting bagi pengembangan perangkat lunak aman:")
    add_bullet_point(doc, "Meskipun mitigasi teknis (seperti validasi dan otorisasi) diimplementasikan pada Backend Presensi, ancaman tetap ditautkan pada elemen spesifik tempat kerentanan atau serangan pertama kali terjadi (misalnya manipulasi parameter pada aliran data F3 dan injeksi SQL pada F11). Hal ini menjaga kejelasan konteks vektor serangan.", "Lokasi Penempatan Ancaman: ")
    add_bullet_point(doc, "Banyak kegagalan keamanan terjadi pada saat data berpindah melintasi batas kepercayaan (trust boundary), terutama dari perangkat klien mahasiswa (TB1) menuju lingkungan backend internal (TB2). Aliran data merupakan representasi saluran komunikasi yang rentan disadap atau dimanipulasi.", "Pentingnya Analisis Aliran Data (Data Flow): ")
    add_bullet_point(doc, "Elemen aktor dosen (E2) dan admin (E3) tidak memiliki entri ancaman langsung bukan karena elemen tersebut kebal terhadap ancaman, melainkan karena potensi risiko terkait wewenang dan manipulasi telah dicakup secara terstruktur pada proses backend P1 dan P3. Pendekatan ini mencegah redundansi dokumentasi.", "Elemen Tanpa Entri Ancaman Langsung: ")
    add_bullet_point(doc, "Seluruh ancaman didokumentasikan dengan Status: Open karena rekomendasi mitigasi yang dirumuskan merupakan usulan kontrol perancangan arsitektur. Status Mitigated hanya boleh diberikan setelah kontrol keamanan selesai diimplementasikan dalam kode program dan lolos verifikasi pengujian keamanan.", "Penetapan Status Ancaman (Open vs Mitigated): ")

    add_heading_2(doc, "7.2 Kesimpulan")
    add_body_paragraph(doc,
        "Pemodelan ancaman berbasis STRIDE terhadap Sistem Presensi Perkuliahan Berbasis QR Code berhasil memetakan arsitektur sistem ke dalam 4 entitas aktor, 3 proses pengolah, 1 penyimpan data persisten, 16 aliran data, serta 4 batas kepercayaan. Dari pemodelan tersebut, teridentifikasi 14 ancaman keamanan yang mencakup seluruh spektrum STRIDE, terdiri atas 9 ancaman berisiko tinggi (skor 6) dan 5 ancaman berisiko sedang (skor 3–4).")
    add_body_paragraph(doc,
        "Vektor risiko paling dominan terkonsentrasi pada dua domain utama: (1) keaslian presensi fisik akibat praktik titip absen dan pembagian kode QR ke luar kelas, serta (2) penegakan otorisasi sisi server pada pembukaan sesi dan pemuatan rekapitulasi data. Rekomendasi kontrol mitigasi yang dirumuskan dalam laporan ini menjadi panduan arsitektural resmi bagi tim rekayasa perangkat lunak untuk mengamankan implementasi sistem presensi sebelum masuk ke fase produksi.")

    # ==========================
    # REFERENSI
    # ==========================
    add_heading_1(doc, "Referensi Standar")
    add_bullet_point(doc, "OWASP Foundation. (2024). OWASP Threat Dragon Project. https://owasp.org/projects/threat-dragon")
    add_bullet_point(doc, "OWASP Foundation. (2023). Threat Dragon Documentation and User Guide. https://www.threatdragon.com/docs/")
    add_bullet_point(doc, "OWASP Cheat Sheet Series. (2024). Threat Modeling Cheat Sheet. https://cheatsheetseries.owasp.org/cheatsheets/Threat_Modeling_Cheat_Sheet.html")
    add_bullet_point(doc, "Shostack, A. (2014). Threat Modeling: Designing for Security. John Wiley & Sons.")
    add_bullet_point(doc, "National Institute of Standards and Technology (NIST). (2020). SP 800-53 Rev. 5: Security and Privacy Controls for Information Systems and Organizations.")

    # ==========================
    # LAMPIRAN BUKTI PRAKTIKUM
    # ==========================
    add_heading_1(doc, "Lampiran: Tangkapan Layar Bukti Praktikum (OWASP Threat Dragon)")
    add_body_paragraph(doc,
        "Berikut merupakan dokumentasi tangkapan layar (screenshot) sebagai bukti pelaksanaan aktivitas threat modeling pada perangkat lunak OWASP Threat Dragon v2:")

    lamp_table = doc.add_table(rows=7, cols=4)
    lamp_headers = ["Kode Lampiran", "Bagian / Layar yang Ditangkap", "Penanggung Jawab", "Keterangan Bukti"]
    lamp_data = [
        ["CAP-01", "Halaman Detail Model (Title & Contributors)", "Fatir Gibran (103112430153)", "Memperlihatkan identitas model dan 4 kontributor."],
        ["CAP-02", "Kanvas Utuh Diagram DFD Level-1", "Fatir Gibran (103112430153)", "Memperlihatkan 8 elemen, 16 aliran data, dan 4 boundary."],
        ["CAP-03", "Modal Input Ancaman Spoofing/Tampering", "Tio Armani (103112430225)", "Memperlihatkan pengisian skenario & mitigasi T01/T03."],
        ["CAP-04", "Modal Input Ancaman R, I, D, E", "Chadafya Putra Zulfikar (103112430173)", "Memperlihatkan pengisian skenario & mitigasi T08/T11/T13."],
        ["CAP-05", "Pengaturan Opsi Fitur Laporan (Report)", "Tedy Permana Putra (103112430157)", "Memperlihatkan opsi centang diagram, empty elements, & properties."],
        ["CAP-06", "Pratinjau Hasil Ekspor Laporan PDF", "Tedy Permana Putra (103112430157)", "Memperlihatkan bukti terbitnya berkas PDF resmi dari tool."]
    ]
    format_table(lamp_table, [1.2, 2.3, 1.8, 1.2], lamp_headers, lamp_data)

    def add_capture_placeholder(cap_code, cap_title, img_path=None):
        box_table = doc.add_table(rows=2, cols=1)
        box_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        c_hdr = box_table.cell(0, 0)
        set_cell_background(c_hdr, "2D3748")
        set_cell_margins(c_hdr, 80, 80, 120, 120)
        p = c_hdr.paragraphs[0]
        r = p.add_run(f"[{cap_code}] {cap_title}")
        r.font.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

        c_body = box_table.cell(1, 0)
        set_cell_background(c_body, "F7FAFC")
        set_cell_margins(c_body, 120, 120, 140, 140)
        set_table_borders(box_table, "CBD5E0")
        pb = c_body.paragraphs[0]
        pb.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if img_path and os.path.exists(img_path):
            rb_img = pb.add_run()
            rb_img.add_picture(img_path, width=Inches(6.0))
        else:
            rb_txt = pb.add_run(f"[Tempelkan screenshot {cap_title} di area ini]")
            rb_txt.font.italic = True
            rb_txt.font.color.rgb = RGBColor(160, 174, 192)
            rb_txt.font.size = Pt(9)
        c_hdr.width = Inches(6.5)
        c_body.width = Inches(6.5)
        doc.add_paragraph().paragraph_format.space_after = Pt(6)

    add_capture_placeholder("CAP-01", "Tangkapan Layar: Halaman Detail Model dan Kontributor (Fatir Gibran)", "cap_01_summary.png")
    add_capture_placeholder("CAP-02", "Tangkapan Layar: Kanvas Utuh DFD Level-1 Sistem Presensi (Fatir Gibran)", "cap_02_dfd.png")
    add_capture_placeholder("CAP-03", "Tangkapan Layar: Form Entri Ancaman Spoofing / Tampering (Tio Armani)")
    add_capture_placeholder("CAP-04", "Tangkapan Layar: Form Entri Ancaman Repudiation / DoS / EoP (Chadafya Putra Zulfikar)")
    add_capture_placeholder("CAP-05", "Tangkapan Layar: Pengaturan Centang Ekspor Report (Tedy Permana Putra)")
    add_capture_placeholder("CAP-06", "Tangkapan Layar: Hasil Ekspor Laporan PDF Tool (Tedy Permana Putra)")

    output_path = "Laporan_Threat_Modeling_Kasus_A_Presensi_QR.docx"
    doc.save(output_path)
    print(f"File laporan berhasil disimpan di {output_path}")

if __name__ == "__main__":
    build_laporan()
