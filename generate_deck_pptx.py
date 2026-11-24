# -*- coding: utf-8 -*-
"""
Script to generate Presentasi_Threat_Modeling_Kasus_A_Presensi_QR.pptx
Professional 16:9 widescreen presentation deck using python-pptx.
High aesthetic quality, clean card layouts, strong typography, structured tables, and embedded diagrams.
"""

import os
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ==========================================
# COLOR PALETTE
# ==========================================
NAVY_DARK = RGBColor(15, 23, 42)       # #0F172A
SLATE_DARK = RGBColor(30, 41, 59)      # #1E293B
SLATE_MUTED = RGBColor(100, 116, 139)  # #64748B
BG_LIGHT = RGBColor(248, 250, 252)     # #F8FAFC
CARD_BG = RGBColor(255, 255, 255)      # #FFFFFF
CARD_BORDER = RGBColor(226, 232, 240)  # #E2E8F0

BLUE_PRIMARY = RGBColor(37, 99, 235)   # #2563EB
BLUE_LIGHT = RGBColor(239, 246, 255)   # #EFF6FF
CYAN_ACCENT = RGBColor(14, 165, 233)   # #0EA5E9

RED_SEV = RGBColor(220, 38, 38)        # #DC2626
RED_BG = RGBColor(254, 242, 242)       # #FEF2F2

AMBER_SEV = RGBColor(217, 119, 6)      # #D97706
AMBER_BG = RGBColor(254, 243, 199)     # #FEF3C7

GREEN_SEV = RGBColor(22, 163, 74)      # #16A34A
GREEN_BG = RGBColor(240, 253, 244)     # #F0FDF4

PURPLE_SEV = RGBColor(147, 51, 234)    # #9333EA
PURPLE_BG = RGBColor(250, 245, 255)    # #FAF5FF

TEXT_MAIN = RGBColor(30, 41, 59)       # #1E293B
TEXT_MUTED = RGBColor(100, 116, 139)   # #64748B
TEXT_WHITE = RGBColor(255, 255, 255)

FONT_HEADING = "Trebuchet MS"
FONT_BODY = "Calibri"

def create_slide_header(slide, title_text, category_text="KEAMANAN SIBER (CAK3CAB3) • STUDI KASUS A"):
    """Creates a consistent, professional header across content slides."""
    # Category / Breadcrumb
    cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.35))
    tf_c = cat_box.text_frame
    tf_c.word_wrap = True
    p_c = tf_c.paragraphs[0]
    p_c.text = category_text.upper()
    p_c.font.name = FONT_BODY
    p_c.font.size = Pt(9.5)
    p_c.font.bold = True
    p_c.font.color.rgb = BLUE_PRIMARY

    # Slide Title
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.68), Inches(11.7), Inches(0.65))
    tf_t = title_box.text_frame
    tf_t.word_wrap = True
    p_t = tf_t.paragraphs[0]
    p_t.text = title_text
    p_t.font.name = FONT_HEADING
    p_t.font.size = Pt(22)
    p_t.font.bold = True
    p_t.font.color.rgb = NAVY_DARK

def add_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=CARD_BORDER, border_width=1):
    """Draws a subtle rounded rectangle container card."""
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = Pt(border_width)
    else:
        shape.line.fill.background()
    return shape

def build_presentation():
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # ==========================================
    # SLIDE 1: COVER / TITLE SLIDE (Dark Elegance)
    # ==========================================
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = NAVY_DARK
    bg1.line.fill.background()

    # Decorative accent glow line
    acc1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.1), Inches(1.5), Inches(0.06))
    acc1.fill.solid()
    acc1.fill.fore_color.rgb = CYAN_ACCENT
    acc1.line.fill.background()

    # Badge: Mata Kuliah & Kasus
    b_tag = s1.shapes.add_textbox(Inches(0.8), Inches(1.3), Inches(10), Inches(0.4))
    p_bt = b_tag.text_frame.paragraphs[0]
    p_bt.text = "KEAMANAN SIBER (CAK3CAB3) • SEMESTER GANJIL 2026/2027"
    p_bt.font.name = FONT_BODY
    p_bt.font.size = Pt(11)
    p_bt.font.bold = True
    p_bt.font.color.rgb = CYAN_ACCENT

    # Main Title
    t_box = s1.shapes.add_textbox(Inches(0.8), Inches(1.7), Inches(11.7), Inches(1.5))
    tf_m = t_box.text_frame
    tf_m.word_wrap = True
    p_m = tf_m.paragraphs[0]
    p_m.text = "Threat Modeling Berbasis STRIDE\nMenggunakan OWASP Threat Dragon"
    p_m.font.name = FONT_HEADING
    p_m.font.size = Pt(32)
    p_m.font.bold = True
    p_m.font.color.rgb = TEXT_WHITE

    # Subtitle
    sub_box = s1.shapes.add_textbox(Inches(0.8), Inches(3.25), Inches(11.7), Inches(0.6))
    p_sub = sub_box.text_frame.paragraphs[0]
    p_sub.text = "Studi Kasus A: Sistem Presensi Perkuliahan Berbasis QR Code"
    p_sub.font.name = FONT_BODY
    p_sub.font.size = Pt(16)
    p_sub.font.color.rgb = RGBColor(203, 213, 225) # Slate 300

    # Team Members 4-Column Card Grid
    members = [
        ("Fatir Gibran", "103112430153", "Ketua Tim & Pemodel DFD"),
        ("Tio Armani", "103112430225", "Analis Spoofing & Tampering"),
        ("Chadafya Putra Zulfikar", "103112430173", "Analis R, I, D, E"),
        ("Tedy Permana Putra", "103112430157", "Penilai Risiko & QA Editor")
    ]

    card_w = Inches(2.7)
    card_h = Inches(1.8)
    card_gap = Inches(0.3)
    start_x = Inches(0.8)
    start_y = Inches(4.3)

    for i, (name, nim, role) in enumerate(members):
        cx = start_x + i * (card_w + card_gap)
        c_shape = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, start_y, card_w, card_h)
        c_shape.fill.solid()
        c_shape.fill.fore_color.rgb = SLATE_DARK
        c_shape.line.color.rgb = RGBColor(51, 65, 85) # Slate 700
        c_shape.line.width = Pt(1)

        tb_mem = s1.shapes.add_textbox(cx + Inches(0.15), start_y + Inches(0.2), card_w - Inches(0.3), card_h - Inches(0.3))
        tf_mem = tb_mem.text_frame
        tf_mem.word_wrap = True

        p1 = tf_mem.paragraphs[0]
        p1.text = name
        p1.font.name = FONT_HEADING
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = TEXT_WHITE

        p2 = tf_mem.add_paragraph()
        p2.text = f"NIM: {nim}"
        p2.font.name = FONT_BODY
        p2.font.size = Pt(10)
        p2.font.color.rgb = CYAN_ACCENT

        p3 = tf_mem.add_paragraph()
        p3.text = role
        p3.font.name = FONT_BODY
        p3.font.size = Pt(9.5)
        p3.font.color.rgb = RGBColor(148, 163, 184) # Slate 400

    # Reviewer footer tag
    rf_box = s1.shapes.add_textbox(Inches(0.8), Inches(6.4), Inches(11.7), Inches(0.4))
    p_rf = rf_box.text_frame.paragraphs[0]
    p_rf.text = "Reviewer: Dosen Praktikum Keamanan Siber • Alat: OWASP Threat Dragon v2.6.2"
    p_rf.font.name = FONT_BODY
    p_rf.font.size = Pt(10)
    p_rf.font.color.rgb = RGBColor(100, 116, 139)

    # ==========================================
    # SLIDE 2: LATAR BELAKANG & RUANG LINGKUP
    # ==========================================
    s2 = prs.slides.add_slide(blank_layout)
    create_slide_header(s2, "Latar Belakang & Ruang Lingkup Sistem Presensi QR")

    # Card 1: Konsep & Alur Sistem (Kiri)
    add_card(s2, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.4))
    tb_c1 = s2.shapes.add_textbox(Inches(1.1), Inches(1.7), Inches(5.0), Inches(5.0))
    tf1 = tb_c1.text_frame
    tf1.word_wrap = True

    p = tf1.paragraphs[0]
    p.text = "Karakteristik & Alur Bisnis Sistem"
    p.font.name = FONT_HEADING
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY_DARK

    bullets1 = [
        ("Tujuan Sistem: ", "Mencatat bukti kehadiran perkuliahan mahasiswa secara digital di ruang kelas melalui pemindaian QR code dinamis."),
        ("Pembukaan Sesi: ", "Dosen pengampu login via SSO kampus, membuka sesi perkuliahan untuk kelas yang diampunya, lalu menayangkan kode QR di proyektor."),
        ("Jendela Sesi 5 Menit: ", "Sesi presensi dibatasi masa aktifnya hanya selama 5 menit untuk memastikan kehadiran fisik pada waktu kelas."),
        ("Alur Mahasiswa: ", "Mahasiswa login via SSO, memindai kode QR, dan aplikasi mengirimkan ID sesi serta token mahasiswa ke Backend Presensi."),
        ("Validasi & Pencatatan: ", "Backend memeriksa token SSO, mengecek sesi masih aktif, memverifikasi relasi kelas, dan membukukan kehadiran ke Database Akademik.")
    ]
    for b_title, b_desc in bullets1:
        p = tf1.add_paragraph()
        p.space_before = Pt(8)
        r1 = p.add_run()
        r1.text = "• " + b_title
        r1.font.bold = True
        r1.font.size = Pt(10.5)
        r1.font.color.rgb = BLUE_PRIMARY
        r2 = p.add_run()
        r2.text = b_desc
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = TEXT_MAIN

    # Card 2: Batasan Lingkup Analisis (Kanan)
    add_card(s2, Inches(6.8), Inches(1.5), Inches(5.7), Inches(5.4))
    tb_c2 = s2.shapes.add_textbox(Inches(7.1), Inches(1.7), Inches(5.1), Inches(5.0))
    tf2 = tb_c2.text_frame
    tf2.word_wrap = True

    p = tf2.paragraphs[0]
    p.text = "Batasan Ruang Lingkup Analisis"
    p.font.name = FONT_HEADING
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY_DARK

    p_in = tf2.add_paragraph()
    p_in.space_before = Pt(10)
    p_in.text = "DALAM LINGKUP ANALISIS (IN-SCOPE)"
    p_in.font.bold = True
    p_in.font.size = Pt(11)
    p_in.font.color.rgb = GREEN_SEV

    in_items = [
        "Pembukaan dan penutupan sesi kelas oleh dosen pengampu.",
        "Pembangkitan, penayangan, dan pemindaian token/kode QR sesi.",
        "Validasi keanggotaan kelas, pencegahan presensi ganda, & pencatatan.",
        "Pertukaran token otorisasi dengan Layanan Autentikasi Kampus (SSO).",
        "Pemuatan rekapitulasi kehadiran oleh dosen dan admin."
    ]
    for it in in_items:
        p = tf2.add_paragraph()
        p.text = "✔ " + it
        p.font.size = Pt(10)
        p.font.color.rgb = TEXT_MAIN

    p_out = tf2.add_paragraph()
    p_out.space_before = Pt(14)
    p_out.text = "DI LUAR LINGKUP ANALISIS (OUT-OF-SCOPE)"
    p_out.font.bold = True
    p_out.font.size = Pt(11)
    p_out.font.color.rgb = RED_SEV

    out_items = [
        "Pendaftaran mata kuliah (KRS) & penginputan nilai semester.",
        "Keamanan fisik ruang kelas & infrastruktur jaringan kabel/Wi-Fi.",
        "Pengujian penetrasi aktif (pentest) & eksploitasi kode nyata."
    ]
    for it in out_items:
        p = tf2.add_paragraph()
        p.text = "✖ " + it
        p.font.size = Pt(10)
        p.font.color.rgb = TEXT_MUTED

    # ==========================================
    # SLIDE 3: ATURAN BISNIS & ASUMSI TEKNIS
    # ==========================================
    s3 = prs.slides.add_slide(blank_layout)
    create_slide_header(s3, "Aturan Bisnis & 5 Asumsi Teknis Perancangan")

    # 3 Aturan Bisnis (Atas)
    add_card(s3, Inches(0.8), Inches(1.5), Inches(11.7), Inches(2.2))
    tb_br = s3.shapes.add_textbox(Inches(1.1), Inches(1.65), Inches(11.1), Inches(1.9))
    tf_br = tb_br.text_frame
    tf_br.word_wrap = True

    p = tf_br.paragraphs[0]
    p.text = "3 Aturan Bisnis Utama (Business Rules)"
    p.font.name = FONT_HEADING
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = BLUE_PRIMARY

    br_items = [
        ("Aturan 1 (Otorisasi Dosen Pengampu): ", "Hanya dosen yang terdaftar resmi sebagai pengampu mata kuliah terkait yang memiliki hak membuka, memantau, dan menutup sesi presensi."),
        ("Aturan 2 (Keunikan Presensi Mahasiswa): ", "Satu akun mahasiswa hanya diperbolehkan memiliki paling banyak satu catatan kehadiran yang sah untuk setiap sesi perkuliahan aktif."),
        ("Aturan 3 (Batas Kedaluwarsa Sesi): ", "Sesi presensi memiliki batas masa aktif ketat selama 5 menit, namun dosen berhak menutup sesi lebih awal jika kelas telah lengkap.")
    ]
    for br_title, br_desc in br_items:
        p = tf_br.add_paragraph()
        p.space_before = Pt(4)
        r1 = p.add_run()
        r1.text = "• " + br_title
        r1.font.bold = True
        r1.font.size = Pt(10.5)
        r1.font.color.rgb = NAVY_DARK
        r2 = p.add_run()
        r2.text = br_desc
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = TEXT_MAIN

    # 5 Asumsi Teknis (Bawah: 5 Kolom Card)
    asumsi_items = [
        ("A1: Enkripsi TLS", "Komunikasi browser ke backend & backend ke SSO wajib menggunakan HTTPS/TLS terenkripsi."),
        ("A2: Standar SSO", "Layanan Autentikasi Kampus menerbitkan token bertanda tangan berbatas waktu (JWT)."),
        ("A3: Isolasi Database", "Database Akademik berada di jaringan internal terisolasi, hanya diakses Backend & DBA."),
        ("A4: QR Statis 5 Menit", "Token QR berisi ID sesi & string acak, belum menerapkan rolling refresh otomatis."),
        ("A5: Kontrol Log Audit", "Audit log terperinci & rate limiting belum dijelaskan (tidak dianggap pasti tidak ada).")
    ]

    a_w = Inches(2.18)
    a_h = Inches(2.9)
    a_gap = Inches(0.2)
    start_ay = Inches(3.9)

    for i, (a_title, a_desc) in enumerate(asumsi_items):
        ax = Inches(0.8) + i * (a_w + a_gap)
        add_card(s3, ax, start_ay, a_w, a_h, bg_color=CARD_BG, border_color=CARD_BORDER)
        
        tb_a = s3.shapes.add_textbox(ax + Inches(0.12), start_ay + Inches(0.15), a_w - Inches(0.24), a_h - Inches(0.3))
        tf_a = tb_a.text_frame
        tf_a.word_wrap = True

        pa1 = tf_a.paragraphs[0]
        pa1.text = a_title
        pa1.font.name = FONT_HEADING
        pa1.font.size = Pt(11)
        pa1.font.bold = True
        pa1.font.color.rgb = NAVY_DARK

        pa2 = tf_a.add_paragraph()
        pa2.space_before = Pt(6)
        pa2.text = a_desc
        pa2.font.name = FONT_BODY
        pa2.font.size = Pt(9.5)
        pa2.font.color.rgb = TEXT_MUTED

    # ==========================================
    # SLIDE 4: DATA FLOW DIAGRAM (DFD) LEVEL-1
    # ==========================================
    s4 = prs.slides.add_slide(blank_layout)
    create_slide_header(s4, "Arsitektur Data Flow Diagram (DFD) Level-1")

    # Kiri: Inventarisasi Komponen DFD
    add_card(s4, Inches(0.8), Inches(1.5), Inches(4.5), Inches(5.4))
    tb_dfd = s4.shapes.add_textbox(Inches(1.0), Inches(1.7), Inches(4.1), Inches(5.0))
    tf_d = tb_dfd.text_frame
    tf_d.word_wrap = True

    p = tf_d.paragraphs[0]
    p.text = "Struktur Komponen Sistem"
    p.font.name = FONT_HEADING
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY_DARK

    comps = [
        ("4 External Entities (Aktor):", "E1: Mahasiswa (Browser)\nE2: Dosen Pengampu (Browser)\nE3: Admin Akademik (Browser)\nE4: Layanan Autentikasi Kampus (SSO)"),
        ("3 Proses Backend Presensi:", "P1: Kelola Sesi & Pembangkitan QR 5m\nP2: Validasi, Cek Relasi, & Catat Presensi\nP3: Muat Rekapitulasi Kehadiran Kelas"),
        ("1 Penyimpan Data Persisten:", "DS1: Database Akademik (Akun, Jadwal, Sesi, Keanggotaan, Presensi)"),
        ("16 Aliran Data (Data Flow):", "F1 s/d F16 menghubungkan interaksi pengguna, backend, basis data, dan SSO.")
    ]
    for c_title, c_desc in comps:
        p = tf_d.add_paragraph()
        p.space_before = Pt(8)
        p.text = c_title
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = BLUE_PRIMARY

        p2 = tf_d.add_paragraph()
        p2.text = c_desc
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = TEXT_MAIN

    # Kanan: Gambar Diagram DFD
    add_card(s4, Inches(5.6), Inches(1.5), Inches(6.9), Inches(5.4))
    if os.path.exists("dfd_kasus_a.png"):
        s4.shapes.add_picture("dfd_kasus_a.png", Inches(5.75), Inches(1.65), width=Inches(6.6))

    # ==========================================
    # SLIDE 5: BATAS KEPERCAYAAN (TRUST BOUNDARIES)
    # ==========================================
    s5 = prs.slides.add_slide(blank_layout)
    create_slide_header(s5, "Zona Keamanan & 4 Batas Kepercayaan (Trust Boundaries)")

    tb_cards = [
        ("TB1: Perangkat Pengguna / Internet", "Melingkupi: E1 Mahasiswa, E2 Dosen, E3 Admin", 
         "Zona Publik / Tidak Terpercaya (Untrusted). Perangkat browser klien berada di luar kontrol institusi, rentan terhadap malware atau modifikasi parameter. Semua input yang masuk wajib divalidasi ketat di server.", 
         AMBER_SEV, AMBER_BG),
        ("TB-Auth: Layanan Autentikasi Kampus", "Melingkupi: E4 Layanan SSO / IdP", 
         "Zona Eksternal Terpercaya (Trusted External IdP). Dipercaya penuh untuk membuktikan identitas pengguna, namun menjadi dependensi kritis (single point of failure) bagi ketersediaan operasional.", 
         PURPLE_SEV, PURPLE_BG),
        ("TB2: Backend Presensi (Aplikasi Internal)", "Melingkupi: P1, P2, P3", 
         "Zona Pemrosesan Aplikasi Terpercaya (Trusted Application Tier). Tempat penegakan logika bisnis, validasi otorisasi kepemilikan kelas, evaluasi masa aktif sesi 5 menit, dan pencegahan presensi ganda.", 
         BLUE_PRIMARY, BLUE_LIGHT),
        ("TB3: Zona Data Internal", "Melingkupi: DS1 Database Akademik", 
         "Zona Data Berintegritas Tinggi (High Trust Data Tier). Terisolasi di jaringan privat internal. Akses langsung dibatasi hanya untuk Backend Presensi terotentikasi dan DBA resmi.", 
         GREEN_SEV, GREEN_BG)
    ]

    tb_w = Inches(5.6)
    tb_h = Inches(2.5)

    coords = [
        (Inches(0.8), Inches(1.5)),
        (Inches(6.8), Inches(1.5)),
        (Inches(0.8), Inches(4.3)),
        (Inches(6.8), Inches(4.3))
    ]

    for i, (title, scope, desc, color, bg) in enumerate(tb_cards):
        x, y = coords[i]
        add_card(s5, x, y, tb_w, tb_h, bg_color=CARD_BG, border_color=color, border_width=1.5)

        tb_box = s5.shapes.add_textbox(x + Inches(0.2), y + Inches(0.18), tb_w - Inches(0.4), tb_h - Inches(0.35))
        tf = tb_box.text_frame
        tf.word_wrap = True

        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.name = FONT_HEADING
        p1.font.size = Pt(13)
        p1.font.bold = True
        p1.font.color.rgb = color

        p2 = tf.add_paragraph()
        p2.text = scope
        p2.font.bold = True
        p2.font.size = Pt(10)
        p2.font.color.rgb = NAVY_DARK

        p3 = tf.add_paragraph()
        p3.space_before = Pt(4)
        p3.text = desc
        p3.font.size = Pt(9.5)
        p3.font.color.rgb = TEXT_MAIN

    # ==========================================
    # SLIDE 6: METODOLOGI STRIDE & RUBRIK RISIKO
    # ==========================================
    s6 = prs.slides.add_slide(blank_layout)
    create_slide_header(s6, "Metodologi Analisis STRIDE & Rubrik Penilaian Risiko")

    # Kiri: 6 Kategori STRIDE
    add_card(s6, Inches(0.8), Inches(1.5), Inches(6.0), Inches(5.4))
    tb_st = s6.shapes.add_textbox(Inches(1.0), Inches(1.7), Inches(5.6), Inches(5.0))
    tf_st = tb_st.text_frame
    tf_st.word_wrap = True

    p = tf_st.paragraphs[0]
    p.text = "Kerangka Kerja STRIDE"
    p.font.name = FONT_HEADING
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY_DARK

    stride_definitions = [
        ("S - Spoofing:", "Menyamarkan identitas sebagai pengguna atau entitas lain secara tidak sah (titip absen akun, pencurian token)."),
        ("T - Tampering:", "Memodifikasi data atau parameter tanpa otorisasi (manipulasi ID mahasiswa, race condition waktu klien)."),
        ("R - Repudiation:", "Menyangkal tindakan kehadiran atau perubahan data karena ketiadaan bukti audit log yang valid."),
        ("I - Information Disclosure:", "Membocorkan informasi rahasia atau data pribadi kepada pihak yang tidak berwenang."),
        ("D - Denial of Service:", "Melumpuhkan ketersediaan layanan sehingga pengguna sah tidak dapat melakukan presensi."),
        ("E - Elevation of Privilege:", "Memperoleh wewenang melebihi haknya (dosen membuka kelas lain, akses rekap global).")
    ]
    for s_title, s_desc in stride_definitions:
        p = tf_st.add_paragraph()
        p.space_before = Pt(6)
        r1 = p.add_run()
        r1.text = s_title + " "
        r1.font.bold = True
        r1.font.size = Pt(10.5)
        r1.font.color.rgb = BLUE_PRIMARY
        r2 = p.add_run()
        r2.text = s_desc
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = TEXT_MAIN

    # Kanan: Rubrik Penilaian Risiko (Skor = L x D)
    add_card(s6, Inches(7.1), Inches(1.5), Inches(5.4), Inches(5.4))
    tb_rb = s6.shapes.add_textbox(Inches(7.3), Inches(1.7), Inches(5.0), Inches(5.0))
    tf_rb = tb_rb.text_frame
    tf_rb.word_wrap = True

    p = tf_rb.paragraphs[0]
    p.text = "Rubrik Evaluasi Risiko (L × D)"
    p.font.name = FONT_HEADING
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY_DARK

    p_f = tf_rb.add_paragraph()
    p_f.space_before = Pt(6)
    p_f.text = "Skor Risiko = Kemungkinan (L) × Dampak (D)"
    p_f.font.bold = True
    p_f.font.size = Pt(12)
    p_f.font.color.rgb = BLUE_PRIMARY

    rubrik_desc = [
        ("Kemungkinan (Likelihood - L):", "Skala Ordinal 1–3\n• 1 (Rendah): Jarang terjadi atau sulit diwujudkan.\n• 2 (Sedang): Masuk akal dalam penggunaan reguler.\n• 3 (Tinggi): Relatif mudah dieksekusi tanpa kendala teknis."),
        ("Dampak (Impact - D):", "Skala Ordinal 1–3\n• 1 (Rendah): Dampak terbatas, mudah dipulihkan.\n• 2 (Sedang): Mengganggu sebagian pengguna/layanan.\n• 3 (Tinggi): Kebocoran besar, data rusak, layanan lumpuh."),
        ("Klasifikasi Tingkat Keparahan:", "• Skor 1–2: Risiko Rendah (Low)\n• Skor 3–4: Risiko Sedang (Medium)\n• Skor 6–9: Risiko Tinggi (High)")
    ]
    for r_title, r_desc in rubrik_desc:
        p = tf_rb.add_paragraph()
        p.space_before = Pt(8)
        p.text = r_title
        p.font.bold = True
        p.font.size = Pt(10.5)
        p.font.color.rgb = NAVY_DARK

        p2 = tf_rb.add_paragraph()
        p2.text = r_desc
        p2.font.size = Pt(9)
        p2.font.color.rgb = TEXT_MAIN

    # ==========================================
    # SLIDE 7: REGISTER ANCAMAN: SPOOFING & TAMPERING
    # ==========================================
    s7 = prs.slides.add_slide(blank_layout)
    create_slide_header(s7, "Register Ancaman: Spoofing & Tampering (T01–T07)", "ANALISIS: TIO ARMANI (103112430225)")

    # Table of T01 - T07
    rows = 8
    cols = 5
    table_shape = s7.shapes.add_table(rows, cols, Inches(0.8), Inches(1.5), Inches(11.7), Inches(5.4))
    table = table_shape.table
    table.columns[0].width = Inches(0.8)  # ID
    table.columns[1].width = Inches(3.2)  # Judul Ancaman
    table.columns[2].width = Inches(1.3)  # STRIDE & El
    table.columns[3].width = Inches(1.5)  # Skor Risiko
    table.columns[4].width = Inches(4.9)  # Rekomendasi Mitigasi Teknis

    headers = ["ID", "Judul Ancaman Spesifik", "Tipe / Elemen", "Skor Risiko", "Rekomendasi Mitigasi Teknis"]
    for c_i, h in enumerate(headers):
        cell = table.cell(0, c_i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = SLATE_DARK
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.bold = True
        p.font.size = Pt(10)
        p.font.color.rgb = TEXT_WHITE

    st_data = [
        ("T01", "Titip absen dengan akun mahasiswa lain", "Spoofing\n[E1. Mahasiswa]", "L:3 × D:2 = 6\n[Tinggi / High]", "MFA pada SSO kampus; pengikatan sesi ke fingerprint perangkat; deteksi multi-presensi."),
        ("T02", "Token autentikasi mahasiswa dicuri & dipakai ulang", "Spoofing\n[F3. Mhs ke P2]", "L:2 × D:3 = 6\n[Tinggi / High]", "Token ber-TTL singkat ber-audience presensi; cookie HttpOnly, Secure, SameSite; HSTS."),
        ("T03", "QR dibagikan jarak jauh ke mahasiswa tidak hadir", "Spoofing\n[P2. Validasi]", "L:3 × D:2 = 6\n[Tinggi / High]", "Dynamic rolling QR code berganti tiap 15–30 detik (TOTP); batasan Wi-Fi kelas."),
        ("T04", "Token QR dapat ditebak atau diprediksi", "Spoofing\n[P1. Kelola Sesi]", "L:2 × D:2 = 4\n[Sedang / Med]", "Pembangkitan token CSPRNG minimal 128-bit; token terikat kriptografis dengan ID sesi."),
        ("T05", "Manipulasi ID sesi atau ID mahasiswa pada presensi", "Tampering\n[F3. Mhs ke P2]", "L:2 × D:3 = 6\n[Tinggi / High]", "Identitas hanya diekstrak dari klaim token terverifikasi di server; validasi relasi kelas."),
        ("T06", "Presensi ganda atau race condition waktu klien", "Tampering\n[P2. Validasi]", "L:2 × D:2 = 4\n[Sedang / Med]", "Unique constraint (id_sesi, id_mhs) database; transaksi atomik; jam server NTP."),
        ("T07", "SQL injection melalui parameter presensi ke database", "Tampering\n[F11. P2 ke DB]", "L:2 × D:3 = 6\n[Tinggi / High]", "Prepared statements (parameterized query); validasi tipe input; akun DB least privilege.")
    ]

    for r_i, r_data in enumerate(st_data):
        for c_i, val in enumerate(r_data):
            cell = table.cell(r_i + 1, c_i)
            cell.fill.solid()
            cell.fill.fore_color.rgb = CARD_BG if r_i % 2 == 0 else BG_LIGHT
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.size = Pt(9)
            if c_i == 0:
                p.font.bold = True
                p.font.color.rgb = BLUE_PRIMARY
            elif c_i == 3:
                p.font.bold = True
                p.font.color.rgb = RED_SEV if "Tinggi" in val else AMBER_SEV
            else:
                p.font.color.rgb = TEXT_MAIN

    # ==========================================
    # SLIDE 8: REGISTER ANCAMAN: R, I, D, E
    # ==========================================
    s8 = prs.slides.add_slide(blank_layout)
    create_slide_header(s8, "Register Ancaman: R, I, D, E (T08–T14)", "ANALISIS: CHADAFYA PUTRA ZULFIKAR (103112430173)")

    table_shape8 = s8.shapes.add_table(rows, cols, Inches(0.8), Inches(1.5), Inches(11.7), Inches(5.4))
    table8 = table_shape8.table
    table8.columns[0].width = Inches(0.8)
    table8.columns[1].width = Inches(3.2)
    table8.columns[2].width = Inches(1.3)
    table8.columns[3].width = Inches(1.5)
    table8.columns[4].width = Inches(4.9)

    for c_i, h in enumerate(headers):
        cell = table8.cell(0, c_i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = SLATE_DARK
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.bold = True
        p.font.size = Pt(10)
        p.font.color.rgb = TEXT_WHITE

    ride_data = [
        ("T08", "Pengubahan / penyangkalan catatan hadir tanpa audit log", "Repudiation\n[P2. Validasi]", "L:2 × D:2 = 4\n[Sedang / Med]", "Centralized append-only audit log mencatat aktor, ID sesi, timestamp server, IP/perangkat."),
        ("T09", "Rekap memuat data pribadi berlebih ke pihak tidak berhak", "Info Disclosure\n[F16. Respons]", "L:2 × D:2 = 4\n[Sedang / Med]", "Minimalisasi data (data minimization DTO); pemisahan DTO dosen & admin; verifikasi wewenang."),
        ("T10", "Akses langsung atau kebocoran cadangan basis data", "Info Disclosure\n[DS1. Database]", "L:1 × D:3 = 3\n[Sedang / Med]", "Enkripsi data at-rest (TDE); isolasi database via bastion host ber-MFA; enkripsi file backup."),
        ("T11", "Lonjakan presensi serentak 5 menit melumpuhkan backend", "Denial of Service\n[P2. Validasi]", "L:3 × D:2 = 6\n[Tinggi / High]", "Message queue asynchronous; in-memory cache (Redis) untuk validasi sesi; rate limiting."),
        ("T12", "Layanan Autentikasi Kampus (SSO) lumpuh", "Denial of Service\n[E4. Layanan SSO]", "L:2 × D:3 = 6\n[Tinggi / High]", "Verifikasi token lokal berbasis kunci publik (cache JWKS); circuit breaker; SOP darurat."),
        ("T13", "Dosen non-pengampu atau pihak lain membuka sesi kelas", "Elevation of Priv\n[P1. Kelola Sesi]", "L:2 × D:3 = 6\n[Tinggi / High]", "Pemeriksaan otorisasi berbasis objek (ABAC): cocokkan NIP pengampu di DB; deny-by-default."),
        ("T14", "Akses data rekapitulasi lintas kelas oleh peran tidak berhak", "Elevation of Priv\n[P3. Muat Rekap]", "L:2 × D:3 = 6\n[Tinggi / High]", "Pemisahan endpoint dosen vs admin; filter kueri scope kelas dosen pengampu; testing CI/CD.")
    ]

    for r_i, r_data in enumerate(ride_data):
        for c_i, val in enumerate(r_data):
            cell = table8.cell(r_i + 1, c_i)
            cell.fill.solid()
            cell.fill.fore_color.rgb = CARD_BG if r_i % 2 == 0 else BG_LIGHT
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.size = Pt(9)
            if c_i == 0:
                p.font.bold = True
                p.font.color.rgb = BLUE_PRIMARY
            elif c_i == 3:
                p.font.bold = True
                p.font.color.rgb = RED_SEV if "Tinggi" in val else AMBER_SEV
            else:
                p.font.color.rgb = TEXT_MAIN

    # ==========================================
    # SLIDE 9: MATRIKS RISIKO & PRIORITAS MITIGASI
    # ==========================================
    s9 = prs.slides.add_slide(blank_layout)
    create_slide_header(s9, "Matriks Evaluasi Risiko & Prioritas Remediasi", "EVALUASI RISIKO: TEDY PERMANA PUTRA (103112430157)")

    # Kiri: Matriks 3x3
    add_card(s9, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.4))
    tb_m = s9.shapes.add_textbox(Inches(1.0), Inches(1.7), Inches(5.2), Inches(5.0))
    tf_m = tb_m.text_frame
    tf_m.word_wrap = True

    p = tf_m.paragraphs[0]
    p.text = "Pemetaan Matriks Risiko 3 × 3"
    p.font.name = FONT_HEADING
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY_DARK

    matrix_rows = [
        ("Kemungkinan 3 (Tinggi):", "• Dampak 2 (Sedang): T01, T03, T11 (Skor 6 — TINGGI)\n• Dampak 3 (Tinggi): —"),
        ("Kemungkinan 2 (Sedang):", "• Dampak 2 (Sedang): T04, T06, T08, T09 (Skor 4 — SEDANG)\n• Dampak 3 (Tinggi): T02, T05, T07, T12, T13, T14 (Skor 6 — TINGGI)"),
        ("Kemungkinan 1 (Rendah):", "• Dampak 3 (Tinggi): T10 (Skor 3 — SEDANG)"),
        ("Rekapitulasi:", "Total 14 Ancaman: 9 Risiko Tinggi (64.3%) & 5 Risiko Sedang (35.7%).")
    ]
    for m_title, m_desc in matrix_rows:
        p = tf_m.add_paragraph()
        p.space_before = Pt(8)
        p.text = m_title
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = BLUE_PRIMARY

        p2 = tf_m.add_paragraph()
        p2.text = m_desc
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = TEXT_MAIN

    # Kanan: 4 Klaster Prioritas Tindakan
    add_card(s9, Inches(6.8), Inches(1.5), Inches(5.7), Inches(5.4))
    tb_p = s9.shapes.add_textbox(Inches(7.0), Inches(1.7), Inches(5.3), Inches(5.0))
    tf_p = tb_p.text_frame
    tf_p.word_wrap = True

    p = tf_p.paragraphs[0]
    p.text = "4 Klaster Urutan Prioritas Mitigasi"
    p.font.name = FONT_HEADING
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY_DARK

    priors = [
        ("Prioritas 1 — Integritas Otorisasi & Validasi Server", RED_SEV,
         "T13, T14, T05, T07 (Skor 6 — Tinggi)\nFokus pada kontrol akses objek dosen pengampu, validasi parameter NIM server-side, dan prepared statements SQL. Dampak integritas masif dengan kompleksitas implementasi terukur."),
        ("Prioritas 2 — Keaslian Presensi Fisik Mahasiswa", AMBER_SEV,
         "T01, T03, T02 (Skor 6 — Tinggi)\nRisiko paling nyata di perkuliahan. Solusi: rolling QR dinamis (TOTP 15–30s), pembatasan subnet Wi-Fi kelas, dan pengikatan sesi ke perangkat."),
        ("Prioritas 3 — Ketahanan Ketersediaan Layanan", BLUE_PRIMARY,
         "T11, T12 (Skor 6 — Tinggi)\nMenjamin ketersediaan pada lonjakan 5 menit melalui message queue asinkron, Redis caching, dan verifikasi lokal JWKS SSO."),
        ("Prioritas 4 — Pertahanan Pendukung & Tata Kelola", GREEN_SEV,
         "T08, T04, T06, T09, T10 (Skor 3–4 — Sedang)\nAudit log terpusat, penguatan CSPRNG, data minimization DTO, dan enkripsi cadangan basis data.")
    ]
    for p_title, p_color, p_desc in priors:
        p = tf_p.add_paragraph()
        p.space_before = Pt(8)
        p.text = p_title
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = p_color

        p2 = tf_p.add_paragraph()
        p2.text = p_desc
        p2.font.size = Pt(9)
        p2.font.color.rgb = TEXT_MAIN

    # ==========================================
    # SLIDE 10: IMPLEMENTASI DI OWASP THREAT DRAGON
    # ==========================================
    s10 = prs.slides.add_slide(blank_layout)
    create_slide_header(s10, "Implementasi Model di OWASP Threat Dragon v2")

    # Kiri: Screenshot Ekspor / Model Report
    add_card(s10, Inches(0.8), Inches(1.5), Inches(6.0), Inches(5.4))
    if os.path.exists("cap_01_summary.png"):
        s10.shapes.add_picture("cap_01_summary.png", Inches(0.95), Inches(1.65), width=Inches(5.7))

    # Kanan: Poin Verifikasi Kelayakan
    add_card(s10, Inches(7.1), Inches(1.5), Inches(5.4), Inches(5.4))
    tb_v = s10.shapes.add_textbox(Inches(7.3), Inches(1.7), Inches(5.0), Inches(5.0))
    tf_v = tb_v.text_frame
    tf_v.word_wrap = True

    p = tf_v.paragraphs[0]
    p.text = "Bukti Autentik Praktikum"
    p.font.name = FONT_HEADING
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY_DARK

    v_points = [
        ("Validasi Model Utuh:", "Model KasusA_KelompokXX_FINAL.json memuat 28 sel diagram: 4 aktor, 3 proses, 1 database, 16 data flow, dan 4 trust boundary."),
        ("14 Ancaman Terpetakan:", "Seluruh 14 ancaman terhubung ke elemen spesifik dengan metadata Title, STRIDE Type, Severity, Skenario, dan Mitigasi."),
        ("Status Seluruh Ancaman 'Open':", "Status tetap Open karena mitigasi yang dirumuskan merupakan rekomendasi desain arsitektur sebelum diverifikasi dalam implementasi kode."),
        ("Ekspor Laporan PDF Resmi:", "Fitur Report Threat Dragon berhasil mencetak laporan 9 halaman lengkap mencakup diagram hingga element properties."),
        ("Uji Buka Ulang (Re-open Test):", "Model berhasil ditutup dan dibuka kembali tanpa korupsi sintaks JSON.")
    ]
    for vp_title, vp_desc in v_points:
        p = tf_v.add_paragraph()
        p.space_before = Pt(6)
        p.text = vp_title
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = BLUE_PRIMARY

        p2 = tf_v.add_paragraph()
        p2.text = vp_desc
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = TEXT_MAIN

    # ==========================================
    # SLIDE 11: KESIMPULAN & REKOMENDASI ARSITEKTURAL
    # ==========================================
    s11 = prs.slides.add_slide(blank_layout)
    create_slide_header(s11, "Kesimpulan & Rekomendasi Keamanan Arsitektural")

    # 3 Big Recommendation Cards
    rec_cards = [
        ("1. Filosofi 'Shift-Left Security'", BLUE_PRIMARY,
         "Threat modeling membuktikan bahwa kerentanan keamanan paling efektif dan efisien diatasi pada tahap perancangan (design level) sebelum kode ditulis, mencegah biaya penambalan mahal di fase produksi."),
        ("2. Dua Domain Risiko Kritis", RED_SEV,
         "Titik paling rawan terkonsentrasi pada:\n(a) Keaslian presensi fisik (titip absen & penyebaran foto QR)\n(b) Penegakan otorisasi di sisi server (ABAC & verifikasi token).\nKeduanya wajib menjadi prioritas rekayasa perangkat lunak."),
        ("3. Pertahanan Berlapis (Defense-in-Depth)", GREEN_SEV,
         "Sistem presensi tidak boleh hanya mengandalkan kode QR statis. Pengamanan harus menggabungkan rolling QR dinamis (TOTP), pengikatan sesi perangkat, audit log tamper-evident, serta arsitektur antrean asinkron.")
    ]

    r_w = Inches(3.7)
    r_h = Inches(5.3)
    r_gap = Inches(0.3)
    r_start_x = Inches(0.8)

    for i, (title, color, text) in enumerate(rec_cards):
        rx = r_start_x + i * (r_w + r_gap)
        add_card(s11, rx, Inches(1.5), r_w, r_h, bg_color=CARD_BG, border_color=color, border_width=1.5)

        tb_r = s11.shapes.add_textbox(rx + Inches(0.2), Inches(1.75), r_w - Inches(0.4), r_h - Inches(0.5))
        tf_r = tb_r.text_frame
        tf_r.word_wrap = True

        p = tf_r.paragraphs[0]
        p.text = title
        p.font.name = FONT_HEADING
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = color

        p2 = tf_r.add_paragraph()
        p2.space_before = Pt(10)
        p2.text = text
        p2.font.name = FONT_BODY
        p2.font.size = Pt(10.5)
        p2.font.color.rgb = TEXT_MAIN

    # ==========================================
    # SLIDE 12: SESI TANYA JAWAB (Q&A) & PENUTUP
    # ==========================================
    s12 = prs.slides.add_slide(blank_layout)
    bg12 = s12.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg12.fill.solid()
    bg12.fill.fore_color.rgb = NAVY_DARK
    bg12.line.fill.background()

    # Center box
    tb_end = s12.shapes.add_textbox(Inches(1.5), Inches(1.8), Inches(10.3), Inches(4.0))
    tf_e = tb_end.text_frame
    tf_e.word_wrap = True

    p = tf_e.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = "TERIMA KASIH"
    p.font.name = FONT_HEADING
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = TEXT_WHITE

    p2 = tf_e.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(8)
    p2.text = "Sesi Tanya Jawab (Q & A) — Threat Modeling Kasus A"
    p2.font.name = FONT_BODY
    p2.font.size = Pt(18)
    p2.font.color.rgb = CYAN_ACCENT

    p3 = tf_e.add_paragraph()
    p3.alignment = PP_ALIGN.CENTER
    p3.space_before = Pt(24)
    p3.text = "Repositori GitHub Proyek:\nhttps://github.com/FatirGibran/Tugas-2-SIber"
    p3.font.name = FONT_BODY
    p3.font.size = Pt(14)
    p3.font.color.rgb = RGBColor(203, 213, 225)

    p4 = tf_e.add_paragraph()
    p4.alignment = PP_ALIGN.CENTER
    p4.space_before = Pt(16)
    p4.text = "Fatir Gibran • Tio Armani • Chadafya Putra Zulfikar • Tedy Permana Putra\nTeknik Informatika / Sistem Informasi — Mata Kuliah Keamanan Siber"
    p4.font.name = FONT_BODY
    p4.font.size = Pt(11)
    p4.font.color.rgb = RGBColor(148, 163, 184)

    output_pptx = "Presentasi_Threat_Modeling_Kasus_A_Presensi_QR.pptx"
    prs.save(output_pptx)
    print(f"Slide presentasi berhasil disimpan di {output_pptx}")

if __name__ == "__main__":
    build_presentation()
