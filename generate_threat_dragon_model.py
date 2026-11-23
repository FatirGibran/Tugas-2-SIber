# -*- coding: utf-8 -*-
"""
Script to generate valid OWASP Threat Dragon v2.x JSON model for Kasus A.
Model contains 8 elements, 16 data flows, 4 trust boundaries, and all 14 threats T01-T14.
"""

import json
import uuid

def make_uuid():
    return str(uuid.uuid4())

def build_threat_dragon_model():
    # Pre-generate IDs
    id_tb1 = "606bdcee-943e-4686-8c43-c81099dcda01"
    id_tb_auth = "606bdcee-943e-4686-8c43-c81099dcda02"
    id_tb2 = "606bdcee-943e-4686-8c43-c81099dcda03"
    id_tb3 = "606bdcee-943e-4686-8c43-c81099dcda04"

    id_e1 = "bb3f40db-ddf0-41cf-bc9a-3ac1924b9001" # Mahasiswa
    id_e2 = "bb3f40db-ddf0-41cf-bc9a-3ac1924b9002" # Dosen
    id_e3 = "bb3f40db-ddf0-41cf-bc9a-3ac1924b9003" # Admin
    id_e4 = "bb3f40db-ddf0-41cf-bc9a-3ac1924b9004" # Layanan SSO

    id_p1 = "369ba8b4-845b-4d9b-97fc-34b363ae9001" # P1. Kelola Sesi & QR
    id_p2 = "369ba8b4-845b-4d9b-97fc-34b363ae9002" # P2. Validasi & Catat Presensi
    id_p3 = "369ba8b4-845b-4d9b-97fc-34b363ae9003" # P3. Muat Rekap Kehadiran

    id_ds1 = "b063aea2-ab95-4c16-83d7-76b9869f9001" # DS1. Database Akademik

    # Ports helper
    def make_ports():
        return {
            "groups": {
                "top": {"position": "top", "attrs": {"circle": {"r": 4, "magnet": True, "stroke": "#5F95FF", "strokeWidth": 1, "fill": "#fff", "style": {"visibility": "hidden"}}}},
                "right": {"position": "right", "attrs": {"circle": {"r": 4, "magnet": True, "stroke": "#5F95FF", "strokeWidth": 1, "fill": "#fff", "style": {"visibility": "hidden"}}}},
                "bottom": {"position": "bottom", "attrs": {"circle": {"r": 4, "magnet": True, "stroke": "#5F95FF", "strokeWidth": 1, "fill": "#fff", "style": {"visibility": "hidden"}}}},
                "left": {"position": "left", "attrs": {"circle": {"r": 4, "magnet": True, "stroke": "#5F95FF", "strokeWidth": 1, "fill": "#fff", "style": {"visibility": "hidden"}}}}
            },
            "items": [
                {"group": "top", "id": make_uuid()},
                {"group": "right", "id": make_uuid()},
                {"group": "bottom", "id": make_uuid()},
                {"group": "left", "id": make_uuid()}
            ]
        }

    # Threats data definitions
    threat_t01 = {
        "id": make_uuid(),
        "number": 1,
        "title": "Titip absen dengan akun mahasiswa lain",
        "status": "Open",
        "severity": "High",
        "type": "Spoofing",
        "modelType": "STRIDE",
        "new": False,
        "score": "6 (L=3, D=2)",
        "description": "Jika mahasiswa yang tidak hadir di kelas memberikan kredensial akun kampus miliknya secara sukarela kepada rekannya yang berada di ruang kelas, dan rekannya tersebut melakukan login serta memindai kode QR dari layar proyektor, maka sistem mencatat kehadiran mahasiswa yang bersangkutan secara sah padahal tidak hadir.",
        "mitigation": "MFA pada autentikasi kampus; pengikatan sesi ke fingerprint perangkat; deteksi anomali satu perangkat mengirim presensi banyak akun atau satu akun dari banyak perangkat; konfirmasi presensi real-time."
    }

    threat_t02 = {
        "id": make_uuid(),
        "number": 2,
        "title": "Token autentikasi mahasiswa dicuri dan dipakai ulang",
        "status": "Open",
        "severity": "High",
        "type": "Spoofing",
        "modelType": "STRIDE",
        "new": False,
        "score": "6 (L=2, D=3)",
        "description": "Jika penyerang memperoleh token autentikasi mahasiswa (misalnya dari perangkat bersama laboratorium atau sesi hijacking) dan mengirimkannya pada aliran F3, maka backend menerima permintaan presensi sebagai mahasiswa tersebut.",
        "mitigation": "Token berumur pendek (TTL singkat) dengan audiens khusus presensi; cookie HttpOnly, Secure, SameSite=Strict; mekanisme token revocation saat logout; penerapan HSTS."
    }

    threat_t03 = {
        "id": make_uuid(),
        "number": 3,
        "title": "QR dibagikan jarak jauh ke mahasiswa yang tidak hadir",
        "status": "Open",
        "severity": "High",
        "type": "Spoofing",
        "modelType": "STRIDE",
        "new": False,
        "score": "6 (L=3, D=2)",
        "description": "Jika mahasiswa di kelas memotret atau menyiarkan tampilan QR code kepada rekannya di luar kelas melalui pesan instan, maka rekannya dapat memindai foto tersebut dalam batas waktu 5 menit dan tercatat hadir.",
        "mitigation": "QR dinamis (rolling QR code) berganti setiap 15-30 detik berbasis algoritma TOTP; pembatasan subnet IP Wi-Fi kelas; deteksi anomali lokasi pemindaian."
    }

    threat_t04 = {
        "id": make_uuid(),
        "number": 4,
        "title": "Token QR dapat ditebak atau diprediksi",
        "status": "Open",
        "severity": "Medium",
        "type": "Spoofing",
        "modelType": "STRIDE",
        "new": False,
        "score": "4 (L=2, D=2)",
        "description": "Jika token sesi dibangkitkan dengan algoritma berentropi rendah atau berurutan, maka penyerang dapat menebak nilai token sesi aktif dan mengirimkan permintaan presensi tanpa perlu melihat QR code.",
        "mitigation": "Pembangkitan token acak kriptografis (CSPRNG minimal 128 bit); token terikat kriptografis dengan ID sesi; rate limiting dan pembatasan percobaan token salah."
    }

    threat_t05 = {
        "id": make_uuid(),
        "number": 5,
        "title": "Manipulasi ID sesi atau ID mahasiswa pada permintaan presensi",
        "status": "Open",
        "severity": "High",
        "type": "Tampering",
        "modelType": "STRIDE",
        "new": False,
        "score": "6 (L=2, D=3)",
        "description": "Jika mahasiswa merekayasa parameter ID mahasiswa atau ID sesi pada body/query permintaan presensi dan backend mempercayainya, maka kehadiran dapat dicatat atas nama orang lain atau untuk kelas yang tidak diikuti.",
        "mitigation": "Identitas hanya diekstraksi dari token autentikasi terverifikasi (bukan dari parameter request); verifikasi relasi keanggotaan kelas di sisi server; validasi skema parameter ketat."
    }

    threat_t06 = {
        "id": make_uuid(),
        "number": 6,
        "title": "Presensi ganda atau lewat masa berlaku akibat race condition dan manipulasi waktu",
        "status": "Open",
        "severity": "Medium",
        "type": "Tampering",
        "modelType": "STRIDE",
        "new": False,
        "score": "4 (L=2, D=2)",
        "description": "Jika mahasiswa mengirim banyak permintaan presensi secara paralel atau memanipulasi waktu perangkat klien, maka dua catatan hadir dapat terbentuk atau presensi diterima setelah batas 5 menit berakhir.",
        "mitigation": "Unique constraint (id_sesi, id_mahasiswa) pada database; transaksi atomik dengan database lock; pemeriksaan waktu dan masa aktif sesi hanya menggunakan jam server (NTP); idempotency key."
    }

    threat_t07 = {
        "id": make_uuid(),
        "number": 7,
        "title": "SQL injection melalui parameter presensi ke database",
        "status": "Open",
        "severity": "High",
        "type": "Tampering",
        "modelType": "STRIDE",
        "new": False,
        "score": "6 (L=2, D=3)",
        "description": "Jika input dari aliran F3 diteruskan ke kueri basis data tanpa parameterisasi, maka penyerang dapat menyuntikkan instruksi SQL berbahaya untuk memanipulasi atau membaca basis data akademik.",
        "mitigation": "Prepared statements / parameterized queries atau ORM teruji; validasi tipe input; akun basis data berprinsip least privilege; pemantauan eksekusi kueri anomali."
    }

    threat_t08 = {
        "id": make_uuid(),
        "number": 8,
        "title": "Pengubahan atau penyangkalan catatan hadir tanpa jejak audit log",
        "status": "Open",
        "severity": "Medium",
        "type": "Repudiation",
        "modelType": "STRIDE",
        "new": False,
        "score": "4 (L=2, D=2)",
        "description": "Jika peristiwa presensi, pembukaan/penutupan sesi, atau perubahan manual tidak tercatat dalam audit log yang memadai, maka mahasiswa dapat menyangkal kehadirannya dan staf dapat mengubah data tanpa dapat dilacak.",
        "mitigation": "Audit log terpusat bersifat append-only mencatat identitas aktor, ID sesi, timestamp server, IP/perangkat, dan aksi; proteksi integritas log; audit berkala."
    }

    threat_t09 = {
        "id": make_uuid(),
        "number": 9,
        "title": "Rekap memuat data pribadi berlebih atau terkirim ke pihak tidak berhak",
        "status": "Open",
        "severity": "Medium",
        "type": "Information disclosure",
        "modelType": "STRIDE",
        "new": False,
        "score": "4 (L=2, D=2)",
        "description": "Jika payload respons rekapitulasi menyertakan atribut pribadi sensitif berlebih atau dikirim tanpa validasi kepemilikan kelas dosen, maka data pribadi mahasiswa dapat terekspos.",
        "mitigation": "Prinsip minimalisasi data (data minimization DTO); pemisahan DTO dosen dan admin; otorisasi ketat kepemilikan kelas di sisi server."
    }

    threat_t10 = {
        "id": make_uuid(),
        "number": 10,
        "title": "Akses langsung atau kebocoran cadangan basis data akademik",
        "status": "Open",
        "severity": "Medium",
        "type": "Information disclosure",
        "modelType": "STRIDE",
        "new": False,
        "score": "3 (L=1, D=3)",
        "description": "Jika file cadangan (backup) basis data akademik disimpan tanpa enkripsi atau kontrol akses segmen internal miskonfigurasi, pihak tidak berwenang dapat membaca riwayat presensi dan data mahasiswa.",
        "mitigation": "Enkripsi data at rest (TDE); isolasi akses database via bastion host ber-MFA; enkripsi penuh file cadangan (backup); audit hak akses berkala."
    }

    threat_t11 = {
        "id": make_uuid(),
        "number": 11,
        "title": "Lonjakan permintaan presensi serentak melumpuhkan backend",
        "status": "Open",
        "severity": "High",
        "type": "Denial of service",
        "modelType": "STRIDE",
        "new": False,
        "score": "6 (L=3, D=2)",
        "description": "Karena sesi presensi hanya berdurasi 5 menit, konsentrasi pengiriman presensi serentak oleh ratusan mahasiswa dapat menyebabkan lonjakan beban yang melumpuhkan ketersediaan backend.",
        "mitigation": "Arsitektur antrean pesan asynchronous; cache in-memory (Redis) untuk validasi status sesi; rate limiting per IP/subnet; SOP perpanjangan darurat oleh dosen."
    }

    threat_t12 = {
        "id": make_uuid(),
        "number": 12,
        "title": "Layanan Autentikasi Kampus tidak tersedia menghentikan presensi",
        "status": "Open",
        "severity": "High",
        "type": "Denial of service",
        "modelType": "STRIDE",
        "new": False,
        "score": "6 (L=2, D=3)",
        "description": "Jika Layanan SSO Kampus mengalami downtime atau kelambatan, mahasiswa tidak dapat login dan backend tidak dapat memvalidasi token, melumpuhkan seluruh proses presensi.",
        "mitigation": "Validasi token secara lokal berbasis kunci publik (cache JWKS); circuit breaker pattern; redundansi SSO; SOP presensi manual darurat."
    }

    threat_t13 = {
        "id": make_uuid(),
        "number": 13,
        "title": "Dosen non-pengampu atau mahasiswa membuka/menutup sesi kelas lain",
        "status": "Open",
        "severity": "High",
        "type": "Elevation of privilege",
        "modelType": "STRIDE",
        "new": False,
        "score": "6 (L=2, D=3)",
        "description": "Jika proses P1 hanya memeriksa klaim 'dosen' tanpa mencocokkan relasi NIP pengampu sah untuk kelas terkait, pengguna lain dapat membuka atau menutup sesi kelas yang bukan miliknya.",
        "mitigation": "Pemeriksaan otorisasi berbasis objek (ABAC) di server: verifikasi relasi pengampu terhadap jadwal di database pada tiap request; tolak secara default (deny-by-default)."
    }

    threat_t14 = {
        "id": make_uuid(),
        "number": 14,
        "title": "Akses rekapitulasi kehadiran lintas kelas oleh peran tidak berhak",
        "status": "Open",
        "severity": "High",
        "type": "Elevation of privilege",
        "modelType": "STRIDE",
        "new": False,
        "score": "6 (L=2, D=3)",
        "description": "Jika proses P3 tidak membatasi hasil rekapitulasi berdasarkan kelas binaan dosen atau tidak memvalidasi endpoint admin, dosen dapat melihat data kelas lain atau mahasiswa mengakses data rekapitulasi kampus.",
        "mitigation": "Pembatasan hak akses berbasis peran dan kepemilikan kelas di server; pemisahan endpoint admin; pengujian otorisasi otomatis (authorization testing)."
    }

    # Flow generator helper
    flow_counter = 0
    def make_flow(src_id, tgt_id, name, threats=None, crossings=None):
        nonlocal flow_counter
        flow_counter += 1
        fid = make_uuid()
        th_list = threats or []
        has_open = len(th_list) > 0
        return {
            "shape": "flow",
            "attrs": {
                "line": {
                    "stroke": "red" if has_open else "#333333",
                    "strokeWidth": 2.5 if has_open else 1.5,
                    "targetMarker": {"name": "block"},
                    "sourceMarker": {"name": ""},
                    "strokeDasharray": None
                }
            },
            "width": 200,
            "height": 100,
            "zIndex": 10 + flow_counter,
            "connector": "smooth",
            "data": {
                "type": "tm.Flow",
                "name": name,
                "description": "",
                "outOfScope": False,
                "isTrustBoundary": False,
                "reasonOutOfScope": "",
                "hasOpenThreats": has_open,
                "isBidirectional": False,
                "isEncrypted": True,
                "isPublicNetwork": False,
                "protocol": "HTTPS",
                "threats": th_list,
                "trustBoundaryIds": crossings or []
            },
            "labels": [name],
            "id": fid,
            "source": {"cell": src_id},
            "target": {"cell": tgt_id}
        }

    # Flows F1 - F16
    f1 = make_flow(id_e2, id_p1, "F1: Buka/Tutup Sesi (Token Dosen)", crossings=[id_tb1, id_tb2])
    f2 = make_flow(id_p1, id_e2, "F2: Token QR Sesi 5 Menit", crossings=[id_tb2, id_tb1])
    f3 = make_flow(id_e1, id_p2, "F3: QR Token + Token Mahasiswa", threats=[threat_t02, threat_t05], crossings=[id_tb1, id_tb2])
    f4 = make_flow(id_p2, id_e1, "F4: Status Presensi (Sukses/Gagal)", crossings=[id_tb2, id_tb1])

    f5 = make_flow(id_e1, id_e4, "F5: Kredensial SSO", crossings=[id_tb1, id_tb_auth])
    f6 = make_flow(id_e4, id_e1, "F6: Token Autentikasi SSO", crossings=[id_tb_auth, id_tb1])

    f7 = make_flow(id_p2, id_e4, "F7: Verifikasi Token SSO", crossings=[id_tb2, id_tb_auth])
    f8 = make_flow(id_e4, id_p2, "F8: Hasil Verifikasi & Peran", crossings=[id_tb_auth, id_tb2])

    f9 = make_flow(id_p1, id_ds1, "F9: Tulis Sesi & QR Token", crossings=[id_tb2, id_tb3])
    f10 = make_flow(id_ds1, id_p1, "F10: Jadwal & Dosen Pengampu", crossings=[id_tb3, id_tb2])

    f11 = make_flow(id_p2, id_ds1, "F11: Validasi Relasi & Catat Hadir", threats=[threat_t07], crossings=[id_tb2, id_tb3])
    f12 = make_flow(id_ds1, id_p2, "F12: Status Sesi & Hasil Simpan", crossings=[id_tb3, id_tb2])

    f13 = make_flow(id_e2, id_p3, "F13: Permintaan Rekap Presensi", crossings=[id_tb1, id_tb2])
    f14 = make_flow(id_p3, id_ds1, "F14: Query Rekapitulasi Hadir", crossings=[id_tb2, id_tb3])
    f15 = make_flow(id_ds1, id_p3, "F15: Data Rekap Kehadiran", crossings=[id_tb3, id_tb2])
    f16 = make_flow(id_p3, id_e2, "F16: Respons Rekapitulasi Presensi", threats=[threat_t09], crossings=[id_tb2, id_tb1])

    # Element Cells
    c_e1 = {
        "position": {"x": 80, "y": 240},
        "size": {"width": 140, "height": 70},
        "attrs": {
            "text": {"text": "E1. Mahasiswa\n(Browser)"},
            "body": {"stroke": "red", "strokeWidth": 2.5, "strokeDasharray": None}
        },
        "visible": True,
        "shape": "actor",
        "zIndex": 2,
        "ports": make_ports(),
        "id": id_e1,
        "data": {
            "type": "tm.Actor",
            "name": "E1. Mahasiswa (Browser)",
            "description": "Pengguna mahasiswa yang mengakses antarmuka presensi melalui browser smartphone/laptop.",
            "isTrustBoundary": False,
            "outOfScope": False,
            "reasonOutOfScope": "",
            "hasOpenThreats": True,
            "providesAuthentication": False,
            "threats": [threat_t01]
        }
    }

    c_e2 = {
        "position": {"x": 80, "y": 90},
        "size": {"width": 140, "height": 70},
        "attrs": {
            "text": {"text": "E2. Dosen Pengampu\n(Browser)"},
            "body": {"stroke": "#333333", "strokeWidth": 1.5, "strokeDasharray": None}
        },
        "visible": True,
        "shape": "actor",
        "zIndex": 2,
        "ports": make_ports(),
        "id": id_e2,
        "data": {
            "type": "tm.Actor",
            "name": "E2. Dosen Pengampu (Browser)",
            "description": "Dosen pengampu yang membuka sesi, menayangkan kode QR di proyektor, dan memantau kehadiran.",
            "isTrustBoundary": False,
            "outOfScope": False,
            "reasonOutOfScope": "",
            "hasOpenThreats": False,
            "providesAuthentication": False,
            "threats": []
        }
    }

    c_e3 = {
        "position": {"x": 80, "y": 390},
        "size": {"width": 140, "height": 70},
        "attrs": {
            "text": {"text": "E3. Admin Akademik\n(Browser)"},
            "body": {"stroke": "#333333", "strokeWidth": 1.5, "strokeDasharray": None}
        },
        "visible": True,
        "shape": "actor",
        "zIndex": 2,
        "ports": make_ports(),
        "id": id_e3,
        "data": {
            "type": "tm.Actor",
            "name": "E3. Admin Akademik (Browser)",
            "description": "Administrator akademik yang melihat rekapitulasi data kehadiran seluruh kelas.",
            "isTrustBoundary": False,
            "outOfScope": False,
            "reasonOutOfScope": "",
            "hasOpenThreats": False,
            "providesAuthentication": False,
            "threats": []
        }
    }

    c_e4 = {
        "position": {"x": 80, "y": 550},
        "size": {"width": 140, "height": 80},
        "attrs": {
            "text": {"text": "E4. Layanan Autentikasi\nKampus (SSO)"},
            "body": {"stroke": "red", "strokeWidth": 2.5, "strokeDasharray": None}
        },
        "visible": True,
        "shape": "actor",
        "zIndex": 2,
        "ports": make_ports(),
        "id": id_e4,
        "data": {
            "type": "tm.Actor",
            "name": "E4. Layanan Autentikasi Kampus (SSO)",
            "description": "Layanan identitas eksternal yang memverifikasi kredensial akun dan menerbitkan token SSO.",
            "isTrustBoundary": False,
            "outOfScope": False,
            "reasonOutOfScope": "",
            "hasOpenThreats": True,
            "providesAuthentication": True,
            "threats": [threat_t12]
        }
    }

    c_p1 = {
        "position": {"x": 420, "y": 90},
        "size": {"width": 160, "height": 85},
        "attrs": {
            "text": {"text": "P1. Kelola Sesi & QR\n[Pembangkitan Sesi 5m]"},
            "body": {"stroke": "red", "strokeWidth": 2.5, "strokeDasharray": None}
        },
        "visible": True,
        "shape": "process",
        "zIndex": 3,
        "ports": make_ports(),
        "id": id_p1,
        "data": {
            "type": "tm.Process",
            "name": "P1. Kelola Sesi & QR",
            "description": "Proses backend yang membangkitkan token QR 5 menit dan memvalidasi wewenang dosen pengampu.",
            "outOfScope": False,
            "isTrustBoundary": False,
            "reasonOutOfScope": "",
            "hasOpenThreats": True,
            "isWebApplication": True,
            "threats": [threat_t04, threat_t13]
        }
    }

    c_p2 = {
        "position": {"x": 420, "y": 240},
        "size": {"width": 160, "height": 95},
        "attrs": {
            "text": {"text": "P2. Validasi & Catat\nPresensi [Aturan 5m & Relasi]"},
            "body": {"stroke": "red", "strokeWidth": 2.5, "strokeDasharray": None}
        },
        "visible": True,
        "shape": "process",
        "zIndex": 3,
        "ports": make_ports(),
        "id": id_p2,
        "data": {
            "type": "tm.Process",
            "name": "P2. Validasi & Catat Presensi",
            "description": "Proses validasi token QR, masa aktif sesi, keanggotaan kelas, dan pencegahan presensi ganda.",
            "outOfScope": False,
            "isTrustBoundary": False,
            "reasonOutOfScope": "",
            "hasOpenThreats": True,
            "isWebApplication": True,
            "threats": [threat_t03, threat_t06, threat_t08, threat_t11]
        }
    }

    c_p3 = {
        "position": {"x": 420, "y": 400},
        "size": {"width": 160, "height": 85},
        "attrs": {
            "text": {"text": "P3. Muat Rekap\nKehadiran [Otorisasi Peran]"},
            "body": {"stroke": "red", "strokeWidth": 2.5, "strokeDasharray": None}
        },
        "visible": True,
        "shape": "process",
        "zIndex": 3,
        "ports": make_ports(),
        "id": id_p3,
        "data": {
            "type": "tm.Process",
            "name": "P3. Muat Rekap Kehadiran",
            "description": "Proses pengambilan dan penyajian rekapitulasi presensi sesuai hak akses peran dosen vs admin.",
            "outOfScope": False,
            "isTrustBoundary": False,
            "reasonOutOfScope": "",
            "hasOpenThreats": True,
            "isWebApplication": True,
            "threats": [threat_t14]
        }
    }

    c_ds1 = {
        "position": {"x": 780, "y": 220},
        "size": {"width": 160, "height": 130},
        "attrs": {
            "text": {"text": "DS1. Database\nAkademik\n(Akun, Jadwal, Sesi,\nKeanggotaan, Presensi)"},
            "topLine": {"stroke": "red", "strokeWidth": 2.5, "strokeDasharray": None},
            "bottomLine": {"stroke": "red", "strokeWidth": 2.5, "strokeDasharray": None}
        },
        "visible": True,
        "shape": "store",
        "zIndex": 3,
        "ports": make_ports(),
        "id": id_ds1,
        "data": {
            "type": "tm.Store",
            "name": "DS1. Database Akademik",
            "description": "Penyimpanan persisten seluruh entitas akademik kampus di jaringan internal.",
            "outOfScope": False,
            "isTrustBoundary": False,
            "reasonOutOfScope": "",
            "hasOpenThreats": True,
            "storesCredentials": True,
            "isEncrypted": True,
            "threats": [threat_t10]
        }
    }

    # Trust Boundaries
    c_tb1 = {
        "position": {"x": 50, "y": 60},
        "size": {"width": 200, "height": 430},
        "attrs": {"label": {"text": "TB1: Perangkat Pengguna / Internet"}},
        "visible": True,
        "shape": "trust-boundary-box",
        "zIndex": -1,
        "id": id_tb1,
        "data": {
            "type": "tm.BoundaryBox",
            "name": "TB1: Perangkat Pengguna / Internet",
            "description": "Zona perangkat klien publik yang tidak terpercaya.",
            "isTrustBoundary": True,
            "hasOpenThreats": False,
            "containedElements": [id_e1, id_e2, id_e3]
        }
    }

    c_tb_auth = {
        "position": {"x": 50, "y": 510},
        "size": {"width": 200, "height": 150},
        "attrs": {"label": {"text": "TB-Auth: Layanan Autentikasi Kampus"}},
        "visible": True,
        "shape": "trust-boundary-box",
        "zIndex": -1,
        "id": id_tb_auth,
        "data": {
            "type": "tm.BoundaryBox",
            "name": "TB-Auth: Layanan Autentikasi Kampus",
            "description": "Zona identitas eksternal kampus.",
            "isTrustBoundary": True,
            "hasOpenThreats": False,
            "containedElements": [id_e4]
        }
    }

    c_tb2 = {
        "position": {"x": 380, "y": 50},
        "size": {"width": 240, "height": 460},
        "attrs": {"label": {"text": "TB2: Backend Presensi (Aplikasi Internal)"}},
        "visible": True,
        "shape": "trust-boundary-box",
        "zIndex": -1,
        "id": id_tb2,
        "data": {
            "type": "tm.BoundaryBox",
            "name": "TB2: Backend Presensi (Aplikasi Internal)",
            "description": "Zona pemrosesan terpercaya yang mengevaluasi logika bisnis.",
            "isTrustBoundary": True,
            "hasOpenThreats": False,
            "containedElements": [id_p1, id_p2, id_p3]
        }
    }

    c_tb3 = {
        "position": {"x": 740, "y": 170},
        "size": {"width": 240, "height": 230},
        "attrs": {"label": {"text": "TB3: Zona Data Internal"}},
        "visible": True,
        "shape": "trust-boundary-box",
        "zIndex": -1,
        "id": id_tb3,
        "data": {
            "type": "tm.BoundaryBox",
            "name": "TB3: Zona Data Internal",
            "description": "Zona basis data terisolasi dan terlindungi.",
            "isTrustBoundary": True,
            "hasOpenThreats": False,
            "containedElements": [id_ds1]
        }
    }

    all_cells = [
        c_tb1, c_tb_auth, c_tb2, c_tb3,
        c_e1, c_e2, c_e3, c_e4,
        c_p1, c_p2, c_p3,
        c_ds1,
        f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
    ]

    model = {
        "version": "2.0.0",
        "release_version": "2.3.0",
        "released_at": "2026-11-23T22:00:00.000Z",
        "summary": {
            "title": "STRIDE - Kasus A - Presensi QR",
            "owner": "Fatir Gibran, Tio Armani, Chadafya Putra Zulfikar, Tedy Permana Putra",
            "description": "Threat modeling sistem presensi perkuliahan berbasis QR code menggunakan kerangka kerja STRIDE.",
            "id": 0
        },
        "detail": {
            "contributors": [
                {"name": "Fatir Gibran (103112430153)"},
                {"name": "Tio Armani (103112430225)"},
                {"name": "Chadafya Putra Zulfikar (103112430173)"},
                {"name": "Tedy Permana Putra (103112430157)"}
            ],
            "reviewer": "Dosen Praktikum Keamanan Siber (CAK3CAB3)",
            "threatTop": 14,
            "diagramTop": 1,
            "diagrams": [
                {
                    "id": 0,
                    "title": "DFD Sistem Presensi QR",
                    "diagramType": "STRIDE",
                    "placeholder": "DFD Level-1 Aliran Data Presensi Perkuliahan",
                    "thumbnail": "./public/content/images/thumbnail.stride.jpg",
                    "version": "2.0.0",
                    "description": "Diagram aliran data Level-1 untuk Sistem Presensi Perkuliahan Berbasis QR Code.",
                    "cells": all_cells
                }
            ]
        }
    }

    out_file = "KasusA_KelompokXX_FINAL.json"
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(model, f, indent=2, ensure_ascii=False)
    print(f"Model JSON berhasil dibuat dan disimpan di {out_file}")

if __name__ == "__main__":
    build_threat_dragon_model()
