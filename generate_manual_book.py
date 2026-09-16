"""
Generate Manual Book Word - Website Informasi Desa
Kelompok 1 KKN Global Institut
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from datetime import datetime

def set_para_spacing(para, before=0, after=6, line_spacing=1.5):
    pf = para.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line_spacing:
        pf.line_spacing = line_spacing

def add_heading(doc, text, level=1, color=None):
    h = doc.add_heading(level=level)
    h.clear()
    run = h.add_run(text)
    if color:
        run.font.color.rgb = RGBColor(*color)
    if level == 1:
        run.font.size = Pt(16)
        run.font.bold = True
        set_para_spacing(h, before=18, after=8)
    elif level == 2:
        run.font.size = Pt(13)
        run.font.bold = True
        set_para_spacing(h, before=12, after=6)
    elif level == 3:
        run.font.size = Pt(11.5)
        run.font.bold = True
        set_para_spacing(h, before=8, after=4)
    return h

def add_body(doc, text, indent=False, bold=False, size=12):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    if indent:
        p.paragraph_format.left_indent = Cm(0.8)
    set_para_spacing(p, before=0, after=8)
    return p

def add_bullet(doc, text, level=0, size=12):
    p = doc.add_paragraph(style='List Bullet')
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.clear()
    run = p.add_run(("    " * level) + "• " + text)
    run.font.size = Pt(size)
    set_para_spacing(p, before=0, after=6)
    return p

def add_numbered(doc, text, num, size=12):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(f"{num}. {text}")
    run.font.size = Pt(size)
    p.paragraph_format.left_indent = Cm(0.5)
    set_para_spacing(p, before=0, after=6)
    return p

def add_info_box(doc, label, value, size=12):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run_l = p.add_run(f"{label}: ")
    run_l.font.bold = True
    run_l.font.size = Pt(size)
    run_v = p.add_run(value)
    run_v.font.size = Pt(size)
    p.paragraph_format.left_indent = Cm(0.5)
    set_para_spacing(p, before=0, after=6)

def add_table_header(table, headers, bg_color="1F7A4D"):
    row = table.rows[0]
    for i, h in enumerate(headers):
        cell = row.cells[i]
        cell.text = h
        # Bold, white text
        for para in cell.paragraphs:
            for run in para.runs:
                run.font.bold = True
                run.font.size = Pt(9.5)
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # Green bg
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), bg_color)
        tcPr.append(shd)

def add_table_row(table, row_idx, values, alt=False):
    row = table.rows[row_idx]
    for i, val in enumerate(values):
        cell = row.cells[i]
        cell.text = str(val)
        for para in cell.paragraphs:
            for run in para.runs:
                run.font.size = Pt(9.5)
        if alt:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), 'F4F7F4')
            tcPr.append(shd)

def add_hr(doc):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'CCCCCC')
    pBdr.append(bottom)
    pPr.append(pBdr)
    set_para_spacing(p, before=4, after=4)

def create_manual():
    doc = Document()

    # --- Page setup ---
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(3.0)
    section.right_margin = Cm(2.5)

    # --- Default style ---
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(12)

    # ====================================================================
    # COVER PAGE
    # ====================================================================
    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph()

    p_cover = doc.add_paragraph()
    p_cover.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_cover.add_run("MANUAL BOOK")
    r.font.size = Pt(28)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x1F, 0x7A, 0x4D)
    set_para_spacing(p_cover, before=0, after=4)

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run("WEBSITE INFORMASI DESA")
    r2.font.size = Pt(18)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(0x1C, 0x2A, 0x22)
    set_para_spacing(p2, before=0, after=6)

    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r3 = p3.add_run("Panduan Lengkap Penggunaan & Pengelolaan\nSistem Informasi Desa Berbasis Web")
    r3.font.size = Pt(12)
    r3.font.color.rgb = RGBColor(0x6B, 0x74, 0x68)
    set_para_spacing(p3, before=0, after=40)

    add_hr(doc)

    doc.add_paragraph()

    p_kkn = doc.add_paragraph()
    p_kkn.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rk = p_kkn.add_run("Disusun oleh:\nKelompok 1 KKN Global Institut")
    rk.font.size = Pt(12)
    rk.font.bold = True
    rk.font.color.rgb = RGBColor(0x1F, 0x7A, 0x4D)
    set_para_spacing(p_kkn, before=0, after=6)

    p_year = doc.add_paragraph()
    p_year.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ry = p_year.add_run(f"Tahun {datetime.now().year}")
    ry.font.size = Pt(11)
    ry.font.color.rgb = RGBColor(0x6B, 0x74, 0x68)

    doc.add_page_break()

    # ====================================================================
    # KATA PENGANTAR
    # ====================================================================
    add_heading(doc, "KATA PENGANTAR", level=1, color=(0x1F, 0x7A, 0x4D))
    add_hr(doc)

    doc.add_paragraph()
    intro_text = (
        "Puji syukur kami panjatkan kepada Tuhan Yang Maha Esa atas selesainya penyusunan "
        "Manual Book Website Informasi Desa ini. Buku panduan ini disusun untuk memudahkan "
        "seluruh pihak yang terlibat — baik administrator desa maupun warga masyarakat — "
        "dalam menggunakan sistem informasi digital desa secara mandiri dan optimal.\n\n"
        "Website Informasi Desa merupakan sistem berbasis web yang dikembangkan oleh "
        "Kelompok 1 KKN Global Institut sebagai wujud kontribusi nyata dalam digitalisasi "
        "pelayanan publik di tingkat desa. Sistem ini dirancang untuk menjadi jembatan "
        "komunikasi antara pemerintah desa dan warga masyarakat secara transparan, "
        "mudah diakses, dan real-time.\n\n"
        "Buku panduan ini mencakup seluruh fitur dan alur penggunaan sistem, mulai dari "
        "tampilan publik bagi warga, panel administrasi bagi operator desa, hingga fitur-fitur "
        "pendukung seperti chatbot cerdas dan sistem pengaduan warga.\n\n"
        "Semoga buku panduan ini bermanfaat dan dapat menjadi acuan yang jelas bagi "
        "seluruh pengguna sistem."
    )
    p_intro = doc.add_paragraph()
    p_intro.add_run(intro_text).font.size = Pt(11)
    set_para_spacing(p_intro, before=0, after=6)

    doc.add_paragraph()
    p_ttd = doc.add_paragraph()
    p_ttd.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_ttd.add_run(f"Hormat kami,\n\n\n\nKelompok 1 KKN Global Institut").font.size = Pt(11)

    doc.add_page_break()

    # ====================================================================
    # DAFTAR ISI
    # ====================================================================
    add_heading(doc, "DAFTAR ISI", level=1, color=(0x1F, 0x7A, 0x4D))
    add_hr(doc)

    toc_items = [
        ("BAB I", "PENDAHULUAN", ""),
        ("1.1", "Latar Belakang", ""),
        ("1.2", "Tujuan Sistem", ""),
        ("1.3", "Ruang Lingkup", ""),
        ("1.4", "Spesifikasi Teknis", ""),
        ("BAB II", "ARSITEKTUR & INFRASTRUKTUR SISTEM", ""),
        ("2.1", "Arsitektur Teknologi", ""),
        ("2.2", "Basis Data (Database)", ""),
        ("2.3", "Penyimpanan File (Storage)", ""),
        ("2.4", "Deployment & Hosting", ""),
        ("BAB III", "PANDUAN PENGUNJUNG / WARGA", ""),
        ("3.1", "Mengakses Website", ""),
        ("3.2", "Halaman Beranda", ""),
        ("3.3", "Fitur Jam WIB & Status Kantor", ""),
        ("3.4", "Menu Profil Desa", ""),
        ("3.5", "Menu Perangkat Desa", ""),
        ("3.6", "Menu UMKM Desa", ""),
        ("3.7", "Menu Berita & Pengumuman", ""),
        ("3.8", "Menu Agenda & Kegiatan", ""),
        ("3.9", "Menu Dokumen & Surat Layanan", ""),
        ("3.10", "Menu Pengaduan & Aspirasi Warga", ""),
        ("3.11", "Menu Galeri Foto", ""),
        ("3.12", "Menu Kontak & Lokasi", ""),
        ("3.13", "Fitur Pencarian Global", ""),
        ("3.14", "Asisten Desa (Chatbot AI)", ""),
        ("3.15", "Navigasi Mobile", ""),
        ("BAB IV", "PANDUAN ADMINISTRATOR DESA", ""),
        ("4.1", "Akses Panel Admin", ""),
        ("4.2", "Login & Keamanan Sesi", ""),
        ("4.3", "Dashboard Admin", ""),
        ("4.4", "Mengelola Profil Desa", ""),
        ("4.5", "Mengelola Perangkat Desa", ""),
        ("4.6", "Mengelola UMKM Desa", ""),
        ("4.7", "Mengelola Agenda & Kegiatan", ""),
        ("4.8", "Mengelola Dokumen & Surat", ""),
        ("4.9", "Mengelola Pengaduan Warga", ""),
        ("4.10", "Mengelola Galeri Foto", ""),
        ("4.11", "Logout & Keamanan", ""),
        ("BAB V", "FITUR KEAMANAN SISTEM", ""),
        ("5.1", "Autentikasi Admin", ""),
        ("5.2", "Session Timeout 12 Jam", ""),
        ("5.3", "Rate Limiting Chatbot", ""),
        ("BAB VI", "TROUBLESHOOTING", ""),
        ("6.1", "Masalah Umum & Solusi", ""),
        ("6.2", "Kontak Bantuan Teknis", ""),
        ("BAB VII", "PENUTUP", ""),
    ]

    for num, title, page in toc_items:
        p_toc = doc.add_paragraph()
        if num.startswith("BAB"):
            r_num = p_toc.add_run(f"{num}  {title}")
            r_num.font.bold = True
            r_num.font.size = Pt(11)
            set_para_spacing(p_toc, before=6, after=2)
        else:
            p_toc.paragraph_format.left_indent = Cm(1.0)
            r_num = p_toc.add_run(f"{num}  {title}")
            r_num.font.size = Pt(10.5)
            set_para_spacing(p_toc, before=1, after=1)

    doc.add_page_break()

    # ====================================================================
    # BAB I: PENDAHULUAN
    # ====================================================================
    add_heading(doc, "BAB I  PENDAHULUAN", level=1, color=(0x1F, 0x7A, 0x4D))
    add_hr(doc)

    add_heading(doc, "1.1  Latar Belakang", level=2)
    add_body(doc,
        "Di era digitalisasi yang semakin berkembang, pemerintah desa dituntut untuk "
        "mampu menyediakan informasi kepada warga secara cepat, akurat, dan dapat diakses "
        "kapan saja maupun di mana saja. Website Informasi Desa hadir sebagai solusi "
        "konkret untuk memenuhi kebutuhan tersebut.\n\n"
        "Sistem ini dikembangkan oleh Kelompok 1 KKN Global Institut sebagai program "
        "pengabdian masyarakat yang bertujuan meningkatkan kapasitas digital pemerintahan "
        "desa. Dengan sistem ini, warga dapat mengakses berbagai informasi resmi desa, "
        "menyampaikan pengaduan, mengetahui jadwal kegiatan, dan berinteraksi dengan "
        "asisten digital yang siap membantu 24 jam penuh."
    )

    add_heading(doc, "1.2  Tujuan Sistem", level=2)
    bullets = [
        "Menyediakan informasi desa yang akurat, terkini, dan mudah diakses oleh seluruh warga.",
        "Meningkatkan transparansi dan keterbukaan informasi pemerintah desa.",
        "Memudahkan warga dalam mengurus keperluan administratif dan dokumen desa.",
        "Menyediakan kanal pengaduan dan aspirasi warga yang responsif.",
        "Mempromosikan potensi UMKM lokal kepada masyarakat luas.",
        "Mendokumentasikan kegiatan dan agenda desa secara digital.",
        "Menyediakan asisten digital berbasis AI untuk menjawab pertanyaan warga secara instan."
    ]
    for b in bullets:
        add_bullet(doc, b)

    add_heading(doc, "1.3  Ruang Lingkup", level=2)
    add_body(doc, "Sistem Website Informasi Desa mencakup dua modul utama:")
    add_bullet(doc, "Modul Publik (Warga): Tampilan yang dapat diakses oleh siapa pun tanpa perlu login, berisi 9 menu informasi desa.")
    add_bullet(doc, "Modul Admin (Operator Desa): Panel khusus yang diakses oleh administrator desa dengan autentikasi email dan password, berisi tools manajemen konten.")

    add_heading(doc, "1.4  Spesifikasi Teknis", level=2)
    specs = [
        ("Platform", "Website berbasis HTML5 / JavaScript (Single Page Application)"),
        ("Backend & Database", "Supabase (PostgreSQL) — layanan Backend-as-a-Service"),
        ("Penyimpanan File", "Supabase Storage (foto, logo, dokumen)"),
        ("AI Chatbot", "Groq API dengan model LLaMA 4 Scout via Supabase Edge Function"),
        ("Hosting", "Vercel (deployment otomatis dari GitHub)"),
        ("Versi Control", "GitHub — repository: github.com/indra1nkuss/webdesa"),
        ("Animasi", "GSAP (GreenSock Animation Platform) — bundled lokal"),
        ("Font", "Fraunces (display/judul) + Poppins (teks UI)"),
        ("Browser Support", "Chrome, Firefox, Safari, Edge — versi modern (2022+)"),
        ("Mobile Support", "Responsif penuh — iOS Safari, Android Chrome"),
    ]
    for label, val in specs:
        add_info_box(doc, label, val)

    doc.add_page_break()

    # ====================================================================
    # BAB II: ARSITEKTUR
    # ====================================================================
    add_heading(doc, "BAB II  ARSITEKTUR & INFRASTRUKTUR SISTEM", level=1, color=(0x1F, 0x7A, 0x4D))
    add_hr(doc)

    add_heading(doc, "2.1  Arsitektur Teknologi", level=2)
    add_body(doc, "Sistem menggunakan arsitektur JAMstack (JavaScript, API, Markup) dengan komponen-komponen berikut:")

    arch_items = [
        ("Frontend (Client-Side)", [
            "index.html — Struktur halaman publik (warga)",
            "admin.html — Panel administrasi desa",
            "assets/js/public.js — Logika tampilan warga (fetch data, render UI)",
            "assets/js/admin.js — Logika panel admin (CRUD, upload foto)",
            "assets/js/chatbot.js — Widget Asisten Desa (Tanya Desa)",
            "assets/js/config.js — Konfigurasi koneksi Supabase (URL & Key)",
            "assets/css/style.css — Seluruh gaya tampilan website",
        ]),
        ("Backend (Supabase)", [
            "Database PostgreSQL — Tabel data (site_config, perangkat_desa, umkm, agenda, dokumen, pengaduan, galeri, berita)",
            "Supabase Auth — Autentikasi email/password untuk admin",
            "Supabase Storage — Penyimpanan foto dan file dokumen",
            "Edge Function (chatbot) — Proxy ke Groq AI API",
        ]),
        ("Deployment", [
            "GitHub — Repositori source code",
            "Vercel — Auto-deploy setiap push ke branch main",
        ]),
    ]

    for section_title, items in arch_items:
        p = doc.add_paragraph()
        r = p.add_run(section_title)
        r.font.bold = True
        r.font.size = Pt(11)
        set_para_spacing(p, before=6, after=2)
        for item in items:
            add_bullet(doc, item)

    add_heading(doc, "2.2  Basis Data (Database)", level=2)
    add_body(doc, "Tabel-tabel database yang digunakan dalam sistem:")

    db_tables = [
        ("site_config", "1", "Profil desa, kontak, visi-misi, warna tema, media sosial"),
        ("perangkat_desa", "Banyak", "Data aparatur/pengurus desa"),
        ("umkm", "Banyak", "Data usaha mikro kecil menengah warga desa"),
        ("berita", "Banyak", "Berita dan pengumuman resmi desa"),
        ("agenda", "Banyak", "Jadwal kegiatan dan acara desa"),
        ("dokumen", "Banyak", "Formulir, Perdes, dan publikasi resmi"),
        ("pengaduan", "Banyak", "Laporan, aspirasi, dan pertanyaan warga"),
        ("galeri", "Banyak", "Foto dokumentasi kegiatan desa"),
    ]

    tbl = doc.add_table(rows=len(db_tables)+1, cols=3)
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    add_table_header(tbl, ["Nama Tabel", "Baris Data", "Keterangan"])
    for i, (name, rows, desc) in enumerate(db_tables):
        add_table_row(tbl, i+1, [name, rows, desc], alt=(i%2==1))

    doc.add_paragraph()
    add_heading(doc, "2.3  Penyimpanan File (Storage)", level=2)
    add_body(doc, "Semua file media (foto, logo, dokumen) disimpan di Supabase Storage dalam bucket bernama 'desa-media'. Format penamaan file: [prefix]_[timestamp].[ekstensi]. Contoh: perangkat_1726000000.jpg")

    add_heading(doc, "2.4  Deployment & Hosting", level=2)
    deploy_steps = [
        "Source code disimpan di GitHub (github.com/indra1nkuss/webdesa)",
        "Setiap perubahan yang di-push ke branch 'main' secara otomatis ter-deploy ke Vercel",
        "Vercel mendistribusikan file statik melalui CDN global",
        "Domain website dapat dikonfigurasi melalui dashboard Vercel",
    ]
    for i, s in enumerate(deploy_steps, 1):
        add_numbered(doc, s, i)

    doc.add_page_break()

    # ====================================================================
    # BAB III: PANDUAN PENGUNJUNG
    # ====================================================================
    add_heading(doc, "BAB III  PANDUAN PENGUNJUNG / WARGA", level=1, color=(0x1F, 0x7A, 0x4D))
    add_hr(doc)

    add_heading(doc, "3.1  Mengakses Website", level=2)
    add_body(doc, "Website dapat diakses melalui berbagai cara:")
    add_bullet(doc, "Browser di HP/Smartphone: buka browser (Chrome/Safari), ketikkan alamat URL website desa.")
    add_bullet(doc, "Scan QR Code: arahkan kamera HP ke QR Code yang tersedia di gapura atau papan pengumuman desa.")
    add_bullet(doc, "Browser di Komputer/Laptop: buka browser dan masukkan alamat URL website.")
    add_body(doc, "Catatan: Pastikan perangkat terhubung ke internet (WiFi atau paket data) untuk mengakses website.")

    add_heading(doc, "3.2  Halaman Beranda", level=2)
    add_body(doc, "Ketika website pertama kali dibuka, pengunjung akan melihat halaman beranda yang terdiri dari:")
    add_bullet(doc, "Navbar (Bilah Navigasi Atas): Logo dan nama desa, menu navigasi, tombol pencarian.")
    add_bullet(doc, "Hero/Banner Utama: Foto latar desa, logo desa, nama desa, dan motto.")
    add_bullet(doc, "Status Kantor Desa: Indikator real-time apakah kantor desa sedang buka (hijau) atau tutup (merah).")
    add_bullet(doc, "Jam WIB Berjalan: Tampilan waktu saat ini dalam zona waktu WIB (GMT+7) yang bergerak otomatis setiap detik, disertai informasi hari dan tanggal.")
    add_bullet(doc, "Grid Menu: 9 menu informasi desa yang dapat diklik untuk membuka masing-masing halaman.")

    add_heading(doc, "3.3  Fitur Jam WIB & Status Kantor", level=2)
    add_body(doc, "Fitur jam berjalan menampilkan waktu real-time dalam format:")
    add_body(doc, "   [Hari, Tanggal Bulan Tahun] | [🕐 HH:MM:SS] WIB", indent=True)
    add_body(doc, "Contoh: Minggu, 14 Sep 2026 | 🕐 21:04:06 WIB")
    doc.add_paragraph()
    add_body(doc, "Status kantor desa diperbarui otomatis setiap detik berdasarkan jam dan hari:")

    jadwal_tbl = doc.add_table(rows=4, cols=3)
    jadwal_tbl.style = 'Table Grid'
    add_table_header(jadwal_tbl, ["Hari", "Jam Layanan", "Status"])
    jadwal_data = [
        ("Senin – Jumat", "08:00 – 16:00 WIB", "Buka (hijau)"),
        ("Sabtu", "08:00 – 12:00 WIB", "Buka (hijau)"),
        ("Minggu / Hari Libur", "—", "Tutup (merah)"),
    ]
    for i, row_data in enumerate(jadwal_data):
        add_table_row(jadwal_tbl, i+1, row_data, alt=(i%2==1))

    doc.add_paragraph()

    add_heading(doc, "3.4  Menu Profil Desa", level=2)
    add_body(doc, "Menampilkan informasi lengkap tentang desa, meliputi:")
    profil_items = [
        "Nama desa, logo, dan motto",
        "Tahun berdirinya desa",
        "Data statistik: luas wilayah, jumlah penduduk, jumlah kepala keluarga",
        "Visi dan Misi desa",
        "Sejarah singkat desa",
        "Potensi desa (dalam bentuk label/tag)",
        "Alamat kantor desa",
        "Peta lokasi (Google Maps Embed) — klik peta untuk membuka di aplikasi Maps",
    ]
    for item in profil_items:
        add_bullet(doc, item)
    add_body(doc, "Cara membuka: Klik menu 'Profil Desa' di halaman beranda atau pilih dari menu navigasi.")

    add_heading(doc, "3.5  Menu Perangkat Desa", level=2)
    add_body(doc, "Menampilkan daftar lengkap aparatur dan pengurus desa. Setiap kartu perangkat memuat:")
    add_bullet(doc, "Foto perangkat desa")
    add_bullet(doc, "Nama lengkap")
    add_bullet(doc, "Jabatan/posisi")
    add_bullet(doc, "Periode jabatan")
    add_bullet(doc, "Tugas dan tanggung jawab")
    add_body(doc, "Kepala Desa ditampilkan di posisi paling atas dengan tanda bintang (⭐).")
    add_body(doc, "Klik kartu untuk melihat detail lengkap di popup/modal.")

    add_heading(doc, "3.6  Menu UMKM Desa", level=2)
    add_body(doc, "Menampilkan direktori usaha warga desa. Setiap kartu UMKM berisi:")
    add_bullet(doc, "Foto produk/usaha")
    add_bullet(doc, "Nama usaha dan kategori")
    add_bullet(doc, "Nama pemilik")
    add_bullet(doc, "Deskripsi singkat usaha")
    add_body(doc, "Klik kartu UMKM untuk melihat detail lengkap termasuk:")
    add_bullet(doc, "Alamat lengkap")
    add_bullet(doc, "Nomor kontak / WhatsApp pemilik")
    add_bullet(doc, "Jam buka")
    add_bullet(doc, "Peta lokasi usaha (jika tersedia)")

    add_heading(doc, "3.7  Menu Berita & Pengumuman", level=2)
    add_body(doc, "Menampilkan berita terbaru dan pengumuman resmi dari pemerintah desa. Fitur:")
    add_bullet(doc, "Tampilan kartu berita dengan foto, judul, dan ringkasan")
    add_bullet(doc, "Tanggal publikasi")
    add_bullet(doc, "Klik kartu untuk membaca berita lengkap di popup")
    add_body(doc, "Berita ditampilkan dari yang terbaru ke yang terlama.")

    add_heading(doc, "3.8  Menu Agenda & Kegiatan", level=2)
    add_body(doc, "Menampilkan jadwal kegiatan, rapat, dan acara yang akan diselenggarakan desa. Informasi yang tersedia:")
    add_bullet(doc, "Judul kegiatan")
    add_bullet(doc, "Tanggal dan waktu pelaksanaan")
    add_bullet(doc, "Lokasi kegiatan")
    add_bullet(doc, "Kategori kegiatan (Rapat, Sosial, Pemerintahan, dll.)")
    add_bullet(doc, "Deskripsi detail kegiatan")
    add_bullet(doc, "Foto kegiatan (jika tersedia)")

    add_heading(doc, "3.9  Menu Dokumen & Surat Layanan", level=2)
    add_body(doc, "Menyediakan akses ke berbagai dokumen dan formulir yang dapat diunduh warga:")
    add_bullet(doc, "Formulir permohonan (KTP, KK, Surat Keterangan, dll.)")
    add_bullet(doc, "Peraturan Desa (Perdes)")
    add_bullet(doc, "Laporan dan publikasi resmi desa")
    add_body(doc, "Cara mengunduh: Klik tombol 'Unduh' atau 'Download' pada dokumen yang diinginkan.")
    add_body(doc, "Catatan: Setiap dokumen menampilkan kategori, deskripsi singkat, dan ukuran file.")

    add_heading(doc, "3.10  Menu Pengaduan & Aspirasi Warga", level=2)
    add_body(doc, "Fitur ini memungkinkan warga menyampaikan laporan, saran, dan pertanyaan langsung kepada pemerintah desa secara digital.")
    add_body(doc, "Cara menyampaikan pengaduan:")
    add_numbered(doc, "Klik menu 'Pengaduan & Aspirasi Warga' dari halaman beranda.", 1)
    add_numbered(doc, "Isi formulir yang tersedia:", 2)
    add_bullet(doc, "Nama lengkap (wajib diisi)", level=1)
    add_bullet(doc, "Nomor kontak / HP (untuk pemberitahuan balasan)", level=1)
    add_bullet(doc, "Subjek / judul laporan", level=1)
    add_bullet(doc, "Isi lengkap laporan atau aspirasi", level=1)
    add_numbered(doc, "Klik tombol 'Kirim Pengaduan'.", 3)
    add_numbered(doc, "Tunggu balasan dari pihak desa. Status pengaduan dapat berupa: Pending, Diproses, atau Selesai.", 4)
    add_body(doc, "Warga dapat melihat status pengaduan mereka di bagian 'Cek Status Laporan' dengan memasukkan nomor kontak yang terdaftar.")

    add_heading(doc, "3.11  Menu Galeri Foto", level=2)
    add_body(doc, "Menampilkan dokumentasi foto kegiatan dan potensi desa. Fitur:")
    add_bullet(doc, "Tampilan grid foto yang menarik")
    add_bullet(doc, "Kategori dan tanggal foto")
    add_bullet(doc, "Klik foto untuk memperbesar dan melihat deskripsi lengkap")

    add_heading(doc, "3.12  Menu Kontak & Lokasi", level=2)
    add_body(doc, "Menampilkan informasi kontak lengkap kantor desa:")
    add_bullet(doc, "Alamat lengkap kantor desa")
    add_bullet(doc, "Nomor telepon / WhatsApp")
    add_bullet(doc, "Alamat email resmi")
    add_bullet(doc, "Jam layanan kantor")
    add_bullet(doc, "Tombol Chat WhatsApp langsung")
    add_bullet(doc, "Peta lokasi interaktif (Google Maps)")
    add_bullet(doc, "Link media sosial resmi desa (Facebook, Instagram)")

    add_heading(doc, "3.13  Fitur Pencarian Global", level=2)
    add_body(doc, "Fitur pencarian memudahkan warga menemukan informasi spesifik tanpa harus menelusuri setiap menu satu per satu.")
    add_body(doc, "Cara menggunakan:")
    add_numbered(doc, "Klik ikon kaca pembesar (🔍) di navbar atau tombol 'Cari' di navigasi bawah (mobile).", 1)
    add_numbered(doc, "Ketikkan kata kunci di kolom pencarian (contoh: 'KTP', 'Pak Kades', 'Kopi', 'Rapat').", 2)
    add_numbered(doc, "Hasil pencarian otomatis muncul dari berbagai kategori: berita, UMKM, perangkat desa, dokumen, dan agenda.", 3)
    add_numbered(doc, "Klik hasil yang relevan untuk membuka detail informasi.", 4)

    add_heading(doc, "3.14  Asisten Desa (Chatbot AI)", level=2)
    add_body(doc,
        "Asisten Desa adalah fitur tanya-jawab cerdas berbasis kecerdasan buatan (AI) "
        "yang dapat menjawab pertanyaan warga tentang desa secara instan, 24 jam sehari, "
        "7 hari seminggu."
    )
    add_body(doc, "Cara menggunakan Asisten Desa:")
    add_numbered(doc, "Klik tombol '🤖 Tanya Desa' yang mengambang di pojok kanan bawah layar.", 1)
    add_numbered(doc, "Panel chat akan terbuka. Anda akan disambut dengan salam dari Asisten Desa.", 2)
    add_numbered(doc, "Pilih topik dari tombol cepat yang tersedia, atau ketik pertanyaan Anda sendiri.", 3)
    add_numbered(doc, "Asisten akan membalas dalam beberapa detik dengan informasi terkini dari database desa.", 4)
    add_body(doc, "Topik yang dapat ditanyakan:")
    ai_topics = [
        "Profil dan sejarah desa",
        "Perangkat dan pengurus desa",
        "Cara mengurus surat dan dokumen",
        "Informasi UMKM dan usaha warga",
        "Jadwal kegiatan dan agenda",
        "Cara menghubungi kantor desa",
    ]
    for t in ai_topics:
        add_bullet(doc, t)
    add_body(doc, "Batasan penggunaan: Maksimal 30 pertanyaan per hari per perangkat untuk menjaga stabilitas layanan.")
    add_body(doc, "Catatan Penting: Jawaban Asisten Desa berasal dari AI dan data yang tersedia — selalu konfirmasi ke kantor desa untuk kepastian resmi.", bold=False)

    add_heading(doc, "3.15  Navigasi Mobile (Smartphone)", level=2)
    add_body(doc, "Di perangkat smartphone, tersedia navigasi bawah layar (bottom navigation) yang memudahkan akses cepat:")

    nav_items = [
        ("🏡 Home", "Kembali ke halaman beranda"),
        ("📰 Berita", "Buka langsung menu berita & pengumuman"),
        ("🔍 Cari", "Buka fitur pencarian global"),
        ("📄 Surat", "Buka menu dokumen & surat layanan"),
        ("💬 Lapor", "Buka menu pengaduan warga"),
    ]
    for icon, desc in nav_items:
        p = doc.add_paragraph()
        r1 = p.add_run(f"  {icon}: ")
        r1.font.bold = True
        r1.font.size = Pt(10.5)
        r2 = p.add_run(desc)
        r2.font.size = Pt(10.5)
        set_para_spacing(p, before=1, after=2)

    add_body(doc, "Fitur swipe-back: Di perangkat mobile, Anda dapat menggeser ke kiri/kanan atau menekan tombol 'Back' bawaan HP untuk kembali ke halaman beranda tanpa keluar dari website.")

    doc.add_page_break()

    # ====================================================================
    # BAB IV: PANDUAN ADMIN
    # ====================================================================
    add_heading(doc, "BAB IV  PANDUAN ADMINISTRATOR DESA", level=1, color=(0x1F, 0x7A, 0x4D))
    add_hr(doc)

    add_body(doc,
        "Panel administrasi (admin panel) adalah area khusus yang hanya dapat diakses "
        "oleh operator/administrator desa yang memiliki akun terdaftar. Seluruh pengelolaan "
        "konten website dilakukan melalui panel ini."
    )

    add_heading(doc, "4.1  Akses Panel Admin", level=2)
    add_body(doc, "Panel admin dapat diakses melalui URL:")
    p_url = doc.add_paragraph()
    r_url = p_url.add_run("   [URL Website Desa]/admin.html")
    r_url.font.bold = True
    r_url.font.size = Pt(11)
    r_url.font.color.rgb = RGBColor(0x1F, 0x7A, 0x4D)
    p_url.paragraph_format.left_indent = Cm(0.5)
    set_para_spacing(p_url, before=2, after=6)

    add_body(doc, "Contoh: https://webdesa-kkn.vercel.app/admin.html")

    add_heading(doc, "4.2  Login & Keamanan Sesi", level=2)
    add_body(doc, "Langkah login:")
    add_numbered(doc, "Buka halaman admin (lihat URL di atas).", 1)
    add_numbered(doc, "Masukkan alamat email dan password yang terdaftar.", 2)
    add_numbered(doc, "Klik tombol 'Masuk / Login'.", 3)
    add_numbered(doc, "Jika berhasil, Anda akan langsung masuk ke dashboard admin.", 4)

    add_body(doc, "Keamanan Sesi:")
    add_bullet(doc, "Sistem menggunakan Supabase Auth untuk autentikasi yang aman dan terenkripsi.")
    add_bullet(doc, "Session Timeout 12 Jam: Jika tidak ada aktivitas (klik, scroll, atau ketik) selama 12 jam, sistem akan otomatis logout dan meminta login ulang.")
    add_bullet(doc, "Aktivitas yang dihitung: gerakan mouse, klik, scroll, dan pengetikan.")
    add_bullet(doc, "Pesan peringatan akan muncul sebelum sesi berakhir.")

    add_heading(doc, "4.3  Dashboard Admin", level=2)
    add_body(doc, "Setelah login, Anda akan melihat tampilan dashboard dengan komponen:")
    add_bullet(doc, "Sidebar Navigasi (kiri): Menu tab untuk berpindah antar modul pengelolaan.")
    add_bullet(doc, "Topbar (atas): Judul halaman aktif dan tombol Logout.")
    add_bullet(doc, "Area Konten (kanan): Form dan tabel data sesuai modul yang dipilih.")
    add_body(doc, "Daftar modul di sidebar:")
    modul_list = ["Profil Desa", "Perangkat Desa", "UMKM Desa", "Agenda Desa", "Dokumen & Surat", "Pengaduan Warga", "Galeri Foto"]
    for m in modul_list:
        add_bullet(doc, m)

    add_heading(doc, "4.4  Mengelola Profil Desa", level=2)
    add_body(doc, "Modul ini digunakan untuk mengatur seluruh informasi utama desa yang tampil di halaman publik.")
    add_body(doc, "Field yang dapat diisi/diubah:")
    profil_fields = [
        ("Nama Desa", "Nama resmi desa"),
        ("Motto", "Semboyan atau tagline desa"),
        ("Logo Desa", "Upload foto logo desa (format JPG/PNG)"),
        ("Foto/Background Desa", "Foto yang tampil sebagai latar hero beranda"),
        ("Tahun Berdiri", "Tahun pendirian desa"),
        ("Visi", "Visi pemerintahan desa"),
        ("Misi", "Misi pemerintahan desa (pisah baris dengan Enter)"),
        ("Sejarah", "Sejarah singkat desa"),
        ("Luas Wilayah", "Contoh: 245 Ha"),
        ("Jumlah Penduduk", "Contoh: 3.450 jiwa"),
        ("Jumlah KK", "Jumlah kepala keluarga"),
        ("Potensi Desa", "Pisahkan dengan koma, contoh: Pertanian, Perikanan, Kerajinan"),
        ("Alamat Kantor", "Alamat lengkap kantor desa"),
        ("Telepon/WA", "Nomor telepon (format internasional: 628xxx)"),
        ("Email", "Email resmi desa"),
        ("Jam Layanan", "Contoh: Senin-Jumat 08.00-16.00"),
        ("Google Maps URL", "URL embed Google Maps — WAJIB format embed"),
        ("Facebook", "URL halaman Facebook desa"),
        ("Instagram", "URL profil Instagram desa"),
        ("Warna Aksen", "Kode warna HEX untuk tema website (contoh: #1f7a4d)"),
    ]
    for field, desc in profil_fields:
        add_info_box(doc, field, desc)

    add_body(doc, "PENTING — Cara mendapatkan URL Google Maps yang benar:")
    add_numbered(doc, "Buka Google Maps di browser.", 1)
    add_numbered(doc, "Cari lokasi kantor desa.", 2)
    add_numbered(doc, "Klik 'Bagikan' → pilih tab 'Sematkan peta'.", 3)
    add_numbered(doc, "Klik 'Salin HTML'.", 4)
    add_numbered(doc, "Dari kode yang disalin, ambil hanya bagian URL di dalam src=\"...\".", 5)
    add_numbered(doc, "Contoh hasil yang benar: https://www.google.com/maps/embed?pb=!1m18...", 6)
    add_body(doc, "Menyimpan: Klik tombol 'Simpan Profil' di bawah form setelah semua field terisi.")

    add_heading(doc, "4.5  Mengelola Perangkat Desa", level=2)
    add_body(doc, "Menambahkan, mengubah, dan menghapus data aparatur desa.")
    add_body(doc, "Cara menambah perangkat baru:")
    add_numbered(doc, "Klik tombol '+ Tambah Perangkat'.", 1)
    add_numbered(doc, "Isi formulir: Nama, Jabatan, Periode Jabatan, Foto (upload atau URL), Tugas & Tanggung Jawab, Urutan Tampil.", 2)
    add_numbered(doc, "Centang 'Ini adalah Kepala Desa' jika data ini adalah Kades (otomatis tampil paling atas).", 3)
    add_numbered(doc, "Klik 'Simpan'.", 4)
    add_body(doc, "Cara mengedit: Klik tombol 'Edit' pada baris data yang ingin diubah.")
    add_body(doc, "Cara menghapus: Klik tombol 'Hapus' dan konfirmasi penghapusan.")

    add_heading(doc, "4.6  Mengelola UMKM Desa", level=2)
    add_body(doc, "Mengelola direktori usaha warga desa.")
    add_body(doc, "Data yang dapat diisi per UMKM:")
    umkm_fields = ["Nama Usaha", "Nama Pemilik", "Kategori (Kuliner, Kerajinan, Jasa, dll.)", "Deskripsi Usaha", "Nomor Kontak WhatsApp", "Alamat Usaha", "Jam Buka", "Google Maps Lokasi Usaha", "Foto Usaha/Produk", "Urutan Tampil"]
    for f in umkm_fields:
        add_bullet(doc, f)
    add_body(doc, "Cara menambah, mengedit, dan menghapus sama seperti modul Perangkat Desa.")

    add_heading(doc, "4.7  Mengelola Agenda & Kegiatan", level=2)
    add_body(doc, "Menjadwalkan dan mengelola agenda kegiatan desa.")
    add_body(doc, "Data per agenda:")
    agenda_fields = ["Judul Kegiatan", "Tanggal (format: YYYY-MM-DD)", "Waktu (format: HH:MM)", "Lokasi", "Kategori", "Deskripsi Lengkap", "Foto Kegiatan"]
    for f in agenda_fields:
        add_bullet(doc, f)

    add_heading(doc, "4.8  Mengelola Dokumen & Surat", level=2)
    add_body(doc, "Mengelola koleksi dokumen dan formulir yang dapat diunduh warga.")
    add_body(doc, "Data per dokumen:")
    dok_fields = ["Judul Dokumen", "Kategori (Formulir, Perdes, Laporan, dll.)", "Deskripsi", "URL File (link Google Drive, Dropbox, atau URL langsung)", "Ukuran File (contoh: 245 KB)"]
    for f in dok_fields:
        add_bullet(doc, f)
    add_body(doc, "Catatan: Untuk file dokumen, gunakan Google Drive dengan pengaturan 'Siapa saja yang memiliki link dapat melihat', lalu salin link-nya.")

    add_heading(doc, "4.9  Mengelola Pengaduan Warga", level=2)
    add_body(doc, "Admin dapat melihat seluruh pengaduan yang masuk dan memberikan tanggapan resmi.")
    add_body(doc, "Alur pengelolaan pengaduan:")
    add_numbered(doc, "Masuk ke tab 'Pengaduan Warga' di dashboard admin.", 1)
    add_numbered(doc, "Lihat daftar pengaduan yang masuk. Setiap pengaduan menampilkan: nama pelapor, kontak, subjek, isi laporan, dan status terkini.", 2)
    add_numbered(doc, "Klik tombol 'Respon' pada pengaduan yang ingin ditanggapi.", 3)
    add_numbered(doc, "Pilih status pengaduan: Pending / Diproses / Selesai.", 4)
    add_numbered(doc, "Tulis tanggapan resmi di kolom tanggapan.", 5)
    add_numbered(doc, "Klik 'Simpan Tanggapan'.", 6)
    add_body(doc, "Tanggapan dan status yang diperbarui akan langsung terlihat oleh warga ketika mereka mengecek status laporan di halaman publik.")

    add_heading(doc, "4.10  Mengelola Galeri Foto", level=2)
    add_body(doc, "Mendokumentasikan kegiatan desa melalui foto.")
    add_body(doc, "Data per foto galeri:")
    galeri_fields = ["Judul Foto", "Kategori (Kegiatan, Infrastruktur, Sosial, dll.)", "Tanggal", "Deskripsi", "Foto (upload file atau masukkan URL)", "Urutan Tampil"]
    for f in galeri_fields:
        add_bullet(doc, f)

    add_heading(doc, "4.11  Logout & Keamanan", level=2)
    add_body(doc, "Cara logout manual:")
    add_numbered(doc, "Klik tombol 'Keluar' / 'Logout' di bagian kanan atas topbar dashboard.", 1)
    add_numbered(doc, "Sistem akan menghapus sesi dan mengarahkan kembali ke halaman login.", 2)
    add_body(doc, "REKOMENDASI KEAMANAN:")
    security_tips = [
        "Selalu logout setelah selesai bekerja, terutama jika menggunakan komputer/HP bersama.",
        "Jangan bagikan email dan password admin kepada siapa pun.",
        "Ganti password secara berkala melalui email/dashboard Supabase.",
        "Pastikan menggunakan password yang kuat (minimal 8 karakter, kombinasi huruf, angka, dan simbol).",
        "Sistem akan otomatis logout setelah 12 jam tidak ada aktivitas sebagai perlindungan tambahan.",
    ]
    for tip in security_tips:
        add_bullet(doc, tip)

    doc.add_page_break()

    # ====================================================================
    # BAB V: KEAMANAN
    # ====================================================================
    add_heading(doc, "BAB V  FITUR KEAMANAN SISTEM", level=1, color=(0x1F, 0x7A, 0x4D))
    add_hr(doc)

    add_heading(doc, "5.1  Autentikasi Admin", level=2)
    add_body(doc,
        "Sistem menggunakan Supabase Auth — layanan autentikasi enterprise-grade dengan "
        "enkripsi standar industri. Password admin tidak pernah disimpan dalam bentuk teks "
        "biasa, melainkan dalam format hash yang tidak dapat dibaca."
    )
    add_body(doc, "Mekanisme keamanan:")
    add_bullet(doc, "Email + Password — autentikasi dua faktor dasar")
    add_bullet(doc, "JWT Token — setiap sesi diautentikasi dengan token terenkripsi")
    add_bullet(doc, "HTTPS — seluruh komunikasi data terenkripsi")
    add_bullet(doc, "Row Level Security (RLS) — hanya admin yang terotentikasi dapat menulis data")

    add_heading(doc, "5.2  Session Timeout 12 Jam", level=2)
    add_body(doc, "Sistem melacak waktu aktivitas terakhir admin menggunakan localStorage browser. Jika rentang waktu sejak aktivitas terakhir melebihi 12 jam, sistem akan:")
    add_numbered(doc, "Secara otomatis memanggil fungsi signOut() dari Supabase Auth.", 1)
    add_numbered(doc, "Menghapus data aktivitas dari penyimpanan lokal.", 2)
    add_numbered(doc, "Menampilkan pesan: 'Sesi Anda telah berakhir karena tidak ada aktivitas selama 12 jam.'", 3)
    add_numbered(doc, "Memuat ulang halaman dan menampilkan form login.", 4)
    add_body(doc, "Aktivitas yang dihitung sebagai 'masih aktif': gerakan mouse, klik mouse, penekanan tombol keyboard, dan scroll halaman.")

    add_heading(doc, "5.3  Rate Limiting Chatbot", level=2)
    add_body(doc, "Untuk mencegah penyalahgunaan layanan AI chatbot, sistem menerapkan pembatasan:")
    add_bullet(doc, "Maksimum 30 pertanyaan per hari per perangkat/browser (tersimpan di localStorage).")
    add_bullet(doc, "Jeda minimal 2 detik antara setiap pertanyaan.")
    add_bullet(doc, "Di sisi server: maksimum 15 permintaan per menit per alamat IP.")
    add_body(doc, "Jika batas terlampaui, sistem akan menampilkan pesan yang jelas kepada pengguna.")

    doc.add_page_break()

    # ====================================================================
    # BAB VI: TROUBLESHOOTING
    # ====================================================================
    add_heading(doc, "BAB VI  TROUBLESHOOTING", level=1, color=(0x1F, 0x7A, 0x4D))
    add_hr(doc)

    add_heading(doc, "6.1  Masalah Umum & Solusi", level=2)

    problems = [
        (
            "Website tidak bisa dibuka / loading terus",
            [
                "Periksa koneksi internet perangkat Anda.",
                "Coba buka di browser lain (Chrome, Firefox).",
                "Hapus cache browser: Ctrl+Shift+Delete (Windows) atau Cmd+Shift+Delete (Mac).",
                "Coba akses dari jaringan internet lain (misalnya beralih dari WiFi ke data seluler).",
            ]
        ),
        (
            "Data tidak muncul di menu (menampilkan loading spinner terus)",
            [
                "Pastikan koneksi internet stabil.",
                "Muat ulang halaman dengan Ctrl+F5 (Windows) atau Cmd+Shift+R (Mac).",
                "Jika masalah berlanjut, hubungi tim teknis — kemungkinan ada gangguan pada server Supabase.",
            ]
        ),
        (
            "Peta Google Maps tidak muncul",
            [
                "Pastikan URL yang dimasukkan admin adalah format EMBED Google Maps (bukan URL biasa).",
                "URL yang benar diawali dengan: https://www.google.com/maps/embed?pb=...",
                "Jika URL bukan format embed, sistem akan menampilkan tombol 'Buka di Google Maps' sebagai alternatif.",
            ]
        ),
        (
            "Admin tidak bisa login",
            [
                "Periksa kembali ejaan email dan password.",
                "Pastikan Caps Lock tidak aktif.",
                "Jika lupa password, gunakan fitur 'Lupa Password' atau hubungi pengelola sistem untuk reset password melalui dashboard Supabase.",
                "Coba hapus cache dan cookie browser lalu coba login ulang.",
            ]
        ),
        (
            "Admin diotomatis logout terlalu cepat",
            [
                "Sistem timeout diatur 12 jam sejak terakhir ada aktivitas.",
                "Pastikan Anda tetap berinteraksi dengan halaman selama bekerja.",
                "Jika dirasa terlalu singkat, hubungi tim teknis untuk menyesuaikan batas waktu.",
            ]
        ),
        (
            "Chatbot tidak merespons / error",
            [
                "Periksa koneksi internet.",
                "Tunggu beberapa saat dan coba lagi — kemungkinan layanan AI sedang sibuk.",
                "Pastikan Anda belum melebihi batas 30 pertanyaan per hari.",
                "Jika masalah berlanjut, hubungi tim teknis — kemungkinan ada masalah pada GROQ_API_KEY atau Edge Function.",
            ]
        ),
        (
            "Foto tidak muncul setelah upload",
            [
                "Pastikan ukuran file foto tidak melebihi batas (disarankan maksimal 5 MB).",
                "Format file yang didukung: JPG, JPEG, PNG, WebP.",
                "Muat ulang halaman admin setelah upload.",
                "Periksa apakah bucket Supabase Storage sudah dikonfigurasi sebagai 'Public'.",
            ]
        ),
        (
            "Pengaduan warga tidak muncul di admin",
            [
                "Pastikan migrasi database (migration_v3.sql) sudah dijalankan di Supabase.",
                "Cek di Supabase Dashboard bahwa tabel 'pengaduan' sudah ada.",
                "Refresh halaman admin dan klik tab Pengaduan Warga kembali.",
            ]
        ),
    ]

    for title, solutions in problems:
        p = doc.add_paragraph()
        r = p.add_run(f"Masalah: {title}")
        r.font.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = RGBColor(0xC0, 0x39, 0x27)
        set_para_spacing(p, before=10, after=2)

        p2 = doc.add_paragraph()
        r2 = p2.add_run("Solusi:")
        r2.font.bold = True
        r2.font.size = Pt(10.5)
        set_para_spacing(p2, before=0, after=2)

        for sol in solutions:
            add_bullet(doc, sol)

    add_heading(doc, "6.2  Kontak Bantuan Teknis", level=2)
    add_body(doc, "Jika Anda mengalami masalah yang tidak tercantum di atas atau membutuhkan bantuan lebih lanjut, silakan hubungi:")
    add_bullet(doc, "Tim Pengembang: Kelompok 1 KKN Global Institut")
    add_bullet(doc, "Melalui: [Nomor WhatsApp Tim Teknis]")
    add_bullet(doc, "Email: [Email Tim Teknis]")
    add_body(doc, "Saat menghubungi, mohon sertakan:")
    add_numbered(doc, "Deskripsi masalah secara detail.", 1)
    add_numbered(doc, "Screenshot atau foto tampilan error (jika ada).", 2)
    add_numbered(doc, "Perangkat dan browser yang digunakan.", 3)

    doc.add_page_break()

    # ====================================================================
    # BAB VII: PENUTUP
    # ====================================================================
    add_heading(doc, "BAB VII  PENUTUP", level=1, color=(0x1F, 0x7A, 0x4D))
    add_hr(doc)

    p_close = doc.add_paragraph()
    p_close.add_run(
        "Demikian Manual Book Website Informasi Desa ini disusun sebagai panduan komprehensif "
        "bagi seluruh pengguna sistem, baik administrator desa maupun warga masyarakat pada umumnya.\n\n"
        "Website Informasi Desa ini dirancang dengan prinsip kemudahan penggunaan (user-friendly) "
        "sehingga dapat digunakan oleh semua kalangan usia tanpa memerlukan keahlian teknologi khusus. "
        "Sistem ini akan terus dikembangkan dan ditingkatkan sesuai kebutuhan desa.\n\n"
        "Kami berharap sistem ini dapat memberikan manfaat nyata dalam meningkatkan kualitas "
        "pelayanan publik dan transparansi pemerintahan desa. Partisipasi aktif seluruh warga "
        "dalam menggunakan fitur-fitur yang tersedia akan sangat membantu terwujudnya desa "
        "yang informatif, transparan, dan modern.\n\n"
        "Terima kasih atas kepercayaan yang diberikan kepada Kelompok 1 KKN Global Institut "
        "dalam mengembangkan sistem informasi digital untuk kemajuan desa."
    ).font.size = Pt(11)
    set_para_spacing(p_close, before=0, after=8)

    doc.add_paragraph()
    p_sign = doc.add_paragraph()
    p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sign.add_run(f"Hormat kami,\n\n\n\n\nKelompok 1 KKN Global Institut\n{datetime.now().strftime('%B %Y')}").font.size = Pt(11)

    doc.add_page_break()

    # ====================================================================
    # LAMPIRAN: REFERENSI CEPAT
    # ====================================================================
    add_heading(doc, "LAMPIRAN — REFERENSI CEPAT ADMIN", level=1, color=(0x1F, 0x7A, 0x4D))
    add_hr(doc)

    add_heading(doc, "Tabel Referensi Menu Admin", level=2)
    quick_ref = [
        ("Profil Desa", "Pengaturan nama, logo, visi-misi, kontak, warna tema"),
        ("Perangkat Desa", "CRUD data aparatur desa + foto"),
        ("UMKM Desa", "CRUD data usaha warga + foto + peta"),
        ("Agenda Desa", "CRUD jadwal kegiatan + foto"),
        ("Dokumen & Surat", "CRUD dokumen & formulir + URL unduhan"),
        ("Pengaduan Warga", "Lihat laporan warga + beri tanggapan + ubah status"),
        ("Galeri Foto", "CRUD foto kegiatan desa"),
    ]

    tbl2 = doc.add_table(rows=len(quick_ref)+1, cols=2)
    tbl2.style = 'Table Grid'
    add_table_header(tbl2, ["Modul", "Fungsi Utama"])
    for i, (m, f) in enumerate(quick_ref):
        add_table_row(tbl2, i+1, [m, f], alt=(i%2==1))

    doc.add_paragraph()
    add_heading(doc, "Checklist Setup Awal Website Desa", level=2)
    checklist = [
        "Jalankan migration SQL di Supabase (migration_v3.sql + migration_v2.sql untuk galeri)",
        "Konfigurasi SUPABASE_URL dan SUPABASE_ANON_KEY di file assets/js/config.js",
        "Buat akun admin melalui Supabase Dashboard > Authentication",
        "Upload logo desa dan foto latar di menu Profil Desa",
        "Isi semua field Profil Desa (nama, motto, visi, misi, kontak)",
        "Tambahkan data perangkat desa minimal Kepala Desa",
        "Konfigurasi Google Maps embed URL untuk lokasi kantor",
        "Tambahkan minimal beberapa data UMKM, berita, dan dokumen",
        "Deploy ke Vercel dan hubungkan dengan repositori GitHub",
        "Uji coba seluruh fitur dari perangkat desktop dan mobile",
        "Atur GROQ_API_KEY di Supabase Edge Function Secrets untuk Chatbot AI",
    ]
    for item in checklist:
        p = doc.add_paragraph()
        r = p.add_run(f"  ☐  {item}")
        r.font.size = Pt(10.5)
        set_para_spacing(p, before=1, after=2)

    # Save
    output_path = r"c:\Users\1nkuss\Documents\WebInformasiDesa\Manual_Book_Website_Informasi_Desa.docx"
    doc.save(output_path)
    print(f"Manual book berhasil dibuat: {output_path}")
    return output_path

if __name__ == "__main__":
    create_manual()
