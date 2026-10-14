import streamlit as st
import sqlite3
import datetime
import os
import subprocess
import tempfile
import pandas as pd
from docx import Document
from generator_bast import generate_bast_docx

# Cross-platform, permission-safe paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_NAME = os.path.join(BASE_DIR, "bast_database.db")
TEMP_OUT_DIR = os.path.join(tempfile.gettempdir(), "bast_outputs")
os.makedirs(TEMP_OUT_DIR, exist_ok=True)

def get_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    
    # 1. Tabel Master BAST
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS bast_records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nomor_bast TEXT UNIQUE NOT NULL,
        tanggal_bast TEXT NOT NULL,
        hari TEXT,
        pihak_kesatu_nama TEXT DEFAULT 'dr. Riswandi, M.M.',
        pihak_kesatu_nip TEXT DEFAULT '19741226 200604 1 002',
        pihak_kesatu_jabatan TEXT DEFAULT 'Plt. Direktur Penanganan Darurat Wilayah II',
        pihak_kedua_nama TEXT NOT NULL,
        pihak_kedua_nip TEXT,
        pihak_kedua_jabatan TEXT,
        pihak_kedua_instansi TEXT NOT NULL,
        pihak_kedua_alamat TEXT,
        pihak_kedua_hp TEXT,
        jenis_bantuan TEXT,
        total_nilai REAL NOT NULL,
        status TEXT DEFAULT 'COMPLETED',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')

    try:
        cursor.execute("ALTER TABLE bast_records ADD COLUMN status TEXT DEFAULT 'COMPLETED'")
    except sqlite3.OperationalError:
        pass

    # 2. Tabel Detail Barang
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS bast_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        bast_id INTEGER NOT NULL,
        nama_barang TEXT NOT NULL,
        penyedia TEXT NOT NULL,
        qty INTEGER NOT NULL,
        satuan TEXT NOT NULL,
        harga_satuan REAL NOT NULL,
        total_harga REAL NOT NULL,
        FOREIGN KEY (bast_id) REFERENCES bast_records (id) ON DELETE CASCADE
    )
    ''')

    # 3. Tabel Activity Log
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS activity_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nomor_bast TEXT,
        action TEXT NOT NULL,
        user TEXT DEFAULT 'System/Operator',
        details TEXT,
        timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')

    conn.commit()
    conn.close()

def generate_next_bast_number(kode_instansi="BNPB"):
    now = datetime.datetime.now()
    tahun = str(now.year)
    romawi_map = {1:'I', 2:'II', 3:'III', 4:'IV', 5:'V', 6:'VI', 7:'VII', 8:'VIII', 9:'IX', 10:'X', 11:'XI', 12:'XII'}
    bulan_romawi = romawi_map[now.month]

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) as count FROM bast_records WHERE nomor_bast LIKE ?", (f"%/{tahun}",))
    row = cursor.fetchone()
    next_seq = (row['count'] if row else 0) + 1
    seq_str = f"{next_seq:03d}"
    
    nomor_bast = f"{seq_str}/BAST/BNPB/{kode_instansi.upper()}/{bulan_romawi}/{tahun}"
    conn.close()
    return nomor_bast

def get_historical_data():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT penyedia FROM bast_items")
    penyedia_list = [row['penyedia'] for row in cursor.fetchall()]
    
    cursor.execute("SELECT DISTINCT nama_barang FROM bast_items")
    barang_list = [row['nama_barang'] for row in cursor.fetchall()]

    cursor.execute("SELECT DISTINCT jenis_bantuan FROM bast_records")
    bantuan_list = [row['jenis_bantuan'] for row in cursor.fetchall()]

    conn.close()
    return penyedia_list, barang_list, bantuan_list

def get_item_historical_price(nama_barang):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT harga_satuan, penyedia, satuan
        FROM bast_items 
        WHERE LOWER(nama_barang) = LOWER(?)
        ORDER BY id DESC LIMIT 1
    ''', (nama_barang,))
    row = cursor.fetchone()
    conn.close()
    if row:
        return {
            "harga": row['harga_satuan'],
            "penyedia": row['penyedia'],
            "satuan": row['satuan']
        }
    return None

def format_rupiah(val):
    return f"Rp {val:,.0f}".replace(",", ".")

# Ensure DB initialized
init_db()

# Page Setup
st.set_page_config(page_title="Generator BAST BNPB", page_icon="📋", layout="wide")

st.title("📋 Generator BAST & Database Terpusat BNPB")
st.caption("Sistem Otomasi Berita Acara Serah Terima Bantuan Logistik & Peralatan — System Builder Han XVI")

# Sidebar Menu
menu = st.sidebar.radio("Navigasi Menu", ["Form Input BAST Baru", "Rekap Database & Log", "Panduan System Builder"])

if menu == "Form Input BAST Baru":
    st.header("📝 Formulir Input BAST Dinamis")
    st.markdown("Isi data transaksi BAST di bawah ini. Sistem menyediakan **penomoran otomatis**, **auto-suggest**, dan **quick-fill harga pembanding**.")

    penyedia_hist, barang_hist, bantuan_hist = get_historical_data()

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1. Administrasi & Pihak Kesatu (BNPB)")
        kode_instansi = st.text_input("Kode Instansi Penerima (Singkatan)", value="KODAM-II", help="Contoh: KODAM-II, PROV-SUMSEL, LANUD, OIKN")
        nomor_bast = generate_next_bast_number(kode_instansi)
        st.text_input("Nomor BAST (Generated Otomatis)", value=nomor_bast, disabled=True)
        
        tgl_bast = st.date_input("Tanggal BAST", datetime.date.today())
        hari_map = {0:"Senin", 1:"Selasa", 2:"Rabu", 3:"Kamis", 4:"Jumat", 5:"Sabtu", 6:"Minggu"}
        hari_nama = hari_map[tgl_bast.weekday()]
        st.text_input("Hari", value=hari_nama, disabled=True)

        p1_nama = st.text_input("Nama PIHAK PERTAMA", value="dr. Riswandi, M.M.")
        p1_nip = st.text_input("NIP PIHAK PERTAMA", value="19741226 200604 1 002")
        p1_jabatan = st.text_input("Jabatan PIHAK PERTAMA", value="Plt. Direktur Penanganan Darurat Wilayah II")

    with col2:
        st.subheader("2. Identitas PIHAK KEDUA (Penerima)")
        is_blanko = st.checkbox("📋 Cetak Blanko / Kosongkan Data Penerima (Diisi Tulis Tangan di Lapangan)")

        if is_blanko:
            st.info("ℹ️ **Mode Blanko Aktif:** Data penerima akan dicetak sebagai garis titik-titik (`....................`). Anda dapat meng-update data penerima riil di menu 'Rekap Database & Log' setelah BAST fisik ditandatangani.")
            p2_instansi = st.text_input("Nama Instansi Penerima *", value="Kodam II Sriwijaya")
            p2_alamat = st.text_area("Alamat Instansi", value="Palembang, Sumatera Selatan", height=68)
            p2_nama = "...................................."
            p2_nip = "...................................."
            p2_jabatan = "...................................."
            p2_hp = "...................................."
            status_bast = "PENDING_RECEIVER"
        else:
            p2_nama = st.text_input("Nama Lengkap Penerima *", value="Sri Mulyanto")
            p2_nip = st.text_input("NIP / NRP Penerima", value="28201418001272")
            p2_jabatan = st.text_input("Jabatan Penerima", value="Kabid II/Logistik")
            p2_instansi = st.text_input("Nama Instansi Penerima *", value="Kodam II Sriwijaya")
            p2_alamat = st.text_area("Alamat Instansi", value="Palembang, Sumatera Selatan", height=68)
            p2_hp = st.text_input("Nomor HP Penerima", value="0812-3456-7890")
            status_bast = "COMPLETED"
        
        jenis_bantuan_input = st.selectbox(
            "Jenis Bantuan / Kategori Logistik", 
            options=["-- Ketik / Pilih Baru --"] + bantuan_hist,
            index=1 if bantuan_hist else 0
        )
        if jenis_bantuan_input == "-- Ketik / Pilih Baru --":
            jenis_bantuan = st.text_input("Masukkan Jenis Bantuan Baru", value="Peralatan Mitigasi Karhutla")
        else:
            jenis_bantuan = jenis_bantuan_input

    st.markdown("---")
    st.subheader("3. Rincian Barang Bantuan Logistik & Peralatan")

    if 'items_data' not in st.session_state:
        st.session_state.items_data = [
            {"nama": "HT Telox TE320K + Akun Pocstar Permanen", "penyedia": "PT Pandhega Jaya Perkasa", "qty": 50, "satuan": "Unit", "harga": 7035000.0}
        ]

    def add_item():
        st.session_state.items_data.append({"nama": "", "penyedia": "", "qty": 1, "satuan": "Unit", "harga": 0.0})

    def remove_item(idx):
        if len(st.session_state.items_data) > 1:
            st.session_state.items_data.pop(idx)

    st.button("➕ Tambah Baris Barang", on_click=add_item)

    updated_items = []
    grand_total = 0.0

    for idx, item in enumerate(st.session_state.items_data):
        st.markdown(f"**Baris #{idx+1}**")
        c1, c2, c3, c4, c5, c6 = st.columns([3, 3, 1, 1, 2, 1])

        with c1:
            nama_val = st.text_input(f"Nama Barang #{idx+1}", value=item["nama"], key=f"nama_{idx}")
        with c2:
            penyedia_val = st.text_input(f"Penyedia #{idx+1}", value=item["penyedia"], key=f"penyedia_{idx}")
        with c3:
            qty_val = st.number_input(f"Qty #{idx+1}", min_value=1, value=int(item["qty"]), key=f"qty_{idx}")
        with c4:
            satuan_val = st.text_input(f"Satuan #{idx+1}", value=item["satuan"], key=f"satuan_{idx}")
        with c5:
            hist = get_item_historical_price(nama_val) if nama_val else None
            default_harga = float(item["harga"])
            if hist and default_harga == 0.0:
                default_harga = float(hist["harga"])
            
            harga_val = st.number_input(f"Harga Satuan (Rp) #{idx+1}", min_value=0.0, value=default_harga, step=1000.0, key=f"harga_{idx}")

            if hist:
                st.caption(f"💡 *Acuan Historis:* {format_rupiah(hist['harga'])} ({hist['penyedia']})")
                if st.button(f"⚡ Quick-Fill ({format_rupiah(hist['harga'])})", key=f"qfill_{idx}"):
                    st.session_state[f"harga_{idx}"] = float(hist["harga"])
                    if not penyedia_val:
                        st.session_state[f"penyedia_{idx}"] = hist["penyedia"]
                    st.rerun()

        with c6:
            subtotal = qty_val * harga_val
            st.markdown(f"**Subtotal:**<br>`{format_rupiah(subtotal)}`", unsafe_allow_html=True)
            if len(st.session_state.items_data) > 1:
                st.button("🗑️ Hapus", key=f"del_{idx}", on_click=remove_item, args=(idx,))

        grand_total += subtotal
        updated_items.append({
            "nama": nama_val,
            "penyedia": penyedia_val,
            "qty": qty_val,
            "satuan": satuan_val,
            "harga": harga_val,
            "subtotal": subtotal
        })

    st.markdown("---")
    st.markdown(f"### 💰 **Total Valuasi BAST:** `{format_rupiah(grand_total)}`")

    if st.button("🚀 Simpan BAST ke Database & Generate Dokumen", type="primary"):
        if not p2_instansi:
            st.error("Mohon lengkapi Nama Instansi Penerima!")
        else:
            conn = get_connection()
            cursor = conn.cursor()
            
            cursor.execute('''
                INSERT INTO bast_records (
                    nomor_bast, tanggal_bast, hari,
                    pihak_kesatu_nama, pihak_kesatu_nip, pihak_kesatu_jabatan,
                    pihak_kedua_nama, pihak_kedua_nip, pihak_kedua_jabatan,
                    pihak_kedua_instansi, pihak_kedua_alamat, pihak_kedua_hp,
                    jenis_bantuan, total_nilai, status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                nomor_bast, str(tgl_bast), hari_nama,
                p1_nama, p1_nip, p1_jabatan,
                p2_nama, p2_nip, p2_jabatan,
                p2_instansi, p2_alamat, p2_hp,
                jenis_bantuan, grand_total, status_bast
            ))
            
            bast_id = cursor.lastrowid

            for it in updated_items:
                cursor.execute('''
                    INSERT INTO bast_items (bast_id, nama_barang, penyedia, qty, satuan, harga_satuan, total_harga)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (bast_id, it["nama"], it["penyedia"], it["qty"], it["satuan"], it["harga"], it["subtotal"]))

            log_desc = f"BAST Blanko dibuat untuk {p2_instansi}" if is_blanko else f"BAST dibuat untuk {p2_instansi} ({p2_nama})"
            cursor.execute('''
                INSERT INTO activity_log (nomor_bast, action, details)
                VALUES (?, 'CREATE', ?)
            ''', (nomor_bast, f"{log_desc} total {format_rupiah(grand_total)}"))

            conn.commit()
            conn.close()

            # Construct Payload for docx generation
            payload = {
                "meta": {
                    "nomor_bast": nomor_bast,
                    "jenis_dokumen": "BERITA ACARA SERAH TERIMA SEMENTARA PENYALURAN BANTUAN LOGISTIK DAN PERALATAN"
                },
                "tanggal_pelaksanaan": {
                    "hari": hari_nama,
                    "tanggal_angka": tgl_bast.day,
                    "tanggal_terbilang": str(tgl_bast.day),
                    "bulan": str(tgl_bast.month),
                    "tahun_angka": tgl_bast.year,
                    "tahun_terbilang": str(tgl_bast.year)
                },
                "pihak_pertama": {
                    "nama": p1_nama,
                    "nip": p1_nip,
                    "jabatan": p1_jabatan,
                    "alamat": "Gedung Graha BNPB, Jl. Pramuka Kav.38 Jakarta Timur 13120"
                },
                "pihak_kedua": {
                    "nama": p2_nama,
                    "nip_nrp": p2_nip,
                    "jabatan": p2_jabatan,
                    "alamat": p2_alamat,
                    "no_hp": p2_hp,
                    "instansi_penerima": p2_instansi
                },
                "konteks_penyerahan": {
                    "lokasi_bencana": p2_instansi,
                    "sumber_dana": "Dana Siap Pakai"
                },
                "logistik_dan_peralatan": {
                    "daftar_barang": [
                        {
                            "no": i + 1,
                            "jenis_barang": it["nama"],
                            "penyedia": it["penyedia"],
                            "jumlah": int(it["qty"]),
                            "satuan": it["satuan"],
                            "harga_satuan": float(it["harga"]),
                            "total_harga": float(it["subtotal"])
                        } for i, it in enumerate(updated_items)
                    ],
                    "total_keseluruhan": float(grand_total),
                    "pengirim_ekspedisi": "PT Talenta Sembada Logistik"
                }
            }

            safe_num = nomor_bast.replace('/', '_')
            docx_path = os.path.join(TEMP_OUT_DIR, f"BAST_{safe_num}.docx")
            pdf_path = os.path.join(TEMP_OUT_DIR, f"BAST_{safe_num}.pdf")

            generate_bast_docx(payload, docx_path)
            
            # Convert to PDF safely
            pdf_success = False
            try:
                subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", docx_path, "--outdir", TEMP_OUT_DIR], check=True, timeout=30)
                if os.path.exists(pdf_path):
                    pdf_success = True
            except Exception as e:
                pdf_success = False

            st.session_state['last_docx'] = docx_path
            st.session_state['last_pdf'] = pdf_path if pdf_success else None
            st.session_state['last_nomor'] = nomor_bast

            st.success(f"✅ BAST `{nomor_bast}` berhasil disimpan ke database (Status: `{status_bast}`)!")
            st.balloons()

    if 'last_docx' in st.session_state and os.path.exists(st.session_state['last_docx']):
        st.markdown("---")
        st.subheader(f"📥 Unduh Dokumen BAST ({st.session_state['last_nomor']})")
        cd1, cd2 = st.columns(2)
        with cd1:
            with open(st.session_state['last_docx'], "rb") as f_docx:
                st.download_button("📄 Download BAST (.docx)", data=f_docx, file_name=os.path.basename(st.session_state['last_docx']), mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        with cd2:
            if st.session_state.get('last_pdf') and os.path.exists(st.session_state['last_pdf']):
                with open(st.session_state['last_pdf'], "rb") as f_pdf:
                    st.download_button("🔴 Download BAST (.pdf)", data=f_pdf, file_name=os.path.basename(st.session_state['last_pdf']), mime="application/pdf")

elif menu == "Rekap Database & Log":
    st.header("📊 Rekapitulasi Database & Audit Log")
    
    conn = get_connection()
    df_records = pd.read_sql_query("SELECT * FROM bast_records ORDER BY id DESC", conn)
    df_items = pd.read_sql_query("SELECT * FROM bast_items", conn)
    df_log = pd.read_sql_query("SELECT * FROM activity_log ORDER BY id DESC", conn)
    
    st.subheader("✏️ Update Data Penerima BAST Blanko (Pasca-Serah Terima)")
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM bast_records WHERE status = 'PENDING_RECEIVER' OR pihak_kedua_nama LIKE '%.%' ORDER BY id DESC")
    pending_basts = cursor.fetchall()

    if pending_basts:
        st.info(f"📌 Terdapat {len(pending_basts)} BAST berstatus **BLANKO / PENDING_RECEIVER** yang menunggu update data penerima riil.")
        selected_nomor = st.selectbox("Pilih Nomor BAST Blanko yang Ingin Di-update:", [row['nomor_bast'] for row in pending_basts])
        
        selected_row = [r for r in pending_basts if r['nomor_bast'] == selected_nomor][0]
        st.write(f"**Instansi Penerima:** {selected_row['pihak_kedua_instansi']} | **Tanggal BAST:** {selected_row['tanggal_bast']} | **Total:** {format_rupiah(selected_row['total_nilai'])}")
        
        with st.form("form_update_penerima"):
            new_nama = st.text_input("Nama Lengkap Penerima Riil *")
            new_nip = st.text_input("NIP / NRP Penerima Riil")
            new_jabatan = st.text_input("Jabatan Penerima Riil")
            new_hp = st.text_input("Nomor HP Penerima Riil")
            
            btn_update = st.form_submit_button("💾 Update Data Penerima")
            if btn_update:
                if not new_nama:
                    st.error("Nama Penerima Riil wajib diisi!")
                else:
                    c_up = conn.cursor()
                    c_up.execute('''
                        UPDATE bast_records 
                        SET pihak_kedua_nama = ?, pihak_kedua_nip = ?, pihak_kedua_jabatan = ?, pihak_kedua_hp = ?, status = 'COMPLETED'
                        WHERE nomor_bast = ?
                    ''', (new_nama, new_nip, new_jabatan, new_hp, selected_nomor))
                    
                    c_up.execute('''
                        INSERT INTO activity_log (nomor_bast, action, details)
                        VALUES (?, 'UPDATE_RECEIVER', ?)
                    ''', (selected_nomor, f"Data penerima riil diperbarui: {new_nama} ({new_jabatan})"))
                    
                    conn.commit()
                    st.success(f"✅ Data penerima BAST `{selected_nomor}` berhasil diperbarui!")
                    st.rerun()
    else:
        st.success("🎉 Seluruh BAST di database sudah memiliki data penerima lengkap (Tidak ada BAST blanko pending).")

    conn.close()

    st.markdown("---")
    st.subheader("📥 Ekspor Rekapitulasi Database ke Excel")
    excel_path = os.path.join(TEMP_OUT_DIR, "Rekapitulasi_BAST_Master.xlsx")
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df_records.to_excel(writer, sheet_name="Master BAST Records", index=False)
        df_items.to_excel(writer, sheet_name="Detail Barang Logistik", index=False)
        df_log.to_excel(writer, sheet_name="Audit Activity Log", index=False)

    with open(excel_path, "rb") as f_excel:
        st.download_button(
            label="📊 Download Rekapitulasi Excel (.xlsx)",
            data=f_excel,
            file_name="Rekapitulasi_BAST_Master.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.markdown("---")
    tab1, tab2, tab3 = st.tabs([" Master BAST Records", " Detail Barang Logistik", " Audit Activity Log"])

    with tab1:
        st.subheader("Data Master BAST")
        st.dataframe(df_records, use_container_width=True)
    
    with tab2:
        st.subheader("Detail Barang per BAST")
        st.dataframe(df_items, use_container_width=True)

    with tab3:
        st.subheader("Audit Trail Aktivitas")
        st.dataframe(df_log, use_container_width=True)

elif menu == "Panduan System Builder":
    st.header("🛠️ Arsitektur Sistem BAST - Han XVI")
    st.markdown("""
    ### **Komponen Utama System Builder:**
    1. **Dynamic Form & Auto-Numbering:** Menghitung format nomor urut BAST otomatis secara terstruktur.
    2. **Blanko Mode (Print-First):** Mendukung pencetakan BAST dengan data penerima kosong (`....................`) dan fitur update susulan pasca-serah terima.
    3. **Smart Comparative Price Alert:** Mengambil acuan harga historis dari database SQLite untuk tombol *Quick-Fill*.
    4. **Relational Database Schema:** 
       - `bast_records` (Master BAST + Status Pending/Completed)
       - `bast_items` (Rincian Logistik)
       - `activity_log` (Audit Trail)
    """)
