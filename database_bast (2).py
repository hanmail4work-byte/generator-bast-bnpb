import sqlite3
import datetime
import os
import tempfile

# Cross-platform, permission-safe database path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_NAME = os.path.join(BASE_DIR, "bast_database.db")

def get_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    
    # 1. Tabel Master BAST Records
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

    # 2. Tabel Detail Barang (BAST Items)
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

    # 3. Tabel Activity Log (Audit Trail)
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
    print("Database initialization completed successfully.")

def generate_next_bast_number(kode_instansi="BNPB", bulan_romawi=None, tahun=None):
    if not bulan_romawi or not tahun:
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
    
    nomor_bast = f"{seq_str}/BAST/BNPB/{kode_instansi}/{bulan_romawi}/{tahun}"
    conn.close()
    return nomor_bast

def query_suggest_penyedia(keyword=""):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT DISTINCT penyedia FROM bast_items 
        WHERE penyedia LIKE ? 
        ORDER BY penyedia LIMIT 5
    ''', (f"%{keyword}%",))
    rows = cursor.fetchall()
    conn.close()
    return [row['penyedia'] for row in rows]

def query_suggest_barang(keyword=""):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT DISTINCT nama_barang FROM bast_items 
        WHERE nama_barang LIKE ? 
        ORDER BY nama_barang LIMIT 5
    ''', (f"%{keyword}%",))
    rows = cursor.fetchall()
    conn.close()
    return [row['nama_barang'] for row in rows]

def query_suggest_jenis_bantuan(keyword=""):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT DISTINCT jenis_bantuan FROM bast_records 
        WHERE jenis_bantuan LIKE ? 
        ORDER BY jenis_bantuan LIMIT 5
    ''', (f"%{keyword}%",))
    rows = cursor.fetchall()
    conn.close()
    return [row['jenis_bantuan'] for row in rows]

def get_historical_price(nama_barang):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT harga_satuan, penyedia, strftime('%Y-%m-%d', i.id) 
        FROM bast_items i
        WHERE LOWER(nama_barang) = LOWER(?)
        ORDER BY i.id DESC LIMIT 1
    ''', (nama_barang,))
    row = cursor.fetchone()
    conn.close()
    if row:
        return {
            "harga_terakhir": row['harga_satuan'],
            "penyedia_terakhir": row['penyedia']
        }
    return None

def insert_sample_data():
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) as count FROM bast_records")
    if cursor.fetchone()['count'] > 0:
        conn.close()
        return

    sample_basts = [
        {
            "nomor": "001/BAST/BNPB/KODAM-II/IX/2026",
            "tanggal": "2026-09-17",
            "hari": "Kamis",
            "p2_nama": "Sri Mulyanto",
            "p2_nip": "28201418001272",
            "p2_jabatan": "Kabid II/Logistik",
            "p2_instansi": "Kodam II Sriwijaya",
            "p2_alamat": "Palembang, Sumatera Selatan",
            "p2_hp": "0812-3456-7890",
            "jenis": "Peralatan Mitigasi Karhutla",
            "total": 3328000000.0,
            "items": [
                ("APD Set Pemadam Karhutla", "PT Pandhega Jaya Perkasa", 250, "Set", 7012000.0, 1753000000.0),
                ("Tabung Motor Pemadam Kebakaran", "PT Pandhega Jaya Perkasa", 50, "Set", 15000000.0, 750000000.0),
                ("Foam Pemadam Kebakaran", "PT Pandhega Jaya Perkasa", 7000, "Liter", 80000.0, 560000000.0),
                ("Gel Pemadam Api", "PT Pandhega Jaya Perkasa", 1750, "Liter", 151428.0, 265000000.0)
            ]
        },
        {
            "nomor": "002/BAST/BNPB/PROV-SUMSEL/IX/2026",
            "tanggal": "2026-09-17",
            "hari": "Kamis",
            "p2_nama": "Mirhansa",
            "p2_nip": "197705072009011007",
            "p2_jabatan": "Staf Bidang Darurat",
            "p2_instansi": "Pemerintah Provinsi Sumatera Selatan",
            "p2_alamat": "Jln. Arah Bandara SMB II Palembang",
            "p2_hp": "0823-7740-1709",
            "jenis": "Peralatan Komunikasi Bencana",
            "total": 399369000.0,
            "items": [
                ("HT Telox TE320K + Akun Pocstar Permanen", "PT Andalas Megatama Persada", 50, "Unit", 7035000.0, 351750000.0),
                ("Kartu Perdana 1 Tahun", "PT Andalas Megatama Persada", 50, "Unit", 952380.0, 47619000.0)
            ]
        }
    ]

    for b in sample_basts:
        cursor.execute('''
            INSERT INTO bast_records (
                nomor_bast, tanggal_bast, hari, pihak_kedua_nama, pihak_kedua_nip,
                pihak_kedua_jabatan, pihak_kedua_instansi, pihak_kedua_alamat, pihak_kedua_hp,
                jenis_bantuan, total_nilai
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (b["nomor"], b["tanggal"], b["hari"], b["p2_nama"], b["p2_nip"],
              b["p2_jabatan"], b["p2_instansi"], b["p2_alamat"], b["p2_hp"],
              b["jenis"], b["total"]))
        
        bast_id = cursor.lastrowid

        for item in b["items"]:
            cursor.execute('''
                INSERT INTO bast_items (
                    bast_id, nama_barang, penyedia, qty, satuan, harga_satuan, total_harga
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (bast_id, item[0], item[1], item[2], item[3], item[4], item[5]))

        cursor.execute('''
            INSERT INTO activity_log (nomor_bast, action, details)
            VALUES (?, 'CREATE', 'Data historis diawal dari BAST fisik')
        ''', (b["nomor"],))

    conn.commit()
    conn.close()
    print("Sample historical data seeded.")

if __name__ == "__main__":
    init_db()
    insert_sample_data()
    
    print("\n--- TEST QUERY HISTORIS ---")
    print("Auto-suggest Penyedia 'Andalas':", query_suggest_penyedia("Andalas"))
    print("Auto-suggest Barang 'HT':", query_suggest_barang("HT"))
    print("Historical Price 'HT Telox TE320K + Akun Pocstar Permanen':", get_historical_price("HT Telox TE320K + Akun Pocstar Permanen"))
    print("Generated Next Number:", generate_next_bast_number("LANUD"))
