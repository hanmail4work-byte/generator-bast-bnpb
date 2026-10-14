import json
import os
import tempfile
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

def set_cell_background(cell, fill_hex):
    """Mengatur warna latar belakang sel tabel."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Mengatur padding sel tabel."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def format_rupiah(amount):
    """Format angka menjadi format Rupiah Indonesia."""
    return f"Rp {amount:,.0f}".replace(",", ".")

def generate_bast_docx(data_json, output_filepath):
    """
    Fungsi utama untuk menggenerate dokumen BAST .docx dari payload JSON.
    """
    doc = Document()

    # Page Setup (A4, Normal Margins 2.54 cm)
    for section in doc.sections:
        section.page_width = Inches(8.27)   # A4 Width
        section.page_height = Inches(11.69) # A4 Height
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Base Styles Configuration
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Arial'
    normal_style.font.size = Pt(10)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)

    # 1. KOP SURAT
    p_kop1 = doc.add_paragraph()
    p_kop1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kop1.paragraph_format.space_after = Pt(2)
    run_kop1 = p_kop1.add_run("BADAN NASIONAL PENANGGULANGAN BENCANA")
    run_kop1.bold = True
    run_kop1.font.size = Pt(13)

    p_kop2 = doc.add_paragraph()
    p_kop2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kop2.paragraph_format.space_after = Pt(12)
    run_kop2 = p_kop2.add_run("Gedung Graha BNPB, Jl. Pramuka Raya Kav. 38, Jakarta Timur 13120\nTelp. (021) 29827444 | Fax. (021) 21281200 | www.bnpb.go.id")
    run_kop2.font.size = Pt(8.5)
    run_kop2.font.color.rgb = RGBColor(80, 80, 80)

    # Garis Pembatas Kop Surat
    p_line = doc.add_paragraph()
    p_line.paragraph_format.space_after = Pt(14)
    p_line_run = p_line.add_run("―" * 58)
    p_line_run.bold = True
    p_line_run.font.size = Pt(10)
    p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 2. JUDUL DOKUMEN & NOMOR
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(2)
    run_title = p_title.add_run(data_json['meta']['jenis_dokumen'].upper())
    run_title.bold = True
    run_title.font.size = Pt(11)

    p_nomor = doc.add_paragraph()
    p_nomor.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_nomor.paragraph_format.space_after = Pt(14)
    run_nomor = p_nomor.add_run(f"NOMOR: {data_json['meta']['nomor_bast']}")
    run_nomor.bold = True
    run_nomor.font.size = Pt(10)

    # 3. PARAGRAF PEMBUKA WAKTU
    tgl = data_json['tanggal_pelaksanaan']
    p_pembuka = doc.add_paragraph()
    p_pembuka.paragraph_format.space_after = Pt(8)
    p_pembuka.paragraph_format.line_spacing = 1.15
    p_pembuka.paragraph_format.first_line_indent = Inches(0.3)
    p_pembuka.add_run(
        f"Pada hari ini {tgl['hari']}, tanggal {tgl['tanggal_terbilang']} bulan {tgl['bulan']} "
        f"tahun {tgl['tahun_terbilang']} ({tgl['tanggal_angka']:02d}-{tgl['bulan']}-{tgl['tahun_angka']}), "
        f"kami yang bertanda tangan di bawah ini:"
    )

    # 4. IDENTITAS PIHAK PERTAMA
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_after = Pt(6)
    p_p1.paragraph_format.left_indent = Inches(0.2)
    p_p1.add_run("1. ").bold = True
    p_p1.add_run("PIHAK PERTAMA:\n").bold = True

    p1_details = [
        ("Nama", data_json['pihak_pertama']['nama']),
        ("NIP", data_json['pihak_pertama']['nip']),
        ("Jabatan", data_json['pihak_pertama']['jabatan']),
        ("Alamat", data_json['pihak_pertama']['alamat']),
    ]
    for label, val in p1_details:
        p_det = doc.add_paragraph()
        p_det.paragraph_format.left_indent = Inches(0.5)
        p_det.paragraph_format.space_after = Pt(2)
        r_lbl = p_det.add_run(f"{label:<12}: ")
        r_lbl.font.name = 'Courier New' # align spacing
        r_lbl.font.size = Pt(9.5)
        r_val = p_det.add_run(val)
        r_val.font.size = Pt(9.5)

    p_p1_note = doc.add_paragraph()
    p_p1_note.paragraph_format.left_indent = Inches(0.5)
    p_p1_note.paragraph_format.space_after = Pt(8)
    p_p1_note.add_run("Dalam hal ini bertindak untuk dan atas nama ").font.size = Pt(9.5)
    p_p1_note.add_run("Badan Nasional Penanggulangan Bencana").bold = True
    p_p1_note.add_run(", selanjutnya disebut ").font.size = Pt(9.5)
    p_p1_note.add_run("PIHAK PERTAMA").bold = True
    p_p1_note.add_run(".")

    # 5. IDENTITAS PIHAK KEDUA
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_after = Pt(6)
    p_p2.paragraph_format.left_indent = Inches(0.2)
    p_p2.add_run("2. ").bold = True
    p_p2.add_run("PIHAK KEDUA:\n").bold = True

    p2_details = [
        ("Nama", data_json['pihak_kedua']['nama']),
        ("NIP/NRP", data_json['pihak_kedua']['nip_nrp']),
        ("Jabatan", data_json['pihak_kedua']['jabatan']),
        ("Alamat", data_json['pihak_kedua']['alamat']),
        ("No. HP", data_json['pihak_kedua']['no_hp']),
    ]
    for label, val in p2_details:
        p_det = doc.add_paragraph()
        p_det.paragraph_format.left_indent = Inches(0.5)
        p_det.paragraph_format.space_after = Pt(2)
        r_lbl = p_det.add_run(f"{label:<12}: ")
        r_lbl.font.name = 'Courier New'
        r_lbl.font.size = Pt(9.5)
        r_val = p_det.add_run(val)
        r_val.font.size = Pt(9.5)

    p_p2_note = doc.add_paragraph()
    p_p2_note.paragraph_format.left_indent = Inches(0.5)
    p_p2_note.paragraph_format.space_after = Pt(10)
    p_p2_note.add_run("Dalam hal ini bertindak untuk dan atas nama ").font.size = Pt(9.5)
    p_p2_note.add_run(data_json['pihak_kedua']['instansi_penerima']).bold = True
    p_p2_note.add_run(", selanjutnya disebut ").font.size = Pt(9.5)
    p_p2_note.add_run("PIHAK KEDUA").bold = True
    p_p2_note.add_run(".")

    # 6. KLAUSUL PENYERAHAN & TABEL BARANG
    p_klausul = doc.add_paragraph()
    p_klausul.paragraph_format.space_after = Pt(8)
    p_klausul.paragraph_format.line_spacing = 1.15
    p_klausul.add_run(
        f"PIHAK PERTAMA menyerahkan Barang Bantuan Logistik dan Peralatan Penanggulangan Bencana kepada "
        f"PIHAK KEDUA untuk penanganan bencana di wilayah {data_json['konteks_penyerahan']['lokasi_bencana']}, "
        f"dan PIHAK KEDUA menyatakan telah menerima dari PIHAK PERTAMA berupa barang bersumber dari "
        f"{data_json['konteks_penyerahan']['sumber_dana']} dengan rincian sebagai berikut:"
    )

    # TABEL BARANG
    barang_list = data_json['logistik_dan_peralatan']['daftar_barang']
    table = doc.add_table(rows=1, cols=7)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    col_widths = [Inches(0.4), Inches(1.8), Inches(1.3), Inches(0.6), Inches(0.5), Inches(0.8), Inches(0.87)]

    hdr_cells = table.rows[0].cells
    hdr_titles = ["No", "Jenis Barang", "Penyedia", "Jumlah", "Satuan", "Harga Satuan", "Total Harga"]
    
    for i, title in enumerate(hdr_titles):
        hdr_cells[i].width = col_widths[i]
        set_cell_background(hdr_cells[i], "1F4E78") # Deep Blue
        set_cell_margins(hdr_cells[i], top=120, bottom=120, left=100, right=100)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(title)
        run.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(8.5)

    for item in barang_list:
        row_cells = table.add_row().cells
        row_data = [
            str(item['no']),
            item['jenis_barang'],
            item['penyedia'],
            f"{item['jumlah']:,}".replace(",", "."),
            item['satuan'],
            format_rupiah(item['harga_satuan']),
            format_rupiah(item['total_harga'])
        ]
        
        for i, val in enumerate(row_data):
            row_cells[i].width = col_widths[i]
            set_cell_margins(row_cells[i], top=80, bottom=80, left=80, right=80)
            p = row_cells[i].paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(8)
            
            if i in [0, 3, 4]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif i in [5, 6]:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    total_row = table.add_row().cells
    total_row[0].width = col_widths[0]
    total_label_cell = total_row[0]
    for i in range(1, 6):
        total_label_cell.merge(total_row[i])
        
    set_cell_background(total_label_cell, "F2F2F2")
    set_cell_margins(total_label_cell, top=100, bottom=100, left=100, right=100)
    p_tot_lbl = total_label_cell.paragraphs[0]
    p_tot_lbl.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_tot_lbl = p_tot_lbl.add_run("TOTAL JUMLAH (Rp): ")
    run_tot_lbl.bold = True
    run_tot_lbl.font.size = Pt(8.5)

    total_val_cell = total_row[6]
    total_val_cell.width = col_widths[6]
    set_cell_background(total_val_cell, "F2F2F2")
    set_cell_margins(total_val_cell, top=100, bottom=100, left=100, right=100)
    p_tot_val = total_val_cell.paragraphs[0]
    p_tot_val.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_tot_val = p_tot_val.add_run(format_rupiah(data_json['logistik_dan_peralatan']['total_keseluruhan']))
    run_tot_val.bold = True
    run_tot_val.font.size = Pt(8.5)

    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>\n'
        f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>\n'
        f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>\n'
        f'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

    # 7. PASAL KETENTUAN DAN KETENTUAN HUKUM
    p_pasal_intro = doc.add_paragraph()
    p_pasal_intro.paragraph_format.space_before = Pt(12)
    p_pasal_intro.paragraph_format.space_after = Pt(4)
    p_pasal_intro.add_run("Selanjutnya penyerahan Barang Bantuan Logistik dan Peralatan tersebut mengikuti ketentuan sebagai berikut:").font.size = Pt(9.5)

    ketentuan_items = [
        "PIHAK KEDUA bertanggung jawab atas pemanfaatan, perawatan, dan pemeliharaan Barang Bantuan Logistik dan Peralatan Bencana tersebut di atas.",
        "Proses Hibah/Pengalihan Aset Barang Bantuan Logistik dan Peralatan Bencana dari PIHAK PERTAMA kepada PIHAK KEDUA dilaksanakan sesuai ketentuan peraturan perundang-undangan:",
        "a. PIHAK KEDUA wajib menyampaikan Laporan Pemanfaatan Barang Bantuan Logistik dan Peralatan Bencana secara berkala kepada PIHAK PERTAMA.",
        "b. PIHAK KEDUA bersedia mencatat Barang Bantuan Logistik dan Peralatan Bencana sebagai Barang Milik Daerah (BMD) setelah Naskah Hibah ditandatangan.",
        "c. Segala biaya perawatan, operasional, dan pemeliharaan sejak ditandatanganinya Berita Acara Serah Terima ini menjadi tanggung jawab PIHAK KEDUA.",
        "d. Pengalihan hak milik akan ditindaklanjuti dengan penerbitan Naskah Hibah Barang Milik Negara (BMN) dari BNPB.",
        "e. Barang bantuan ini dipergunakan sepenuhnya untuk kepentingan penanggulangan bencana dan masyarakat yang terdampak.",
        "f. Pengirim / Ekspedisi penyalur logistik dan peralatan ini dilaksanakan oleh " + data_json['logistik_dan_peralatan']['pengirim_ekspedisi'] + "."
    ]

    for idx, item_text in enumerate(ketentuan_items):
        p_k = doc.add_paragraph()
        p_k.paragraph_format.space_after = Pt(3)
        p_k.paragraph_format.line_spacing = 1.15
        
        if idx == 0:
            p_k.paragraph_format.left_indent = Inches(0.2)
            p_k.add_run("1. ").bold = True
            p_k.add_run(item_text).font.size = Pt(9)
        elif idx == 1:
            p_k.paragraph_format.left_indent = Inches(0.2)
            p_k.add_run("2. ").bold = True
            p_k.add_run(item_text).font.size = Pt(9)
        else:
            p_k.paragraph_format.left_indent = Inches(0.5)
            p_k.add_run(item_text).font.size = Pt(8.5)

    # 8. PARAGRAF PENUTUP
    p_penutup = doc.add_paragraph()
    p_penutup.paragraph_format.space_before = Pt(8)
    p_penutup.paragraph_format.space_after = Pt(16)
    p_penutup.paragraph_format.line_spacing = 1.15
    p_penutup.paragraph_format.first_line_indent = Inches(0.3)
    p_penutup.add_run(
        "Demikian Berita Acara Serah Terima ini dibuat dalam rangkap secukupnya untuk dipergunakan sebagaimana mestinya."
    ).font.size = Pt(9.5)

    # 9. TANDA TANGAN (2 KOLOM BORDERLESS TABLE)
    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_table.autofit = False

    sig_col_widths = [Inches(3.1), Inches(3.1)]
    sig_cells = sig_table.rows[0].cells

    # PIHAK KEDUA (Kiri)
    cell_kiri = sig_cells[0]
    cell_kiri.width = sig_col_widths[0]
    p_kiri = cell_kiri.paragraphs[0]
    p_kiri.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kiri.paragraph_format.space_after = Pt(2)
    p_kiri.add_run("PIHAK KEDUA\n").bold = True
    p_kiri.add_run(f"{data_json['pihak_kedua']['instansi_penerima']}\n\n\n\n\n").font.size = Pt(9)
    p_kiri.add_run(f"{data_json['pihak_kedua']['nama']}\n").bold = True
    p_kiri.add_run(f"NIP/NRP. {data_json['pihak_kedua']['nip_nrp']}").font.size = Pt(8.5)

    # PIHAK PERTAMA (Kanan)
    cell_kanan = sig_cells[1]
    cell_kanan.width = sig_col_widths[1]
    p_kanan = cell_kanan.paragraphs[0]
    p_kanan.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kanan.paragraph_format.space_after = Pt(2)
    p_kanan.add_run("PIHAK PERTAMA\n").bold = True
    p_kanan.add_run(f"{data_json['pihak_pertama']['jabatan']}\n\n\n\n\n").font.size = Pt(9)
    p_kanan.add_run(f"{data_json['pihak_pertama']['nama']}\n").bold = True
    p_kanan.add_run(f"NIP. {data_json['pihak_pertama']['nip']}").font.size = Pt(8.5)

    doc.save(output_filepath)
    print(f"Dokumen BAST berhasil di-generate: {output_filepath}")

if __name__ == "__main__":
    sample_payload = {
        "meta": {
            "nomor_bast": "001/BAST/BNPB/OIKN/X/2026",
            "jenis_dokumen": "BERITA ACARA SERAH TERIMA SEMENTARA PENYALURAN BANTUAN LOGISTIK DAN PERALATAN"
        },
        "tanggal_pelaksanaan": {
            "hari": "Kamis",
            "tanggal_angka": 17,
            "tanggal_terbilang": "Tujuh Belas",
            "bulan": "Agustus",
            "tahun_angka": 2026,
            "tahun_terbilang": "Dua Ribu Dua Puluh Enam"
        },
        "pihak_pertama": {
            "nama": "dr. Riswandi, M.M.",
            "nip": "19741226 200604 1 002",
            "jabatan": "Plt. Direktur Penanganan Darurat Wilayah II",
            "alamat": "Gedung Graha BNPB, Jl. Pramuka Kav.38 Jakarta Timur 13120"
        },
        "pihak_kedua": {
            "nama": "SRI MULYANTO",
            "nip_nrp": "28201418001272",
            "jabatan": "Paban Dya Bekang Slogdam II/Swj",
            "alamat": "Kabantu Bekang Slogdam II/Swj",
            "no_hp": "081272171254",
            "instansi_penerima": "Pemerintah Kodam II Sriwijaya, Provinsi Sumatra Selatan"
        },
        "konteks_penyerahan": {
            "lokasi_bencana": "Kodam II Sriwijaya, Provinsi Sumatra Selatan",
            "sumber_dana": "Dana Siap Pakai"
        },
        "logistik_dan_peralatan": {
            "daftar_barang": [
                {
                    "no": 1,
                    "jenis_barang": "APD",
                    "penyedia": "PT Alfa Pranata Yudha",
                    "jumlah": 250,
                    "satuan": "Set",
                    "harga_satuan": 7012000,
                    "total_harga": 1753000000
                }
            ],
            "total_keseluruhan": 1753000000,
            "pengirim_ekspedisi": "PT Talenta Sembada Logistik"
        }
    }

    test_dir = os.path.join(tempfile.gettempdir(), "bast_scratch")
    os.makedirs(test_dir, exist_ok=True)
    out_path = os.path.join(test_dir, "BAST_Sample_Generated.docx")
    generate_bast_docx(sample_payload, out_path)
