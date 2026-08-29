import docx
import pdfplumber


def _cocokkan_baris_tabel(rows):
    """
    Ambil baris tabel yang kolom pertamanya angka (nomor item),
    dan kolom terakhirnya dianggap sebagai jawaban.
    """
    hasil = {}
    for row in rows:
        if not row or len(row) < 2:
            continue
        sel_pertama = (row[0] or '').strip()
        if not sel_pertama.isdigit():
            continue
        no_item = int(sel_pertama)
        jawaban = (row[-1] or '').strip()
        if jawaban:
            hasil[no_item] = jawaban
    return hasil


def parse_docx(file_obj):
    """Baca semua tabel dalam file .docx, cocokkan format No | ... | Jawaban."""
    document = docx.Document(file_obj)
    hasil = {}
    for table in document.tables:
        rows = []
        for row in table.rows:
            rows.append([cell.text for cell in row.cells])
        hasil.update(_cocokkan_baris_tabel(rows))
    return hasil


def parse_pdf(file_obj):
    """Coba baca tabel dari PDF (hasil export Word). Tidak mendukung PDF hasil scan/foto."""
    hasil = {}
    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                hasil.update(_cocokkan_baris_tabel(table))
    return hasil


def parse_kuesioner_file(file_obj, nama_file):
    """Deteksi ekstensi file, panggil parser yang sesuai. Return dict {no_item: jawaban}."""
    ekstensi = nama_file.lower().rsplit('.', 1)[-1]
    if ekstensi == 'docx':
        return parse_docx(file_obj), None
    elif ekstensi == 'pdf':
        return parse_pdf(file_obj), None
    else:
        return {}, f"Ekstensi .{ekstensi} tidak didukung. Gunakan .docx atau .pdf."