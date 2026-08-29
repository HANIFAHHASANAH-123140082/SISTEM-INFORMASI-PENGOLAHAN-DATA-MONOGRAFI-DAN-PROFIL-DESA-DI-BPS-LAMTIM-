import re


def bersihkan_teks_tempel(teks):
    """
    Rapikan teks hasil copy-paste dari PDF/Word yang biasanya berantakan:
    - Non-breaking space (kadang ikut kecopy dari PDF) diganti spasi biasa
    - Baris-baris pendek akibat word-wrap di sumber asli digabung jadi satu paragraf
    - Spasi/enter berlebih dirapikan jadi satu spasi
    - Baris kosong ganda (enter 2x) tetap dianggap sebagai batas paragraf beneran

    Return: teks bersih, antar-paragraf dipisah '\n\n' (bisa dipecah lagi kalau perlu).
    """
    if not teks:
        return teks

    teks = teks.replace('\xa0', ' ').replace('\u200b', '')
    teks = teks.replace('\r\n', '\n').replace('\r', '\n')

    # Baris kosong ganda (>=2 kali enter) = batas paragraf yang memang disengaja
    daftar_paragraf = re.split(r'\n\s*\n', teks)

    hasil = []
    for paragraf in daftar_paragraf:
        # Gabungkan semua baris/spasi berlebih di dalam satu paragraf jadi satu spasi
        satu_baris = re.sub(r'\s+', ' ', paragraf).strip()
        if satu_baris:
            hasil.append(satu_baris)

    return '\n\n'.join(hasil)