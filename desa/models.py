from django.db import models
from django.contrib.auth.models import User


class Bab(models.Model):
    kode = models.CharField(max_length=2, unique=True)  # contoh: "A", "B", ..., "K"
    nama_bab = models.CharField(max_length=100)
    urutan = models.IntegerField()

    class Meta:
        ordering = ['urutan']

    def __str__(self):
        return f"{self.kode}. {self.nama_bab}"


class SubbabTabel(models.Model):
    bab = models.ForeignKey(Bab, on_delete=models.CASCADE, related_name='subbab_list')
    kode_tabel = models.CharField(max_length=10)  # contoh: "A.1", "B.2"
    nama_tabel = models.CharField(max_length=150)
    urutan = models.IntegerField()
    label_kategori = models.CharField(max_length=100, default="Uraian", help_text="Judul kolom ke-2, contoh: 'Organisasi', 'Jenis Sarana'")
    label_nilai = models.CharField(max_length=100, default="Jumlah", help_text="Judul kolom ke-3, contoh: 'Jumlah Organisasi'")
    
    class Meta:
        ordering = ['urutan']

    def __str__(self):
        return f"{self.kode_tabel} {self.nama_tabel}"


class ItemKuesioner(models.Model):
    TIPE_DATA_CHOICES = [
        ('angka', 'Angka'),
        ('teks', 'Teks'),
        ('pilihan', 'Pilihan (Ada/Tidak Ada, dll)'),
    ]
    no_item = models.IntegerField(unique=True)
    pertanyaan = models.TextField()
    konsep_definisi = models.TextField(blank=True)
    tipe_data = models.CharField(max_length=20, choices=TIPE_DATA_CHOICES)
    satuan = models.CharField(max_length=50, blank=True)
    pilihan_jawaban = models.CharField(max_length=500, blank=True)  # dipisah koma, contoh: "Ada,Tidak Ada"

    class Meta:
        ordering = ['no_item']

    def __str__(self):
        return f"Item {self.no_item}: {self.pertanyaan[:50]}"

    def daftar_pilihan(self):
        return [p.strip() for p in self.pilihan_jawaban.split(',')] if self.pilihan_jawaban else []

class ItemSubbab(models.Model):
    item = models.ForeignKey(ItemKuesioner, on_delete=models.CASCADE, related_name='pemetaan_subbab')
    subbab = models.ForeignKey(SubbabTabel, on_delete=models.CASCADE, related_name='item_list')
    nama_kolom = models.CharField(max_length=150)  # nama yang tampil di tabel dokumen

    def __str__(self):
        return f"Item {self.item.no_item} -> {self.subbab.kode_tabel}"


class Desa(models.Model):
    nama_desa = models.CharField(max_length=100)
    kecamatan = models.CharField(max_length=100)
    kabupaten = models.CharField(max_length=100)
    tahun_profil = models.IntegerField()
    garis_bujur = models.DecimalField(max_digits=10, decimal_places=6, null=True, blank=True)
    garis_lintang = models.DecimalField(max_digits=10, decimal_places=6, null=True, blank=True)
    status_wilayah = models.CharField(max_length=50, blank=True)  # Desa / Kelurahan
    sejarah_desa = models.TextField(blank=True)
    nama_kepala_desa = models.CharField(max_length=100, blank=True)
    nomor_kode = models.CharField(max_length=20, blank=True, help_text="Nomor kode desa/kelurahan, contoh: 2007")
    keadaan_data = models.CharField(max_length=100, blank=True, help_text="Periode data, contoh: Desember 2024 dan April 2025")
    peta_wilayah = models.ImageField(upload_to='peta_wilayah/', null=True, blank=True)
    foto_kepala_desa = models.ImageField(upload_to='foto_kepala_desa/', null=True, blank=True)

    def __str__(self):
        return self.nama_desa


class JawabanKuesioner(models.Model):
    desa = models.ForeignKey(Desa, on_delete=models.CASCADE, related_name='jawaban_list')
    item = models.ForeignKey(ItemKuesioner, on_delete=models.CASCADE, related_name='jawaban_list')
    user = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    nilai = models.CharField(max_length=255, blank=True, null=True)  # kosong = belum diisi
    tanggal_input = models.DateTimeField(auto_now=True)

    class Meta:
        unique_together = ('desa', 'item')  # 1 desa hanya boleh 1 jawaban per item

    def __str__(self):
        return f"{self.desa.nama_desa} - Item {self.item.no_item}: {self.nilai}"
