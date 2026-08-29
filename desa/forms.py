from django import forms
from .models import Desa
from .utils import bersihkan_teks_tempel


class DesaForm(forms.ModelForm):
    class Meta:
        model = Desa
        fields = [
            'nama_desa', 'kecamatan', 'kabupaten', 'tahun_profil',
            'garis_bujur', 'garis_lintang', 'status_wilayah',
            'sejarah_desa', 'nama_kepala_desa',
            'nomor_kode', 'keadaan_data',
            'peta_wilayah', 'foto_kepala_desa',
        ]
        widgets = {
            'sejarah_desa': forms.Textarea(attrs={'rows': 4}),
        }

    def clean_sejarah_desa(self):
        teks = self.cleaned_data.get('sejarah_desa', '')
        return bersihkan_teks_tempel(teks)