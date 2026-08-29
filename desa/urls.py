from django.urls import path
from . import views
from .document_generator import download_profil_desa, download_monografi_desa

urlpatterns = [
    path('', views.daftar_desa, name='daftar_desa'),
    path('tambah/', views.tambah_desa, name='tambah_desa'),
    path('<int:desa_id>/edit/', views.edit_desa, name='edit_desa'),
    path('<int:desa_id>/input-jawaban/', views.input_jawaban, name='input_jawaban'),
    path('<int:desa_id>/upload-kuesioner/', views.upload_kuesioner, name='upload_kuesioner'),
    path('<int:desa_id>/hapus/', views.hapus_desa, name='hapus_desa'),
    path('<int:desa_id>/download-profil/', download_profil_desa, name='download_profil'),
    path('<int:desa_id>/download-monografi/', download_monografi_desa, name='download_monografi'),
]