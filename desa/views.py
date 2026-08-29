from django.shortcuts import render, redirect, get_object_or_404
from .models import Desa, ItemKuesioner, JawabanKuesioner
from .forms import DesaForm
from .kuesioner_parser import parse_kuesioner_file

def daftar_desa(request):
    semua_desa = Desa.objects.all().order_by('nama_desa')
    return render(request, 'desa/daftar_desa.html', {'semua_desa': semua_desa})


def tambah_desa(request):
    if request.method == 'POST':
        form = DesaForm(request.POST, request.FILES)
        if form.is_valid():
            desa_baru = form.save()
            return redirect('input_jawaban', desa_id=desa_baru.id)
    else:
        form = DesaForm()
    return render(request, 'desa/tambah_desa.html', {'form': form})


def input_jawaban(request, desa_id):
    desa = get_object_or_404(Desa, id=desa_id)
    semua_item = ItemKuesioner.objects.all()

    if request.method == 'POST':
        jawaban_dict = {}
        for item in semua_item:
            nilai = request.POST.get(f'item_{item.id}', '').strip()
            jawaban_dict[item.no_item] = nilai

        for item in semua_item:
            nilai = jawaban_dict[item.no_item]
            JawabanKuesioner.objects.update_or_create(
                desa=desa, item=item,
                defaults={'nilai': nilai if nilai else None}
            )

        return redirect('daftar_desa')

    jawaban_tersimpan = {j.item_id: j.nilai for j in JawabanKuesioner.objects.filter(desa=desa)}
    daftar_item_jawaban = [
        {'item': item, 'nilai': jawaban_tersimpan.get(item.id, '')}
        for item in semua_item
    ]
    return render(request, 'desa/input_jawaban.html', {
        'desa': desa,
        'daftar_item_jawaban': daftar_item_jawaban,
    })

def edit_desa(request, desa_id):
    desa = get_object_or_404(Desa, id=desa_id)
    if request.method == 'POST':
        form = DesaForm(request.POST, request.FILES, instance=desa)
        if form.is_valid():
            form.save()
            return redirect('daftar_desa')
    else:
        form = DesaForm(instance=desa)
    return render(request, 'desa/edit_desa.html', {'form': form, 'desa': desa})

def hapus_desa(request, desa_id):
    desa = get_object_or_404(Desa, id=desa_id)
    if request.method == 'POST':
        desa.delete()
    return redirect('daftar_desa')

def upload_kuesioner(request, desa_id):
    desa = get_object_or_404(Desa, id=desa_id)
    error = None
    jumlah_terisi = 0

    if request.method == 'POST' and request.FILES.get('file_kuesioner'):
        file_upload = request.FILES['file_kuesioner']
        hasil_parse, error = parse_kuesioner_file(file_upload, file_upload.name)

        if not error:
            if not hasil_parse:
                error = "Tidak ada data yang berhasil dibaca dari file ini. Pastikan formatnya sesuai (lihat petunjuk di halaman ini), atau isi manual."
            else:
                for no_item, jawaban in hasil_parse.items():
                    try:
                        item = ItemKuesioner.objects.get(no_item=no_item)
                    except ItemKuesioner.DoesNotExist:
                        continue
                    JawabanKuesioner.objects.update_or_create(
                        desa=desa, item=item,
                        defaults={'nilai': jawaban}
                    )
                    jumlah_terisi += 1

                if jumlah_terisi > 0:
                    return redirect('input_jawaban', desa_id=desa.id)

    return render(request, 'desa/upload_kuesioner.html', {
        'desa': desa,
        'error': error,
        'jumlah_terisi': jumlah_terisi,
    })