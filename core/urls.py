from django.contrib import admin
from django.urls import path
from inventory.views import (
    proses_suara_massal, home, halaman_prediksi, simpan_data_historis, 
    api_kalkulasi_prediksi, halaman_gudang, update_stok_manual, 
    terapkan_prediksi_stok, halaman_login, halaman_register, 
    proses_logout, hapus_riwayat, download_excel_stok, simpan_barang_masuk
)

urlpatterns = [
    path('admin/', admin.site.urls), # Pakai .urls, bukan .register
    path('login/', halaman_login, name='halaman_login'),
    path('register/', halaman_register, name='halaman_register'),
    path('logout/', proses_logout, name='logout'),
    path('proses/', proses_suara_massal),
    path('prediksi/', halaman_prediksi, name='halaman_prediksi'),
    path('prediksi/simpan/', simpan_data_historis, name='simpan_data_historis'),
    path('prediksi/api/', api_kalkulasi_prediksi, name='api_kalkulasi_prediksi'),
    path('prediksi/terapkan/', terapkan_prediksi_stok, name='terapkan_prediksi_stok'),
    path('gudang/', halaman_gudang, name='halaman_gudang'),
    path('gudang/update/', update_stok_manual, name='update_stok_manual'),
    path('gudang/masuk/', simpan_barang_masuk, name='simpan_barang_masuk'),
    path('gudang/download/', download_excel_stok, name='download_excel_stok'),

    path('riwayat/hapus/<int:log_id>/', hapus_riwayat, name='hapus_riwayat'),
    path('', home, name='home'),
]