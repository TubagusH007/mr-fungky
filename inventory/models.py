from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone

class BahanBaku(models.Model):
    ZONA_CHOICES = [('TOPPING', 'Topping'), ('TEA', 'Tea Base'), ('DAIRY', 'Dairy'), ('BAHAN_BAKU', 'Bahan Baku')]
    nama_bahan = models.CharField(max_length=100)
    stok_sekarang = models.FloatField(default=0)
    stok_minimal = models.FloatField(default=50) # Info stok menipis
    satuan = models.CharField(max_length=20, default='gr')
    kategori_zona = models.CharField(max_length=20, choices=ZONA_CHOICES)

    def is_menipis(self):
        return self.stok_sekarang <= self.stok_minimal

    def __str__(self):
        return self.nama_bahan

class RiwayatStok(models.Model):
    petugas = models.ForeignKey(User, on_delete=models.CASCADE)
    bahan = models.ForeignKey(BahanBaku, on_delete=models.CASCADE)
    jumlah_baru = models.FloatField()
    waktu = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-waktu'] # History terbaru di atas

class DataPenjualan(models.Model):
    tanggal = models.DateField(auto_now_add=True)
    jumlah_pembeli = models.IntegerField()
    omset = models.DecimalField(max_digits=12, decimal_places=2)

    def __str__(self):
        return f"Data {self.tanggal}"

    class Meta:
        ordering = ['-tanggal'] # Mengurutkan berdasarkan tanggal terbaru

class DataHistorisHarian(models.Model):
    tanggal = models.DateField(auto_now_add=True)
    jumlah_antrian = models.IntegerField()
    omset = models.DecimalField(max_digits=12, decimal_places=2)
    
    # Bahan Baku Terpakai (Berdasarkan Excel)
    deeproast = models.FloatField(default=0)
    four_season = models.FloatField(default=0)
    jasmine = models.FloatField(default=0)
    boba = models.FloatField(default=0)
    paper_cup_besar = models.IntegerField(default=0)
    paper_cup_kecil = models.IntegerField(default=0)
    gula = models.FloatField(default=0)
    gula_cair = models.FloatField(default=0)
    honey_syrup = models.FloatField(default=0)
    konjac_jelly = models.FloatField(default=0)
    creamer = models.FloatField(default=0)

    def __str__(self):
        return f"Historis {self.tanggal} - Antrian: {self.jumlah_antrian}"

    class Meta:
        ordering = ['-tanggal']