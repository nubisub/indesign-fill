# Fill InDesign Table Automaticaly From Excel File

## Introduction

Script untuk mengisi tabel di InDesign secara otomatis dari file Excel. Script ini dimodifikasi dari [https://github.com/gerynastiar/indesign-beta.git](https://github.com/gerynastiar/indesign-beta.git).

## Workflow

### Filter Data

1. Karena data dari sipaling pusat gak difilter, maka kita filter dulu fren data yang mau dipakai. Jalankan [filter.ipynb](Data/filter.ipynb) untuk mengambil data yang diperlukan (sesuaikan kab/kec yang diinginkan).
2. Pastikan urutan file excel di folder urut sesuai tabel di InDesign.
3. Pastikan jumlah dan susunan kolom sesuai dengan table di indesign.

### Generate Script

- Jalankan [scipt_generator.ipynb](Syntax/script_generator.ipynb) untuk menghasilkan script yang akan dijalankan di InDesign (sesuaikan kab/kec yang diinginkan).

### Run Script

- Jalankan script di indesign `Window > Utilities > Scripts`.
