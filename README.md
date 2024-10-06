# Fill InDesign Table Automaticaly From Excel File

## Introduction

Script untuk mengisi tabel di InDesign secara otomatis dari file Excel. Script ini dimodifikasi dari [https://github.com/gerynastiar/indesign-beta.git](https://github.com/gerynastiar/indesign-beta.git).

## Workflow

1. Karena data dari sipaling pusat gak difilter, maka kita filter dulu fren data yang mau diambil. Taruh sheet kec dan desa ke sheet 1 dan 2 (Untuk membuat per kecamatan).
2. Pastikan jumlah dan susunan kolom cocok dengan templatenya.
3. Jalankan [filter.ipynb](Data/filter.ipynb) untuk mengambil data yang diperlukan (sesuaikan kab yang diinginkan).
4. Jalankan [scipt_generator.ipynb](Syntax/script_generator.ipynb) untuk menghasilkan script (per kecamatan) yang akan dijalankan di InDesign.
