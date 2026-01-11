# Arabic Number to Words Converter (Pro+ Edition)

**Arabic Number to Words Converter Pro+** is a Custom Function (UDF) for Microsoft Excel designed to accurately convert numeric values into Arabic words following proper grammar (*Nahwu*) rules.

## ✨ Key Features

* **Grammatical Accuracy**: Automatically handles gender rules (*'Adad Ma'dud*) where numbers 3-10 use the opposite gender.
* **Irab Support**: Supports suffix changes based on sentence position (*Marfu* `-uun` or *Mansub/Majrur* `-iin`).
* **Writing Styles**: Choose between Classic or Modern styles (specifically for *Mi'ah* spelling).
* **International Currencies**: Supports various currencies including IDR (Rupiah), SAR (Riyal), KWD (Dinar), AED (Dirham), and more.
* **Ordinal Numbers**: Universal function to create sequences like "5th Grade" or "10th Level".
* **Unicode Safe**: Uses Unicode characters to ensure Arabic text displays correctly across devices.

## 🚀 How to Use

### 1. Main Function: `=Arab()`

`=Arab(Number; [Mode]; [Gender]; [Style]; [IsCurrency]; [Irab])`

* **Mode**: 1 = Sentence (Default), 2 = Spell per digit.
* **Gender**: 1 = Muannas (Default), 2 = Muzakkar.
* **Irab**: 1 = Marfu (Default), 2 = Mansub/Majrur.

### 2. Currency Function: `=ArabCurrency()`

`=ArabCurrency(1500; "id")`

*Result: Alfun wa Khamsumi'ati Rubiyyatun*

### 3. Ordinal Function: `=ArabUniversal()`

`=ArabUniversal(5; 15; 2; TRUE)`

*Result: Al-Fashlu Al-Khamis (5th Class)*

---

# Konverter Angka ke Terbilang Arab (Edisi Pro+)

**Konverter Angka ke Terbilang Arab Pro+** adalah fungsi kustom (UDF) untuk Microsoft Excel yang dirancang khusus untuk mengubah angka numerik menjadi kalimat terbilang dalam bahasa Arab secara akurat sesuai kaidah tata bahasa (*Nahwu*).

## ✨ Fitur Utama

* **Akurasi Tata Bahasa**: Menangani aturan gender angka (*'Adad Ma'dud*) secara otomatis (angka 3-10 berlawanan jenis).
* **Dukungan Irab**: Mendukung perubahan akhiran kata sesuai posisi kalimat (*Marfu* `-uun` atau *Mansub/Majrur* `-iin`).
* **Gaya Penulisan**: Pilihan antara gaya penulisan klasik atau modern (khusus penulisan *Mi'ah*).
* **Mata Uang Internasional**: Mendukung berbagai mata uang seperti Rupiah (ID), Riyal (SA), Dinar (KW), Dirham (AE), dan lainnya.
* **Angka Urutan (Ordinal)**: Fungsi universal untuk membuat urutan seperti "Kelas ke-5" atau "Tingkat ke-10".
* **Unicode Safe**: Menggunakan karakter Unicode agar teks Arab tetap rapi di berbagai perangkat.

## 🚀 Cara Penggunaan

### 1. Fungsi Utama: `=Arab()`

`=Arab(Angka; [Mode]; [Gender]; [Gaya]; [IsCurrency]; [Irab])`

* **Mode**: 1 = Kalimat (Default), 2 = Eja Per Digit.
* **Gender**: 1 = Muannas (Default), 2 = Muzakkar.
* **Irab**: 1 = Marfu (Default), 2 = Mansub/Majrur.

### 2. Fungsi Mata Uang: `=ArabCurrency()`

`=ArabCurrency(1500; "id")`

*Hasil: Alfun wa Khamsumi'ati Rubiyyatun*

### 3. Fungsi Urutan: `=ArabUniversal()`

`=ArabUniversal(5; 15; 2; TRUE)`

*Hasil: Al-Fashlu Al-Khamis (Kelas ke-5)*

---

## 🛠 Installation / Instalasi

1. Open Excel and press `ALT + F11` / Buka Excel dan tekan `ALT + F11`.
2. Select `Insert` > `Module` / Pilih `Insert` > `Module`.
3. Copy and paste the code from `Module1.bas` into the module / Salin dan tempel kode dari `Module1.bas` ke dalam modul.
4. Save as **Excel Macro-Enabled Workbook (.xlsm)** / Simpan sebagai **Excel Macro-Enabled Workbook (.xlsm)**.

## 📜 License / Lisensi (MIT)

This software is free to use under the MIT License / Perangkat lunak ini tersedia gratis di bawah Lisensi MIT.

**Author / Penulis:** Rida Rahman DH 96-02

**Blog:** [ridahaja.blogspot.com](https://ridahaja.blogspot.com)

**GitHub:** [github.com/RidhaHaja](https://github.com/RidhaHaja)

---

Apakah Anda ingin saya membantu menambahkan **Tabel Parameter** yang lebih detail ke dalam draf dua bahasa ini?
