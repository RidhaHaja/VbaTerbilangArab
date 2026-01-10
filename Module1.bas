Attribute VB_Name = "Module1"
Rem VB_Name = "Module1"
Option Explicit

Rem ===============================================================================================================
Rem PROYEK       : Konverter Angka ke Terbilang Arab (Edisi Pro+)
Rem PENULIS      : Rida Rahman DH 96-02
Rem BLOG         : https://ridahaja.blogspot.com
Rem TANGGAL      : Januari 2026
Rem ===============================================================================================================
Rem MIT LICENSE (BILINGUAL):
Rem
Rem [ENGLISH]
Rem Copyright (c) 2026 Rida Rahman DH 96-02
Rem
Rem Permission is hereby granted, free of charge, to any person obtaining a copy of this software
Rem and associated documentation files (the "Software"), to deal in the Software without restriction,
Rem including without limitation the rights to use, copy, modify, merge, publish, distribute,
Rem sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
Rem furnished to do so, subject to the following conditions:
Rem
Rem The above copyright notice and this permission notice shall be included in all copies or
Rem substantial portions of the Software.
Rem
Rem THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
Rem BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
Rem NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
Rem DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
Rem OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
Rem
Rem [BAHASA INDONESIA]
Rem Hak Cipta (c) 2026 Rida R DH 96-02
Rem
Rem Izin diberikan secara gratis kepada siapa pun yang mendapatkan salinan perangkat lunak ini
Rem untuk menggunakan, menyalin, mengubah, dan mendistribusikan tanpa batasan, dengan syarat
Rem pemberitahuan hak cipta di atas wajib disertakan dalam semua salinan perangkat lunak.
Rem Perangkat lunak ini disediakan "APA ADANYA", tanpa jaminan apa pun dari penulis.
Rem ===============================================================================================================

Rem ===============================================================================================================
Rem CARA PENGGUNAAN DI EXCEL (DOKUMENTASI):
Rem ===============================================================================================================
Rem 1. FUNGSI UTAMA: =Arab(Angka; [Mode]; [Gender]; [Gaya]; [IsCurrency]; [Irab])
Rem    - Angka      : Nilai yang akan dikonversi (contoh: 125).
Rem    - Mode       : 1 = Terbilang Kata (Default), 2 = Eja Per Digit.
Rem    - Gender     : 1 = Muannas (Default), 2 = Muzakkar.
Rem    - Gaya       : 1 = Klasik, 2 = Modern (Default).
Rem    - IsCurrency : True = Mode Mata Uang (Koma dieja sebagai satuan), False = Desimal Biasa.
Rem    - Irab       : 1 = Marfu (Default: -uun), 2 = Mansub/Majrur (-iin).
Rem
Rem 2. FUNGSI UNIVERSAL: =ArabUniversal(Angka; [Satuan]; [Gender]; [IsOrdinal])
Rem    - Satuan     : 15 = Fashlun (Kelas), 16 = Martabah (Tingkat).
Rem    - IsOrdinal  : True = Angka Urutan (Pertama, Kedua, dst), False = Angka Biasa.
Rem
Rem CONTOH RUMUS:
Rem =Arab(125; 1; 2)             --> Menghasilkan 125 dalam bentuk Muzakkar.
Rem =ArabUniversal(5; 15; 2; TRUE) --> Menghasilkan "Al-Fashlu Al-Khamis" (Kelas ke-5).
Rem ===============================================================================================================
' --- ENUMERASI ---

Public Enum GenderArab
    Muannas = 1
    Muzakkar = 2
End Enum

Public Enum GayaArab
    Klasik = 1
    Modern = 2
End Enum

Public Enum IrabArab
    Marfu = 1
    MansubMajrur = 2
End Enum

Private Kata(0 To 9, 1 To 2) As String
Private Puluhan(2 To 9) As String
Private SatuanBesar(0 To 10) As String
Private Ordinal(1 To 12, 1 To 2) As String
Private IsInitialized As Boolean

Sub InitializeArabicWords()
    If IsInitialized Then Exit Sub
    Dim i As Integer

    Kata(0, 2) = ChrW(1589) & ChrW(1601) & ChrW(1585) ' Sifrun
    Kata(1, 2) = ChrW(1608) & ChrW(1575) & ChrW(1581) & ChrW(1583) ' Wahid
    Kata(2, 2) = ChrW(1575) & ChrW(1579) & ChrW(1606) & ChrW(1575) & ChrW(1606) ' Ithnan
    Kata(3, 2) = ChrW(1579) & ChrW(1604) & ChrW(1575) & ChrW(1579) ' Thalath
    Kata(4, 2) = ChrW(1571) & ChrW(1585) & ChrW(1576) & ChrW(1593) ' Arba
    Kata(5, 2) = ChrW(1582) & ChrW(1605) & ChrW(1587) ' Khams
    Kata(6, 2) = ChrW(1587) & ChrW(1578) ' Sitt
    Kata(7, 2) = ChrW(1587) & ChrW(1576) & ChrW(1593) ' Sab'a
    Kata(8, 2) = ChrW(1579) & ChrW(1605) & ChrW(1575) & ChrW(1606) ' Thaman
    Kata(9, 2) = ChrW(1578) & ChrW(1587) & ChrW(1593) ' Tis'a

    Kata(0, 1) = Kata(0, 2)
    
    For i = 3 To 9
        Kata(i, 1) = Kata(i, 2) & ChrW(1577)    'Ta Marbutah
    Next i
    
    Kata(1, 1) = Kata(1, 2) & ChrW(1577) ' Wahidah
    Kata(2, 1) = ChrW(1575) & ChrW(1579) & ChrW(1606) & ChrW(1578) & ChrW(1575) & ChrW(1606) ' Ithnatan
    Kata(8, 1) = Kata(8, 2) & ChrW(1610) & ChrW(1577) ' Thamaniyah

    Puluhan(2) = ChrW(1593) & ChrW(1588) & ChrW(1585) & ChrW(1608) & ChrW(1606) ' Isyruun
    For i = 3 To 9
        Puluhan(i) = Kata(i, 2) & ChrW(1608) & ChrW(1606) ' Tambah akhiran uun
    Next i
    
    SatuanBesar(0) = ""
    SatuanBesar(1) = ChrW(1571) & ChrW(1604) & ChrW(1601)
    SatuanBesar(2) = ChrW(1605) & ChrW(1604) & ChrW(1610) & ChrW(1608) & ChrW(1606)
    SatuanBesar(3) = ChrW(1605) & ChrW(1604) & ChrW(1610) & ChrW(1575) & ChrW(1585)
    SatuanBesar(4) = ChrW(1578) & ChrW(1585) & ChrW(1610) & ChrW(1604) & ChrW(1610) & ChrW(1608) & ChrW(1606)
    SatuanBesar(5) = ChrW(1576) & ChrW(1604) & ChrW(1610) & ChrW(1575) & ChrW(1585)
    SatuanBesar(6) = ChrW(1587) & ChrW(1604) & ChrW(1610) & ChrW(1608) & ChrW(1606)
    SatuanBesar(7) = ChrW(1587) & ChrW(1604) & ChrW(1610) & ChrW(1575) & ChrW(1585)
    SatuanBesar(8) = ChrW(1578) & ChrW(1604) & ChrW(1610) & ChrW(1608) & ChrW(1606)
    SatuanBesar(9) = ChrW(1578) & ChrW(1604) & ChrW(1610) & ChrW(1575) & ChrW(1585)
    SatuanBesar(10) = ChrW(1583) & ChrW(1610) & ChrW(1587) & ChrW(1610) & ChrW(1604) & ChrW(1610) & ChrW(1608) & ChrW(1606)

    ' ORDINAL NUMBER (1-12)
    Dim Asyar As String: Asyar = ChrW(1593) & ChrW(1588) & ChrW(1585) ' Asyar
    Dim Asyarah As String: Asyarah = Asyar & ChrW(1577) ' Asyarah
    For i = 3 To 5
        Ordinal(i, 2) = Left(Kata(i, 2), 1) & ChrW(1575) & Mid(Kata(i, 2), 2)
        Ordinal(i, 1) = Ordinal(i, 2) & ChrW(1577)
    Next i
    For i = 7 To 9
        Ordinal(i, 2) = Left(Kata(i, 2), 1) & ChrW(1575) & Mid(Kata(i, 2), 2)
        Ordinal(i, 1) = Ordinal(i, 2) & ChrW(1577)
    Next i

    ' 1: Awwal / Ula
    Ordinal(1, 2) = ChrW(1571) & ChrW(1608) & ChrW(1604)
    Ordinal(1, 1) = ChrW(1571) & ChrW(1608) & ChrW(1604) & ChrW(1609)
    ' 2: Thani / Thaniyah
    Ordinal(2, 2) = ChrW(1579) & ChrW(1575) & ChrW(1606) & ChrW(1610)
    Ordinal(2, 1) = Ordinal(2, 2) & ChrW(1577)
    ' 6: Sadis / Sadisah (Perubahan akar kata dari Sitt)
    Ordinal(6, 2) = ChrW(1587) & ChrW(1575) & ChrW(1583) & ChrW(1587)
    Ordinal(6, 1) = Ordinal(6, 2) & ChrW(1577)
    ' 10: Asyir / Asyirah
    Ordinal(10, 2) = ChrW(1593) & ChrW(1575) & ChrW(1588) & ChrW(1585)
    Ordinal(10, 1) = Ordinal(10, 2) & ChrW(1577)
    ' 11: Hadi Asyar / Hadiyata Asyarah
    Ordinal(11, 2) = ChrW(1581) & ChrW(1575) & ChrW(1583) & ChrW(1610) & " " & Asyar
    Ordinal(11, 1) = ChrW(1581) & ChrW(1575) & ChrW(1583) & ChrW(1610) & ChrW(1577) & " " & Asyarah
    ' 12: Thani Asyar / Thaniyata Asyarah
    Ordinal(12, 2) = Ordinal(2, 2) & " " & Asyar
    Ordinal(12, 1) = ChrW(1579) & ChrW(1575) & ChrW(1606) & ChrW(1610) & ChrW(1577) & " " & Asyarah
    IsInitialized = True
End Sub

Private Function ConvertThreeDigitsArab(ByVal num As Variant, _
                                        ByVal PilihGender As GenderArab, _
                                        ByVal Gaya As GayaArab, _
                                        Optional Irab As IrabArab = Marfu) As String
    
    Dim s As String, SPACE As String, WAW As String, Miah As String
    Dim valNum As Integer: valNum = Val(num)
    
    If valNum = 0 Then Exit Function
    
    SPACE = ChrW(32)
    WAW = SPACE & ChrW(1608) & SPACE
    ' Klasik: Mi'ah dengan Alif | Modern: Mi'ah tanpa Alif
    Miah = IIf(Gaya = Klasik, ChrW(1605) & ChrW(1575) & ChrW(1574) & ChrW(1577), ChrW(1605) & ChrW(1574) & ChrW(1577))

    ' --- 1. BAGIAN RATUSAN (100-900) ---
    If valNum >= 100 Then
        Dim h As Integer: h = Int(valNum / 100)
        Select Case h
            Case 1: s = Miah
            Case 2: s = ChrW(1605) & ChrW(1574) & ChrW(1578) & IIf(Irab = Marfu, ChrW(1575), ChrW(1610)) & ChrW(1606) ' Mi'atani / Mi'ataini
            Case Else: s = Kata(h, Muzakkar) & IIf(Gaya = Klasik, SPACE, "") & Miah
        End Select
        valNum = valNum Mod 100: If valNum > 0 Then s = s & WAW
    End If

    ' --- 2. BAGIAN SATUAN, BELASAN, & PULUHAN ---
    If valNum > 0 Then
        Dim st As Integer: st = valNum Mod 10
        Dim TeksSatuan As String
        Dim GenderTarget As GenderArab
        
        ' Aturan Gender Arab 3-10: Berlawanan dengan benda yang dihitung
        GenderTarget = IIf(valNum >= 3, IIf(PilihGender = Muzakkar, Muannas, Muzakkar), PilihGender)
        
        ' Ambil teks dasar dari array Kata (Indeks disesuaikan jika tunggal atau majemuk)
        TeksSatuan = Kata(IIf(valNum < 10, valNum, st), GenderTarget)
        
        ' Jika angka 8 (Muzakkar) bertemu angka lain (belasan atau puluhan), tambahkan Ya (1610)
        ' Jika tunggal (8 saja), ia tetap Thaman (????) sesuai isi array Kata(8,2)
        If st = 8 And GenderTarget = Muzakkar And valNum > 10 Then
            TeksSatuan = TeksSatuan & ChrW(1610)
        End If

        ' Penentuan Struktur Kalimat
        If valNum < 10 Then
            ' 1 - 9
            s = s & TeksSatuan
            
        ElseIf valNum = 10 Then
            ' 10 (Asyarah / Asyar)
            s = s & IIf(PilihGender = Muzakkar, ChrW(1593) & ChrW(1588) & ChrW(1585) & ChrW(1577), ChrW(1593) & ChrW(1588) & ChrW(1585))
            
        ElseIf valNum >= 11 And valNum <= 12 Then
            ' 11 & 12 (Khusus)
            If valNum = 11 Then
                s = s & IIf(PilihGender = Muzakkar, ChrW(1571) & ChrW(1581) & ChrW(1583) & SPACE & ChrW(1593) & ChrW(1588) & ChrW(1585), _
                                                   ChrW(1573) & ChrW(1581) & ChrW(1583) & ChrW(1609) & SPACE & ChrW(1593) & ChrW(1588) & ChrW(1585) & ChrW(1577))
            Else
                Dim Ithna As String
                Ithna = ChrW(1575) & ChrW(1579) & ChrW(1606) & IIf(PilihGender = Muannas, ChrW(1578), "") & IIf(Irab = Marfu, ChrW(1575), ChrW(1610))
                s = s & Ithna & SPACE & ChrW(1593) & ChrW(1588) & ChrW(1585) & IIf(PilihGender = Muannas, ChrW(1577), "")
            End If
            
        ElseIf valNum >= 13 And valNum <= 19 Then
            ' 13 - 19 (Menggunakan TeksSatuan)
            s = s & TeksSatuan & SPACE & ChrW(1593) & ChrW(1588) & ChrW(1585) & IIf(PilihGender = Muannas, ChrW(1577), "")
            
        Else
            ' 20 - 99
            Dim Pulu As String: Pulu = Puluhan(Int(valNum / 10))
            ' Ubah akhiran uun menjadi iin jika Mansub/Majrur
            If Irab = MansubMajrur Then Pulu = Replace(Pulu, ChrW(1608) & ChrW(1606), ChrW(1610) & ChrW(1606))
            
            If st = 0 Then
                s = s & Pulu
            Else
                ' Format: Satuan + WAW + Puluhan (Contoh: Wahid wa 'Isyrun)
                s = s & TeksSatuan & WAW & Pulu
            End If
        End If
    End If
    
    ConvertThreeDigitsArab = Trim(s)
End Function

Public Function Arab(ByVal angka As Variant, _
              Optional Mode As Byte = 1, _
              Optional gender As GenderArab = Muannas, _
              Optional Gaya As GayaArab = Modern, _
              Optional IsCurrency As Boolean = False, _
              Optional Irab As IrabArab = Marfu) As String
    
    Application.Volatile
    If IsError(angka) Or angka = "" Or Not IsNumeric(angka) Then
        Arab = ""
        Exit Function
    End If
    InitializeArabicWords

    Dim strNum As String: strNum = Replace(CStr(angka), " ", "")
    Dim Prefix As String: Prefix = ""
    If Left(strNum, 1) = "-" Then
        Prefix = ChrW(1587) & ChrW(1575) & ChrW(1604) & ChrW(1576) & " "
        strNum = Mid(strNum, 2)
    End If

    If Mode = 2 Then
        Dim i As Long, res As String
        For i = 1 To Len(strNum)
            If IsNumeric(Mid(strNum, i, 1)) Then
                res = res & Kata(CInt(Mid(strNum, i, 1)), gender) & " "
            End If
        Next i
        Arab = Trim(Prefix & res)
    Else
        ' Bagian Terbilang
        Dim PosKoma As Long: PosKoma = InStr(strNum, ".")
        If PosKoma = 0 Then PosKoma = InStr(strNum, ",")
        
        Dim bulat As String, Desimal As String
        If PosKoma > 0 Then
            bulat = Left(strNum, PosKoma - 1)
            Desimal = Mid(strNum, PosKoma + 1)
        Else
            bulat = strNum
        End If

        Dim FinalBulat As String
        If Trim(Replace(bulat, "0", "")) = "" Then
            FinalBulat = Kata(0, gender)
        Else
            FinalBulat = ProsesBlokRibuan(bulat, gender, Gaya, Irab)
        End If

        ' Desimal
        Dim FinalDes As String
        If Len(Desimal) > 0 And Val(Desimal) > 0 Then
            Dim Pemisah As String
            Pemisah = IIf(IsCurrency, " " & ChrW(1608) & " ", " " & ChrW(1601) & ChrW(1575) & ChrW(1589) & ChrW(1604) & ChrW(1577) & " ")
            FinalDes = Pemisah
            If IsCurrency Then
                FinalDes = FinalDes & ConvertThreeDigitsArab(Val(Desimal), Muannas, Gaya, Irab)
            Else
                Dim j As Integer
                For j = 1 To Len(Desimal)
                    FinalDes = FinalDes & Kata(Val(Mid(Desimal, j, 1)), Muannas) & " "
                Next j
            End If
        End If
        Arab = Trim(Prefix & FinalBulat & FinalDes)
    End If
End Function

Private Function ProsesBlokRibuan(ByVal strNum As String, ByVal PilihGender As GenderArab, ByVal Gaya As GayaArab, Optional Irab As IrabArab = Marfu) As String
    Dim Blocks() As String, blockCount As Integer: blockCount = 0
    Dim res As String, WAW As String: WAW = " " & ChrW(1608) & " "
    
    strNum = Trim(strNum)
    Do While Len(strNum) > 0
        ReDim Preserve Blocks(blockCount)
        Dim take As Integer: take = IIf(Len(strNum) >= 3, 3, Len(strNum))
        Blocks(blockCount) = Right(strNum, take)
        strNum = Left(strNum, Len(strNum) - take)
        blockCount = blockCount + 1
    Loop

    Dim i As Integer
    For i = blockCount - 1 To 0 Step -1
        Dim bVal As Integer: bVal = Val(Blocks(i))
        If bVal > 0 Then
            Dim Cur As String
            If i > UBound(SatuanBesar) Then
                Cur = "[Overflow]"
            Else
                Select Case i
                    Case 0: Cur = ConvertThreeDigitsArab(bVal, PilihGender, Gaya, Irab)
                    Case 1: ' RIBUAN
                        Select Case bVal
                            Case 1: Cur = SatuanBesar(1)
                            Case 2: Cur = ChrW(1571) & ChrW(1604) & ChrW(1601) & IIf(Irab = Marfu, ChrW(1575), ChrW(1610)) & ChrW(1606)
                            Case 3 To 10: Cur = Kata(bVal, Muannas) & " " & ChrW(1570) & ChrW(1604) & ChrW(1575) & ChrW(1601)
                            Case Else: Cur = ConvertThreeDigitsArab(bVal, Muzakkar, Gaya, Irab) & " " & SatuanBesar(1)
                        End Select
                    Case Else: ' JUTAAN - DESILYUN
                        If bVal = 1 Then
                            Cur = SatuanBesar(i)
                        ElseIf bVal = 2 Then
                            Cur = SatuanBesar(i) & IIf(Irab = Marfu, ChrW(1575) & ChrW(1606), ChrW(1610) & ChrW(1606))
                        Else
                            Cur = ConvertThreeDigitsArab(bVal, Muzakkar, Gaya, Irab) & " " & SatuanBesar(i)
                        End If
                End Select
            End If
            res = IIf(res = "", Cur, res & WAW & Cur)
        End If
    Next i
    ProsesBlokRibuan = res
End Function

Public Function ArabCurrency(ByVal angka As Variant, _
                             Optional ByVal KodeNegara As String = "id") As String
    ' Parameter KodeNegara (Gunakan Huruf Kecil):
    ' id (Indonesia), sa (Saudi), kw (Kuwait), ae (UEA), qa (Qatar)
    ' my (Malaysia), sg (Singapura), bn (Brunei), jp (Jepang), cn (China)
    
    Dim HasilUtama As String, StrAngka As String
    Dim BagianBulat As String, BagianDesimal As String
    Dim PosKoma As Long
    Dim TeksMataUang As String, TeksSen As String, WAW As String
    
    ' Karakter Unicode untuk kata hubung "dan" (Waw)
    WAW = " " & ChrW(1608) & " "
    KodeNegara = LCase(Trim(KodeNegara))
    
    Select Case KodeNegara
        Case "id" ' Indonesia: Rupiah & Sen
            TeksMataUang = ChrW(1585) & ChrW(1608) & ChrW(1576) & ChrW(1610) & ChrW(1577)
            TeksSen = ChrW(1587) & ChrW(1606)
            
        Case "sa" ' Saudi Arabia: Riyal & Halalah
            TeksMataUang = ChrW(1585) & ChrW(1610) & ChrW(1575) & ChrW(1604)
            TeksSen = ChrW(1607) & ChrW(1604) & ChrW(1604) & ChrW(1577)
            
        Case "kw" ' Kuwait: Dinar & Fils
            TeksMataUang = ChrW(1583) & ChrW(1610) & ChrW(1606) & ChrW(1575) & ChrW(1585)
            TeksSen = ChrW(1601) & ChrW(1604) & ChrW(1587)
            
        Case "ae" ' Uni Emirat Arab: Dirham & Fils
            TeksMataUang = ChrW(1583) & ChrW(1585) & ChrW(1607) & ChrW(1605)
            TeksSen = ChrW(1601) & ChrW(1604) & ChrW(1587)
            
        Case "qa" ' Qatar: Riyal & Dirham
            TeksMataUang = ChrW(1585) & ChrW(1610) & ChrW(1575) & ChrW(1604)
            TeksSen = ChrW(1583) & ChrW(1585) & ChrW(1607) & ChrW(1605)
            
        Case "my" ' Malaysia: Ringgit & Sen
            TeksMataUang = ChrW(1585) & ChrW(1610) & ChrW(1606) & ChrW(1580) & ChrW(1610) & ChrW(1578)
            TeksSen = ChrW(1587) & ChrW(1606)
            
        Case "sg", "bn" ' Singapura & Brunei: Dollar & Sen
            TeksMataUang = ChrW(1583) & ChrW(1608) & ChrW(1604) & ChrW(1575) & ChrW(1585)
            TeksSen = ChrW(1587) & ChrW(1606)
            
        Case "jp" ' Jepang: Yen
            TeksMataUang = ChrW(1610) & ChrW(1606)
            TeksSen = ""
            
        Case "cn" ' China: Yuan
            TeksMataUang = ChrW(1610) & ChrW(1608) & ChrW(1575) & ChrW(1606)
            TeksSen = ""
            
        Case Else
            TeksMataUang = ""
            TeksSen = ""
    End Select

    ' Logika pemisahan angka bulat dan desimal
    StrAngka = CStr(angka)
    PosKoma = InStr(StrAngka, ",")
    If PosKoma = 0 Then PosKoma = InStr(StrAngka, ".")
    
    If PosKoma > 0 Then
        BagianBulat = Left(StrAngka, PosKoma - 1)
        BagianDesimal = Mid(StrAngka, PosKoma + 1)
    Else
        BagianBulat = StrAngka
        BagianDesimal = ""
    End If
    
    HasilUtama = Arab(BagianBulat, 1, Muannas, Modern, False, Marfu) & " " & TeksMataUang
    
    ' Proses bagian sen (maksimal 2 digit untuk mata uang)
    If Val(BagianDesimal) > 0 And TeksSen <> "" Then
        If Len(BagianDesimal) > 2 Then BagianDesimal = Left(BagianDesimal, 2)
        ' Pastikan desimal diproses sebagai satuan terkecil
        HasilUtama = HasilUtama & WAW & _
                     Arab(BagianDesimal, 1, Muannas, Modern, False, Marfu) & " " & TeksSen
    End If
    
    ArabCurrency = Trim(HasilUtama)
End Function

Public Function ArabUniversal(ByVal angka As Variant, _
                              Optional ByVal JenisSatuan As Integer = 0, _
                              Optional ByVal PilihGender As Integer = 1, _
                              Optional ByVal IsOrdinal As Boolean = False) As String
    InitializeArabicWords
    
    Dim HasilTerbilang As String, TeksSatuan As String, GenderFinal As GenderArab
    Dim AL As String: AL = ChrW(1575) & ChrW(1604) ' Karakter Alif-Lam
    
    If PilihGender = 2 Then GenderFinal = Muzakkar Else GenderFinal = Muannas
    
    ' --- 1. DAFTAR SATUAN  ---
    Select Case JenisSatuan
        Case 1: TeksSatuan = ChrW(1588) & ChrW(1582) & ChrW(1589) ' Syakhsh
        Case 2: TeksSatuan = ChrW(1585) & ChrW(1571) & ChrW(1587) ' Ra's
        Case 3: TeksSatuan = ChrW(1581) & ChrW(1576) & ChrW(1577) ' Habbah
        Case 4: TeksSatuan = ChrW(1603) & ChrW(1578) & ChrW(1575) & ChrW(1576) ' Kitab
        Case 5: TeksSatuan = ChrW(1605) & ChrW(1578) & ChrW(1585) ' Metre
        Case 6: TeksSatuan = ChrW(1603) & ChrW(1610) & ChrW(1604) & ChrW(1608) & ChrW(1594) & ChrW(1585) & ChrW(1575) & ChrW(1605) ' Kg
        Case 7: TeksSatuan = ChrW(1587) & ChrW(1575) & ChrW(1593) & ChrW(1577) ' Sa'ah
        Case 8: TeksSatuan = ChrW(1610) & ChrW(1608) & ChrW(1605) ' Yaum
        Case 9: TeksSatuan = ChrW(1583) & ChrW(1608) & ChrW(1585) ' Dawr
        Case 10: TeksSatuan = ChrW(1591) & ChrW(1606) ' Ton
        Case 11: TeksSatuan = ChrW(1604) & ChrW(1578) & ChrW(1585) ' Litre
        Case 12: TeksSatuan = ChrW(1603) & ChrW(1610) & ChrW(1604) & ChrW(1608) & ChrW(1605) & ChrW(1578) & ChrW(1585) ' Km
        Case 15: TeksSatuan = IIf(IsOrdinal, AL, "") & ChrW(1601) & ChrW(1589) & ChrW(1604) ' Al-Fashlu / Fashl
        Case 16: TeksSatuan = IIf(IsOrdinal, AL, "") & ChrW(1605) & ChrW(1585) & ChrW(1578) & ChrW(1576) & ChrW(1577) ' Al-Martabah / Martabah
        Case 17: TeksSatuan = IIf(IsOrdinal, AL, "") & ChrW(1601) & ChrW(1584) & ChrW(1602) & ChrW(1577) ' Al-Firqah / Firqah
        Case Else: TeksSatuan = ""
    End Select
    
    ' --- 2. LOGIKA TERBILANG ---
    If IsOrdinal Then
        ' Pola Ma'rifah: Alif Lam + Ordinal
        HasilTerbilang = AL & GetArabicOrdinal(Val(angka), GenderFinal)
    Else
        ' Pola Nakirah: Angka Kardinal Biasa
        HasilTerbilang = Arab(angka, 1, GenderFinal)
    End If

    ' --- 3. PENGGABUNGAN AKHIR (Final Output) ---
    If IsOrdinal And TeksSatuan <> "" Then
        ' Urutan (Na't): Satuan + Angka (Contoh: Al-Fashlu Al-Khamis)
        ArabUniversal = TeksSatuan & " " & HasilTerbilang
    ElseIf TeksSatuan <> "" Then
        ' Jumlah ('Adad): Angka + Satuan (Contoh: Khamsatu Fashlin)
        ArabUniversal = HasilTerbilang & " " & TeksSatuan
    Else
        ' Hanya Angka
        ArabUniversal = HasilTerbilang
    End If
End Function

Private Function GetArabicOrdinal(ByVal n As Integer, ByVal g As GenderArab) As String
    ' Memastikan data sudah diinisialisasi
    If Not IsInitialized Then InitializeArabicWords
    
    ' Validasi rentang 1-12
    If n >= 1 And n <= 12 Then
        ' Mengambil langsung dari array Ordinal(Angka, Gender)
        GetArabicOrdinal = Ordinal(n, g)
    Else
        GetArabicOrdinal = ""
    End If
End Function
