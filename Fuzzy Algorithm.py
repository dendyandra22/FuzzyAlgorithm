
import pandas as pd
import xlwt 
from xlwt import Workbook 
import random

# fungsi write excel
def writeExcel(namefile, data):
    wb = Workbook()
    row, column = 1,0
    sheet1 = wb.add_sheet('Sheet 1') 
    sheet1.write(0,0,'Nomor penerima')
    for item in data:
        sheet1.write(row, column, item)
        row += 1
    wb.save(namefile)

# fungsi read excel
def readExcel(name):
    return pd.read_excel(name)

# fungsi hitung nilai fuzzy penghasilan
def fuzzyPenghasilan(penghasilan):
    if penghasilan <= 5:
        hasil = [['kecil', 1], ['sedang', 0]]
    elif 5 < penghasilan < 7:
        x = -(penghasilan - 7)/(7 - 5)
        hasil = [['kecil', x], ['sedang', 1-x]]
    elif 7 <= penghasilan <= 10:
        hasil = [['sedang', 1], ['besar', 0]]
    elif 10 < penghasilan < 12:
        x = -(penghasilan - 12)/(12 - 10)
        hasil = [['sedang', x], ['besar', 1-x]]
    elif 12 <= penghasilan <= 15:
        hasil = [['besar', 1], ['sangat besar', 0]]
    elif 15 < penghasilan < 17:
        x = -(penghasilan - 17)/(17 - 15)
        hasil = [['besar', x], ['sangat besar', 1-x]]
    elif penghasilan >= 17:
        hasil = [['besar', 0], ['sangat besar', 1]]
    return hasil

# fungsi hitung nilai fuzzy pengeluaran
def fuzzyPengeluaran(pengeluaran):
    if pengeluaran <= 4:
        hasil = [['kecil', 1], ['sedang', 0]]
    elif 4 < pengeluaran < 6:
        x = -(pengeluaran - 6)/(6 - 4)
        hasil = [['kecil', x], ['sedang', 1-x]]
    elif 6 <= pengeluaran <= 7:
        hasil = [['sedang', 1], ['besar', 0]]
    elif 7 < pengeluaran < 9:
        x = -(pengeluaran - 9)/(9 - 7)
        hasil = [['sedang', x], ['besar', 1-x]]
    elif 9 <= pengeluaran <= 10:
        hasil = [['besar', 1], ['sangat besar', 0]]
    elif 10 < pengeluaran < 11:
        x = -(pengeluaran - 11)/(11 - 10)
        hasil = [['besar', x], ['sangat besar', 1-x]]
    elif pengeluaran >= 11:
        hasil = [['besar', 0], ['sangat besar', 1]]
    return hasil

# fungsi fuzzification
def fuzzyfication(penghasilan, pengeluaran):
    hasilPenghasilan = fuzzyPenghasilan(penghasilan)
    hasilPengeluaran = fuzzyPengeluaran(pengeluaran)
    x = hasilPenghasilan + hasilPengeluaran
    return x

# fungsi konjungsi aturan fuzzy
def fuzzyConjunction(nilai1, nilai2):
    if nilai1[0] == 'kecil' and nilai2[0] == 'kecil':
        if(nilai1[1] < nilai2[1]):
            nilaiNk = nilai1[1]
        else:
            nilaiNk = nilai2[1]
        nk = ['tinggi', nilaiNk]
    elif nilai1[0] == 'kecil' and nilai2[0] == 'sedang':
        if(nilai1[1] < nilai2[1]):
            nilaiNk = nilai1[1]
        else:
            nilaiNk = nilai2[1]
        nk = ['tinggi', nilaiNk]
    elif nilai1[0] == 'kecil' and nilai2[0] == 'besar':
        if(nilai1[1] < nilai2[1]):
            nilaiNk = nilai1[1]
        else:
            nilaiNk = nilai2[1]
        nk = ['tinggi', nilaiNk]
    elif nilai1[0] == 'kecil' and nilai2[0] == 'sangat besar':
        if(nilai1[1] < nilai2[1]):
            nilaiNk = nilai1[1]
        else:
            nilaiNk = nilai2[1]
        nk = ['tinggi', nilaiNk]
    elif nilai1[0] == 'sedang' and nilai2[0] == 'kecil':
        if(nilai1[1] < nilai2[1]):
            nilaiNk = nilai1[1]
        else:
            nilaiNk = nilai2[1]
        nk = ['rendah', nilaiNk]
    elif nilai1[0] == 'sedang' and nilai2[0] == 'sedang':
        if(nilai1[1] < nilai2[1]):
            nilaiNk = nilai1[1]
        else:
            nilaiNk = nilai2[1]
        nk = ['tinggi', nilaiNk]
    elif nilai1[0] == 'sedang' and nilai2[0] == 'besar':
        if(nilai1[1] < nilai2[1]):
            nilaiNk = nilai1[1]
        else:
            nilaiNk = nilai2[1]
        nk = ['tinggi', nilaiNk]
    elif nilai1[0] == 'sedang' and nilai2[0] == 'sangat besar':
        if(nilai1[1] < nilai2[1]):
            nilaiNk = nilai1[1]
        else:
            nilaiNk = nilai2[1]
        nk = ['tinggi', nilaiNk]
    elif nilai1[0] == 'besar' and nilai2[0] == 'kecil':
        if(nilai1[1] < nilai2[1]):
            nilaiNk = nilai1[1]
        else:
            nilaiNk = nilai2[1]
        nk = ['rendah', nilaiNk]
    elif nilai1[0] == 'besar' and nilai2[0] == 'sedang':
        if(nilai1[1] < nilai2[1]):
            nilaiNk = nilai1[1]
        else:
            nilaiNk = nilai2[1]
        nk = ['rendah', nilaiNk]
    elif nilai1[0] == 'besar' and nilai2[0] == 'besar':
        if(nilai1[1] < nilai2[1]):
            nilaiNk = nilai1[1]
        else:
            nilaiNk = nilai2[1]
        nk = ['rendah', nilaiNk]
    elif nilai1[0] == 'besar' and nilai2[0] == 'sangat besar':
        if(nilai1[1] < nilai2[1]):
            nilaiNk = nilai1[1]
        else:
            nilaiNk = nilai2[1]
        nk = ['tinggi', nilaiNk]
    elif nilai1[0] == 'sangat besar' and nilai2[0] == 'kecil':
        if(nilai1[1] < nilai2[1]):
            nilaiNk = nilai1[1]
        else:
            nilaiNk = nilai2[1]
        nk = ['rendah', nilaiNk]
    elif nilai1[0] == 'sangat besar' and nilai2[0] == 'sedang':
        if(nilai1[1] < nilai2[1]):
            nilaiNk = nilai1[1]
        else:
            nilaiNk = nilai2[1]
        nk = ['rendah', nilaiNk]
    elif nilai1[0] == 'sangat besar' and nilai2[0] == 'besar':
        if(nilai1[1] < nilai2[1]):
            nilaiNk = nilai1[1]
        else:
            nilaiNk = nilai2[1]
        nk = ['rendah', nilaiNk]
    elif nilai1[0] == 'sangat besar' and nilai2[0] == 'sangat besar':
        if(nilai1[1] < nilai2[1]):
            nilaiNk = nilai1[1]
        else:
            nilaiNk = nilai2[1]
        nk = ['rendah', nilaiNk]
    return nk

# fungsi disjungsi aturan fuzzy
def fuzzyDisjunction(hasilConjunction):
    rendah = []
    tinggi = []
    hasil = []

    nkRendah = 0
    nkTinggi = 0

    for item in hasilConjunction:
        if (item[0] == 'rendah'):
            rendah.append(item)
        else:
            tinggi.append(item)
    if len(rendah) != 0:
        nkRendah = 0
        for item in rendah:
            if item[1] >= nkRendah:
                nkRendah = item[1]
    hasil.append(['rendah', nkRendah])

    if len(tinggi) != 0:
        nkTinggi = 0
        for item in tinggi:
            if item[1] >= nkTinggi:
                nkTinggi = item[1]
    hasil.append(['tinggi', nkTinggi])

    return hasil

# fungsi inferensi
def inference(nilaiFuzzy):
    con = []
    for i in range(2):
        j = 2
        while j < 4:    
            con.append(fuzzyConjunction(nilaiFuzzy[i], nilaiFuzzy[j]))
            j += 1
    return fuzzyDisjunction(con)

# fungsi defuzzyfication
def defuzzyfication(hasilInference):
    res = 0
    nkRendah = 0
    nkTinggi = 0
    tempRen, tempTeng, tempTeng2, tempTing = 0,0,0,0
    rendah = []
    tengah = []
    tinggi = []
    
    if len(hasilInference) == 1:
        if hasilInference[0][0] == 'rendah':
            nkRendah = hasilInference[0][1]
        elif hasilInference[0][0] == 'tinggi':
            nkTinggi = hasilInference[0][1]
    else:
        nkRendah = hasilInference[0][1]
        nkTinggi = hasilInference[1][1]
    
    for i in range(10):
        x = random.randrange(0,101)
        if x <= 50:
            rendah.append(x)
            
        elif 50 < x <= 80:
            tengah.append(x)
            
        elif x > 80:
            tinggi.append(x)

    for n in rendah:
        tempRen += n
    tempRen = tempRen * nkRendah

    for n in range(len(tengah)):
        if hasilInference[0][1] <= hasilInference[1][1]:
            tempTeng += tengah[n] * (tengah[n] / 130)
            tempTeng2 += tengah[n] / 130
        else:
            tempTeng += -((tengah[n] / 130) - 1) * tengah[n]
            tempTeng2 += -((tengah[n] / 130) - 1) * tengah[n]

    for n in tinggi:
        tempTing += n
    tempTing = tempTing * nkTinggi
    penyebut = (nkRendah * len(rendah)) + tempTeng2 + (nkTinggi * len(tinggi))
    res = (tempRen + tempTeng + tempTing) / penyebut
    return res

# Main program
data = readExcel('Mahasiswa.xls')
nilaiLayak, penerima = [],[]
for i in range(0,100):
    x = fuzzyfication(data['Penghasilan'][i], data['Pengeluaran'][i])
    nk = inference(x)
    nilaiLayak.append([i+1,defuzzyfication(nk)])
hasil = sorted(nilaiLayak, key = lambda x: x[1], reverse=True)
for i in range(20):
    penerima.append(hasil[i][0])
writeExcel('Bantuan.xls',penerima)

