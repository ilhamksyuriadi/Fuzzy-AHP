# -*- coding: utf-8 -*-
"""
Created on Mon Apr 23 15:39:29 2018

@author: ilhamksyuriadi
"""

import xlrd

#read data nilai mahasiswa
mahasiswa = []
fileLoc = "nilai mahasiswa.xlsx"
workBook = xlrd.open_workbook(fileLoc)
sheet = workBook.sheet_by_index(0)
for row in range(10):
    data = []
    for col in range(sheet.ncols):
        data.append(sheet.cell_value(row,col))
    mahasiswa.append(data)
    
#angka diisi sesuai hasil survey
def generateCriteria():
            # urutan : APPL DAP PBO PBD
    criteria.append([1.0, 1.0, 0.2, 3.0]) #APPL
    criteria.append([1.0, 1.0, 0.14, 5.0]) #DAP
    criteria.append([5.0, 7.0, 1.0, 5.0]) #PBO
    criteria.append([0.33, 0.2, 0.2, 1.0]) #PBD
 
def sumColumn(criteria):
    APPL = 0.0
    DAP = 0.0
    PBO = 0.0
    PBD = 0.0
    for i in range(len(criteria)):
        APPL += criteria[i][0]
        DAP += criteria[i][1]
        PBO += criteria[i][2]
        PBD += criteria[i][3]
    return APPL,DAP,PBO,PBD

def divElm(c,APPL,DAP,PBO,PBD):
    dividedCriteria = []
    for i in range(len(c)):
        data = c[i][0]/APPL,c[i][1]/DAP,c[i][2]/PBO,c[i][3]/PBD
        dividedCriteria.append(data)
    return dividedCriteria

def avgRow(c):
    APPL = (c[0][0]+c[0][1]+c[0][2]+c[0][3])/4
    DAP = (c[1][0]+c[1][1]+c[1][2]+c[1][3])/4
    PBO = (c[2][0]+c[2][1]+c[2][2]+c[2][3])/4
    PBD = (c[3][0]+c[3][1]+c[3][2]+c[3][3])/4
    return APPL,DAP,PBO,PBD

def convertIndeks(indeks):
    if indeks == "A":
        return 4.0
    elif indeks == "AB":
        return 3.5
    elif indeks == "B":
        return 3.0
    elif indeks == "BC":
        return 2.5
    elif indeks == "C":
        return 2.0
    elif indeks == "D":
        return 1.0
    else:
        return 0.0

def decision(mhs,APPl,DAP,PBO,PBD):
    result = []
    for i in range(len(mhs)):
        appl = convertIndeks(mhs[i][1])*APPL
        dap = convertIndeks(mhs[i][2])*DAP
        pbo = convertIndeks(mhs[i][3])*PBO
        pbd = convertIndeks(mhs[i][4])*PBD
        data = appl+dap+pbo+pbd
        result.append([data,mhs[i][0]])
    results  = sorted(result,reverse=True)
    for i in range(4):
        print(results[i])
        
def consistency(c, APPL,DAP,PBO,PBD):
    #step 1, mencari weighted sum
    wc = []
    for i in range(len(c)):
        temp = c[i][0]*APPL,c[i][1]*DAP,c[i][2]*PBO,c[i][3]*PBD
        wc.append(temp)
    wAPPL = wc[0][0]+wc[0]  [1]+wc[0][2]+wc[0][3]
    wDAP = wc[1][0]+wc[1][1]+wc[1][2]+wc[1][3]
    wPBO = wc[2][0]+wc[2][1]+wc[2][2]+wc[2][3]
    wPBD = wc[3][0]+wc[3][1]+wc[3][2]+wc[3][3]
    #step 2, membagi weighted sum dengan corresponding priority value
    appl = wAPPL/APPL
    dap = wDAP/DAP
    pbo = wPBO/PBO
    pbd = wPBD/PBD
    semua = appl + dap + pbo + pbd
    #step 3, mencari lmax
    avg = semua / 4
    #step 4, mencari CI
    CI = (avg-4)/(4-1)
    #step 5, mencari CR
    RI = 0.89
    CR = CI / RI
    return CR
    
criteria = []
generateCriteria()#criteria sudah di isi sesuai hasil survey
APPL,DAP,PBO,PBD = sumColumn(criteria)#step 1
pairwise = divElm(criteria,APPL,DAP,PBO,PBD)#step 2
cAPPL,cDAP,cPBO,cPBD = avgRow(pairwise)#step 3
cons = consistency(pairwise,cAPPL,cDAP,cPBO,cPBD)#penghitungan consistency

#melakukan perulangan selama consistency < 0.10
while cons > 0.1:
    APPL,DAP,PBO,PBD = sumColumn(pairwise)
    newPairwise = divElm(pairwise,APPL,DAP,PBO,PBD)
    cAPPL,cDAP,cPBO,cPBD = avgRow(newPairwise)
    cons = consistency(newPairwise,cAPPL,cDAP,cPBO,cPBD)
    pairwise = newPairwise

#memberikan saran keputusan    
decision(mahasiswa,cAPPL,cDAP,cPBO,cPBD)
print(cons)


