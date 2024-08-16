import os, re
import pandas as pd

"""
    Dotais skripts nolasa datus no abiem excel failiem un attiecīgi sagatavo XML izejas failu.
    XML faila struktūras atkāpes pēc tam iegūst caur Notepad++ XML tools piespraudnes (plugin), kur izvēlas 'Pretty print' opciju
    
    Skripts darbojas atrodoties vienā mapē ar abiem excel failiem 
"""

# kopīgais dokumentu skaits
DOC_COUNT = 16
EXCEL_FILE = 'Dokumentu_saraksts_8.xlsx'
EXCEL_FILE_2 = 'Glabajamo_vienibu_saraksts_8.xlsx'

type = ""

# pirmā excel faila dati
class Document:
    def __init__(self, obj_nr, obj_name, data_name, checksum, checksum_val, comment):
        self.obj_nr = obj_nr
        self.obj_name = obj_name
        self.data_name = data_name
        self.checksum = checksum
        self.checksum_val = checksum_val
        self.comment = comment

# otrā excel faila dati        
class Document2:
    def __init__(self, obj_name, obj_date):
        self.obj_name = obj_name
        self.obj_date = obj_date

# izdrukā XML failu        
def print_to_xml():
    global type
    global list, list2

    with open('output.xml', 'a') as f:
    
        f.write(f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>\n<Registrs>')
        for i in range(16):
            f.write(f'<Dokuments>\n<DokumentaTips>{type}</DokumentaTips>\n<DokumentaNosaukums>{list[i].obj_name}</DokumentaNosaukums>\n<DokumentaDatums>{str(list2[i].obj_date).split()[0]}</DokumentaDatums>')
            
            f.write(f'<Datne>\n<DatnesNosaukums>{list[i].data_name}</DatnesNosaukums>\n<Kontrolsummas>\n<KontrolsummasAlg>{str(list[i].checksum).split("; ")[0]}</KontrolsummasAlg>\n<Kontrolsumma>{str(list[i].checksum_val).split("; ")[0]}</Kontrolsumma>\n</Kontrolsummas><Kontrolsummas>\n<KontrolsummasAlg>{str(list[i].checksum).split("; ")[1]}</KontrolsummasAlg>\n<Kontrolsumma>{str(list[i].checksum_val).split("; ")[1]}</Kontrolsumma>\n</Kontrolsummas>\n</Datne>')
            
            f.write(f'<GV_Numurs>{list[i].obj_nr}</GV_Numurs>')
            
            # ja ir 'piezīmes', tad pievieno attiecīgo XML tagu
            if list[i].comment != 0:
                f.write(f'<Piezimes>{list[i].comment}</Piezimes>')
            
            f.write(f'</Dokuments>')
        
        f.write(f'</Registrs>\n')    

### galvenā skripta daļa

df = pd.read_excel(EXCEL_FILE)
# Null jeb pandas 'nan' vētības pārveido par '0'
df = df.fillna(0)

df2 = pd.read_excel(EXCEL_FILE_2)
list = []
list2 = []

# print(df)
# print(df2)
if "tekstuāli" in df2.iat[3,0]:
    type = "tekstuālais" 

# 7 ir kur sākas dokumentu dati no dotā excel faila pēc pandas loģikas
for i in range(7, 7 + DOC_COUNT):
    tmp = Document(df.iat[i, 0], df.iat[i, 2], df.iat[i, 3], df.iat[i, 4], df.iat[i, 5], df.iat[i, 6])
    list.append(tmp)

# 12 ir kur sākas dokumentu dati no dotā otrā excel faila pēc pandas loģikas
for i in range(12, 12 + DOC_COUNT):
    tmp = Document2(df2.iat[i, 2], df2.iat[i, 3])
    list2.append(tmp)
      
print_to_xml()
