import pandas as pd
import json
import sys

def main(kaynak_excel_dosyasi, baslangic_sutunu, bitis_sutunu, cikis_dosyasi="output.json"):
    # Excel dosyasını oku ve boş hücreleri boş string ile doldur
    df = pd.read_excel(kaynak_excel_dosyasi, engine='openpyxl').fillna('')

    # Sütunları indeks numaralarına çevir
    baslangic_index = ord(baslangic_sutunu.upper()) - ord('A')
    bitis_index = ord(bitis_sutunu.upper()) - ord('A')
    
    # Sütunlar arasındaki verileri al
    data = []
    for index, row in df.iterrows():
        entry = {}
        for col in range(baslangic_index, bitis_index + 1):
            column_name = df.columns[col]
            entry[column_name] = row[column_name]
        data.append(entry)
    
    # JSON dosyasına yaz
    with open(cikis_dosyasi, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

if __name__ == "__main__":
    if len(sys.argv) < 4 or len(sys.argv) > 5:
        print("Usage: python main.py sourceExcel FirstColumn LastColumn [OutputFile.json]")
        sys.exit(1)
    
    kaynak_excel_dosyasi = sys.argv[1]
    baslangic_sutunu = sys.argv[2]
    bitis_sutunu = sys.argv[3]
    cikis_dosyasi = sys.argv[4] if len(sys.argv) == 5 else "output.json"

    main(kaynak_excel_dosyasi, baslangic_sutunu, bitis_sutunu, cikis_dosyasi)
