import openpyxl


# Excel dosyası ve işlem parametrelerini belirleyin
excel_dosya = 'de.xlsx'
sayi = 2.4
islem = 'bölme'
sira = 'AI' # Değiştirmek istediğiniz sütunun harf değerini belirtin


async def islem_yap(excel_dosya, sayi, islem, sira):
    # Excel dosyasını yükle
    workbook = openpyxl.load_workbook(excel_dosya)
    sheet = workbook.active

    # İlgili sütunu seç
    sütun = sheet[sira + '1':sira + str(sheet.max_row)]

    # Sütun üzerinde işlem yap
    for row in sütun:
        for cell in row:
            if isinstance(cell.value, int) or isinstance(cell.value, float):
                if islem == 'çarpma':
                    cell.value = cell.value * sayi
                elif islem == 'bölme':
                    cell.value = cell.value / sayi
                elif islem == 'toplama':
                    cell.value = cell.value + sayi
                elif islem == 'çıkarma':
                    cell.value = cell.value - sayi

    # Değişiklikleri kaydet
    workbook.save(excel_dosya)

# İşlemi gerçekleştir
import asyncio
loop = asyncio.get_event_loop()
loop.run_until_complete(islem_yap(excel_dosya, sayi, islem, sira))
print('İşlem tamamlandı')
