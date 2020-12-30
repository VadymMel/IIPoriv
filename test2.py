from fuzzywuzzy import fuzz
import docx
import openpyxl
from pprint import pprint
from openpyxl.styles import Font
from openpyxl import load_workbook
from tkinter.filedialog import askopenfilename, asksaveasfilename

KEY_COMPARE = ['Кінцевою датою погашення заборгованості по позикових коштах ',
               'За користування позиковими коштами Позичальник виплачує',
               'позначень, ідентифікації об’єктів', 'ЄДРПОУ', 'ТОВ', 'ДОГОВІР ПОЗИКИ', 'Обсяг Позики у сумі']
MONTH = {'01': 'січеня', '02': 'лютого', '03': 'березня', '04': 'квітня', '05': 'травня', '06': 'червня',
          '07': 'липня', '08': 'серпня', '09': 'вересня', '10': 'жовтня',
          '11': 'листопада', '12': 'грудня'}
VALUTA = 'євро'
DATA = []
KOD_NAZ = []
z = 0
rezult = 0
max_num = 0

wb = openpyxl.Workbook()
wb.create_sheet(title='ЗЛ', index=0)
sheet = wb['ЗЛ']
doc = docx.Document('dogovor2.docx')
dlinadoc = (len(doc.paragraphs))
table = doc.tables[0]
table_size = len(table.rows)

for i in KEY_COMPARE:
    # z += 1
    # print(f'CRUG PROVERKI = {z}')
    # for abzac in doc.paragraphs:
    #     x = fuzz.partial_ratio(abzac.text, i)
    #     if x > 70:
    #         print(f'{x} - {abzac.text}')
    #         # value = abzac.text
    #         # cell = sheet.cell(row=i, column=1)
    #         # cell.value = value
    for j in range(0, table_size):
        for zx in table.rows[j].cells[0].paragraphs:
            x = fuzz.partial_ratio(zx.text, i)
            if x > 85:
                if zx.text == ' ':
                    pass

                else:
                    if 'позики' in zx.text.lower():
                        value = zx.text[0:(zx.text.find("№"))]
                        cell = sheet.cell(row=5, column=2)
                        cell.value = value
                        value = zx.text[zx.text.find("№") + 1:]
                        cell = sheet.cell(row=6, column=2)
                        cell.value = value
                    for slova in zx.text.replace('"', '').split():
                        '''определение ЕДРПОУ'''
                        try:
                            EDRPOU = int(slova)
                            if len(str(EDRPOU)) == 8:
                                print(slova)
                                KOD_NAZ.append(f'{slova}; ')
                        except ValueError:
                            pass

                        '''определение даты'''
                        try:
                            datte = int(slova)
                            if 2020 < datte < 2050 or 1 <= datte <= 31:
                                DATA.append(slova)
                        except ValueError:
                            for key, value in MONTH.items():
                                if slova == value:
                                    DATA.append(key)
                                else:
                                    pass

                    if 'ТОВ' in zx.text:
                        '''определение названия'''
                        print(zx.text[zx.text.find('ТОВ'):(zx.text.find(','))])
                        KOD_NAZ.append(zx.text[zx.text.find('ТОВ'):(zx.text.find(','))])
                    if 'сумі' in zx.text:
                        for slova in zx.text.replace('.', '').replace(',', '.').split():

                            try:
                                summ = float(slova)
                                print(summ)
                            except ValueError:
                                        pass
                            if VALUTA == slova.lower():
                                print(VALUTA)
                        value = f'{summ}; {VALUTA.upper()}'
                        cell = sheet.cell(row=10, column=2)
                        cell.value = value

value = ''.join(KOD_NAZ)
cell = sheet.cell(row=4, column=2)
cell.value = value
value = '.'.join(DATA)
cell = sheet.cell(row=11, column=2)
cell.value = value
wb.save("123.xlsx")
