from fuzzywuzzy import fuzz
import docx
import openpyxl
from pprint import pprint
from openpyxl.styles import Font
from openpyxl import load_workbook
from tkinter.filedialog import askopenfilename, asksaveasfilename

wb = openpyxl.Workbook()
wb.create_sheet(title='ЗЛ', index=0)
sheet = wb['ЗЛ']

KEY_COMPARE = ['Кінцевою датою погашення заборгованості по позикових коштах ',
               'За користування позиковими коштами Позичальник виплачує',
               'позначень, ідентифікації об’єктів']
doc = docx.Document('dogovor2.docx')
# # количество абзацев в документе
dlinadoc = (len(doc.paragraphs))
SLOVAR = {}
# # текст первого абзаца в документе
# print(doc.paragraphs[0].text)
table = doc.tables[0]
table_size = len(table.rows)
z = 0
rezult = 0
max_num = 0
for i in KEY_COMPARE:
    z += 1
    print(f'CRUG PROVERKI = {z}')
    for abzac in doc.paragraphs:
        x = fuzz.partial_ratio(abzac.text, i)
        if x > 70:
            print(f'{x} - {abzac.text}')
            value = abzac.text
            cell = sheet.cell(row=i, column=1)
            cell.value = value
    for j in range(1, table_size):
        for zx in table.rows[j].cells[0].paragraphs:
            x = fuzz.partial_ratio(zx.text, i)
            if x > 85:
                if zx.text == ' ':
                    pass
                else:
                    pprint(f'{x} - {zx.text}')
                    value = zx.text
                    cell = sheet.cell(row=j, column=1)
                    cell.value = value

wb.save("123.xlsx")
#         if x >= 30:
#             SLOVAR.update(({x : i}))
#     for keys in SLOVAR:
#         if keys >= rezult:
#             rezult = keys
#         else:
#             print(keys)
# print(SLOVAR.keys())


# file_name = asksaveasfilename(
#     filetypes=(("Excel files", "*.xlsx"),))
# wb.save(file_name + ".xlsx")
