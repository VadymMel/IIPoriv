from fuzzywuzzy import fuzz
import docx

KEY_COMPARE = ['Послідовність зміст виконання та   робіт кожного', 'Положення політики безпеки,  пов’язані ',
               'позначень, ідентифікації об’єктів']
doc = docx.Document('kszi.docx')
# # количество абзацев в документе
dlinadoc = (len(doc.paragraphs))
SLOVAR = {}
# # текст первого абзаца в документе
# print(doc.paragraphs[0].text)
z = 0
rezult = 0
max_num = 0
for i in KEY_COMPARE:
    z += 1
    print(f'CRUG PROVERKI = {z}')
    for abzac in doc.paragraphs:
        x = fuzz.partial_ratio(abzac.text, i)
        if x > rezult:
            if x >= 60:
                rezult = x
            else:
                print(x)
#         if x >= 30:
#             SLOVAR.update(({x : i}))
#     for keys in SLOVAR:
#         if keys >= rezult:
#             rezult = keys
#         else:
#             print(keys)
# print(SLOVAR.keys())
