from fuzzywuzzy import fuzz
import docx

KEY_COMPARE = ['зміст Послідовність виконання та   робіт кожного', 'Положення політики безпеки,  пов’язані ',
               'позначень, ідентифікації об’єктів']
doc = docx.Document('kszi.docx')
# # количество абзацев в документе
dlinadoc = (len(doc.paragraphs))
#
# # текст первого абзаца в документе
# print(doc.paragraphs[0].text)
z = 0

for i in KEY_COMPARE:
    z += 1
    print(f'CRUG PROVERKI = {z}')
    for abzac in doc.paragraphs:
        x = fuzz.partial_ratio(abzac.text, i)
        if x >= 80:
            print(x)
            print(abzac.text)
        elif x >= 60:
            print(f'x = {x}')
            print(f'abzac = {abzac.text}')

