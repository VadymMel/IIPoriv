import docx
import numpy as np
from fuzzywuzzy import fuzz
import docx

KEY_COMPARE = ['Послідовність виконання та  зміст робіт кожного', 'Положення політики безпеки,  пов’язані ',
               'позначень, ідентифікації об’єктів']
doc = docx.Document('kszi.docx')
# # количество абзацев в документе
dlinadoc = (len(doc.paragraphs))
#
# # текст первого абзаца в документе
# print(doc.paragraphs[0].text)
z = 0

for i in KEY_COMPARE:

    for abzac in doc.paragraphs:
        x = fuzz.partial_ratio(abzac.text, i)
        if x >= 80:

            print(x)
            print(abzac.text)

