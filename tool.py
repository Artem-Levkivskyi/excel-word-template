from docxtpl import DocxTemplate
import openpyxl

test = []
wb = openpyxl.load_workbook('E:/исходник.xlsx')
sheet = wb['Лист1']

for row in sheet['A4':'D323']:
    for cellObj in row:
        if cellObj.value is None or cellObj.value == " ":
            continue
        test.append(cellObj.value)

x = 0
while x < len(test):
    doc = DocxTemplate("E:/шаблон.docx")
    context = {'contract': test[x], 'name': str(test[x+1]), 'flat': str(test[x+2]), 'phone': str(test[x+3])}
    doc.render(context)
    doc.save("E:/" + str(test[x+2]) + '.docx')
    x += 4
