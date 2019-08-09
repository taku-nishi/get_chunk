import xlrd
import xlwt
import spacy
nlp = spacy.load('en')

wb = xlrd.open_workbook('List_data.xls')
sheet = wb.sheet_by_name('Sample')

sheet_write = xlwt.Workbook()
new_sheet = sheet_write.add_sheet('new_sheet', cell_overwrite_ok=True)

for i in range(1, sheet.nrows):
    cell = sheet.cell(i, 1)

    sentence = cell.value
    noun_chunk_test = nlp(sentence)

    products = []
    
    for noun_chunk in noun_chunk_test.noun_chunks:
        products.append(noun_chunk)
        new_sheet.write(i, 1, str(products))
sheet_write.save('new_sheet.xls')

