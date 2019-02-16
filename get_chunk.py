#!/usr/bin/env python
# coding: utf-8

# In[1]:


get_ipython().system('pip3 install spacy')


# In[2]:


# get_ipython().system('python -m spacy download en')


# In[5]:


import xlrd
import xlwt
import spacy
nlp = spacy.load('en')

wb = xlrd.open_workbook('20180930_Sample_Data_USSIC_0100-0971_Listed_Agri.xls')
sheet = wb.sheet_by_name('20180925_Sample_Data for QC')

workwrite = xlwt.Workbook()
sheet1 = workwrite.add_sheet('sheet1', cell_overwrite_ok=True)


for i in range(2, sheet.nrows):
    cell = sheet.cell(i, 8)

    sentence = cell.value
    noun_chunk_test = nlp(sentence)

    products = []
    
    for noun_chunk in noun_chunk_test.noun_chunks:
        products.append(noun_chunk)
        sheet1.write(i, 1, str(products))
    print(products)
    

# workwrite.save('20180930_Sample_Data_USSIC_0100-0971_Listed_Agri.xls')
#     for token in products:
#         if token.ent_type_ == "PRODUCT" or "FAC":
#             products.append(token)
#     print(products)




# In[ ]:




