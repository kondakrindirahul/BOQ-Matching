
# coding: utf-8

# Importing SpaCy and loading the english model.

# In[18]:

import spacy
nlp = spacy.load('en')


# Importing xlrd package to extract the contents of excel file
# into the python code.

# In[2]:

import xlrd
workbook = xlrd.open_workbook('master_boq_from_table.xls')


# In[3]:

workbook.sheet_names()


# In[4]:

worksheet = workbook.sheet_by_name('Sheet1')


# Obtaining the Column values (features) from the excel sheet

# In[26]:

columns = []
n = worksheet.ncols

for i in range(0, n):
    columns.append(worksheet.cell(0, i))

# print(columns)


# Extracting all the BOQ_codes from the boq_codes 
# column in the excel sheet

# In[25]:

boq_codes = []

for i in range(0, worksheet.nrows):
    boq_codes.append(worksheet.cell(i, 0))
#print(boq_codes)    


# Extracting the types of units used in the BOQ descriptions

# In[7]:

boq_units_description = []

for i in range(0, worksheet.nrows):
    boq_units_description.append(worksheet.cell(i, 4))

# print(boq_units_description)


# creating a dictionary where the key is the Master_BOQ_Id 
# and value is the Master_BOQ_Description

# In[8]:

description_dict = {}

for i in range(1, worksheet.nrows):
    description_dict[boq_codes[i]] = worksheet.cell(i, 2)

# for each in description_dict.keys():
#     print(each)
#     print(description_dict[each])


# Using the similarity function provided by SpaCy, 
# we calculate the similarity between a given Project_BOQ_description
# and all the Master_BOQ_Descriptions.
# 
# The top 100 ranked documents are shown in the output.

# In[ ]:

import operator

sentence = (input('enter the project_boq : '))
query = nlp(sentence)
doc_score = {}

for each in description_dict.keys():
    text = nlp(str(description_dict[each]))
#     query.similarity(text)
    
    doc_score[description_dict[each]] = query.similarity(text)

# for each in doc_score.keys():
#     print(each, end = ':')
#     print(doc_score[each])
#     print('\n')

sorted_doc_score = sorted(doc_score.items(), key=operator.itemgetter(1), reverse = True)

for i in range(0,101):
    print(sorted_doc_score[i])




