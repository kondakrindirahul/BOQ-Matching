
# coding: utf-8

# --> The required master BOQ's are extracted from the excel sheet and 
#     stored into python dictionaires.

# In[1]:

import xlrd
from collections import Counter
import math
import operator
from nltk.corpus import stopwords
from nltk.stem import PorterStemmer
from PyDictionary import PyDictionary

workbook = xlrd.open_workbook('master_boq_from_table.xls')
worksheet = workbook.sheet_by_name('Sheet1')

columns = []
n = worksheet.ncols
stemmer = PorterStemmer()
dictionary = PyDictionary()

for i in range(0, n):
    columns.append(worksheet.cell(0, i))

boq_codes = []
for i in range(0, worksheet.nrows):
    boq_codes.append(worksheet.cell(i, 0))
    
boq_units_description = []
for i in range(0, worksheet.nrows):
    boq_units_description.append(worksheet.cell(i, 4))
    
corpus = {}
description_corpus = {}
for i in range(1, worksheet.nrows):
    description_corpus[boq_codes[i]] = worksheet.cell(i, 2)

# for each in description_dict.keys():
#     print(each)
#     print(description_dict[each])


# --> Creating a dictionary called corpus in which
#     key is the doc_id and value is the content of document. 

# In[20]:


# corpus = {'doc1': 'Electrical Works Earth wire 15mm',
#           'doc2': 'Electrical Works Earth wire 66mm',
#           'doc3': 'Electrical Works Dismantling of Earth wire and cable wire',
#           'doc4': 'Electrical Works High Tensile steel wire',
#           'doc5': 'Electrical Works Erection of Hardware & Wire under Normal conditions',
#           'doc6': 'Electrical Works Erection of Hardware & Wire under Shut down conditions',
#           'doc7': 'Drainage & Protection Works Retaining wall Earth work Excavation',
#           'doc8': 'Drainage & Protection Works Retaining wall RCC M30.0 for Sub Structure'}

# print(corpus)


# --> Converting the content of document into list of tokens 
#     that are converted into lower case.
# --> Removing the stopwords from the corpus text.
# --> Stemming the words in the list

# In[2]:

for each in description_corpus.keys():
    lower_case = str(description_corpus[each]).lower().split()
    filtered_words = [word for word in lower_case if word not in stopwords.words('english')]
    
    stemmed_words = []
    for each_word in filtered_words:
        stemmed_word = stemmer.stem(each_word)
        stemmed_words.append(stemmed_word)
    
    corpus[each] = stemmed_words
    
    
for each in corpus.keys():
    print(each)
    print(corpus[each])


# --> A function to build an inverted index in which 
#     the key is token and value is again a dictionary
#     in which key is document id and value is count of
#     the token in the document.

# In[3]:

def build_inverted_index(corpus):
    inverted_index = {}
    
    for each in corpus.keys():
        doc_id = each
        content = corpus[each]
        word_count = dict(Counter(content))
        
        for token in content:
            if token not in inverted_index:
                temp_dict = dict()
                temp_dict[doc_id] = word_count[token]
                inverted_index[token] = temp_dict
            
            else:
                temp = inverted_index[token]
                temp[doc_id] = word_count[token]
                inverted_index[token] = temp
    
    return inverted_index
       
    
#     for each in inverted_index.keys():
#         print(each)
#         print(inverted_index[each])
        

build_inverted_index(corpus)


# --> Calculating the tf_idf score for each document
#     when given a project_BOQ.
# --> In the dictionary tf_idf_scores, the key is doc_id and
#     the value is the corresponding tf_idf score.

# In[4]:

def calculate_TF_IDF(corpus, description_corpus):
    input_BOQ = input("enter a project_BOQ : ")
    inverted_index = build_inverted_index(corpus)
    
    tf_idf_scores = {}
    # scores_description = {}
    lowercase_input = input_BOQ.lower()
    inputBOQ_list = lowercase_input.split()
    stopped_terms = [word for word in inputBOQ_list if word not in stopwords.words('english')]
    
    project_terms = []
    synonyms_list = []
    for each_word in stopped_terms:
        stemmed_term = stemmer.stem(each_word)
        project_terms.append(stemmed_term)
        
    # adding synonyms to the project_terms list
    for term in project_terms:
        synonyms_list.extend(dictionary.synonym(term))
    
    project_terms.extend(synonyms_list)
    print(project_terms)
    
    doc_count = len(corpus)
    
    for each_doc in corpus.keys():
        doc_id = each_doc
        text = description_corpus[each_doc]
        
        doc_tf_idf = 0
        
        for each_term in project_terms:
            if each_term not in inverted_index:
                continue
            
            tf_component = corpus[each_doc].count(each_term)/float(len(corpus[each_doc]))
            idf_component = math.log(float(doc_count)/len(inverted_index[each_term]))
            doc_tf_idf += tf_component * idf_component
            
        tf_idf_scores[text] = doc_tf_idf
        
    sorted_tf_idf_scores = sorted(tf_idf_scores.items(), key=operator.itemgetter(1), reverse = True)
    
    return sorted_tf_idf_scores

 
calculate_TF_IDF(corpus, description_corpus)

