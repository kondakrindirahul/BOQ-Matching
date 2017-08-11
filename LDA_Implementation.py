
# coding: utf-8

# --> Required packages for LDA Implementation

# In[1]:

from nltk.tokenize import RegexpTokenizer
from nltk.corpus import stopwords
from nltk.stem.porter import PorterStemmer
from gensim import corpora, models
import gensim
import xlrd

tokenizer = RegexpTokenizer(r'\w+')

# Create p_stemmer of class PorterStemmer
p_stemmer = PorterStemmer()


# --> extracting data from the excel sheets and storing it in a dictionary.

# In[2]:

workbook = xlrd.open_workbook('master_boq_from_table.xls')
worksheet = workbook.sheet_by_name('Sheet1')

columns = []
n = worksheet.ncols

for i in range(0, n):
    columns.append(worksheet.cell(0, i))

boq_codes = []
for i in range(0, worksheet.nrows):
    text = str(worksheet.cell(i, 0))
    boq_codes.append(text[5:])
    
description_dict = {}
for i in range(1, worksheet.nrows):
    description = str(worksheet.cell(i, 2))
    description_dict[boq_codes[i]] = description[5:]

for each in description_dict.keys():
    print(each)
    print(description_dict[each])


# --> Creating a list containing master BOQ descriptions

# In[3]:

document_set = []
for each in description_dict.keys():
    document_set.append(description_dict[each])
    


# --> Preprossing the descriptions and creating a document-term matrix for converting into LDA model.

# In[4]:

stop_words = set(stopwords.words('english'))

# list for tokenized documents in loop
texts = []

# loop through document list
for i in document_set:
    
    # clean and tokenize document string
    raw = str(i).lower()
    tokens = tokenizer.tokenize(raw)

    # remove stop words from tokens
    stopped_tokens = [i for i in tokens if not i in stop_words]
    
    # stem tokens
    stemmed_tokens = [p_stemmer.stem(i) for i in stopped_tokens]
    
    # add tokens to list
    texts.append(stemmed_tokens)

for each in texts:
    print(each)


# In[46]:

# turn our tokenized documents into a id <-> term dictionary
dictionary = corpora.Dictionary(texts)

# print(dictionary)
# for each in dictionary:
#     print(each)
#     print(dictionary[each])

# convert tokenized documents into a document-term matrix
corpus = [dictionary.doc2bow(text) for text in texts]

# for content in corpus:
#     print(content)


# In[47]:

# generate LDA model
ldamodel = gensim.models.ldamodel.LdaModel(corpus, num_topics=12, id2word = dictionary, passes=20)


# In[48]:

# ldamodel.print_topics(num_topics=12, num_words=2)

for each in ldamodel.print_topics(num_topics=12, num_words=4):
    print(each)


# --> Assigns the topics to the documents in corpus
# 

# In[22]:

lda_corpus = ldamodel[corpus]

# for each in lda_corpus:
#     print(each)


# --> Inputing Project BOQ and preprocessing the input BOQs'.
# 

# --> Converting the input BOQs into document term matrix and testing them with the pretrained LDAModel.

# In[51]:

stop_words = set(stopwords.words('english'))

# id2word = gensim.corpora.Dictionary()

query_corpus = []
processed_queries = []
query1 = input("enter a Project BOQ : ")
query2 = input("enter another Project BOQ: ")

query_corpus.append(query1)
query_corpus.append(query2)
# print(query_corpus)

for query in query_corpus:
    raw = str(query).lower()
    tokens = tokenizer.tokenize(raw)
    # remove stop words from query
    stopped_query = [i for i in tokens if not i in stop_words]
    # stem query
    processed_query = [p_stemmer.stem(i) for i in stopped_query]
    processed_queries.append(processed_query)
    
# print(processed_queries)

# turn our tokenized documents into a id <-> term dictionary
query2id = corpora.Dictionary(processed_queries)
print(query2id)

# convert tokenized documents into a document-term matrix
query_corpus = [query2id.doc2bow(processed_query) for processed_query in processed_queries]
# for each in query_corpus:
#     print(each)

# assigns topics to the input porject BOQs'
topics_query = ldamodel[query_corpus]
for each_query in topics_query:
    print("the topics and their probability distributions in the Project BOQ are : ")
    print(each_query)

