#!This code advances slide in a powerpoint based on the similarity between a dynamically being produced speech to a refrence text file

# *Importing required libraries
import gensim
import nltk
import numpy as np
from nltk.tokenize import word_tokenize, sent_tokenize
import speech_recognition as sr
import win32com.client

# *Opening presentation
app = win32com.client.gencache.EnsureDispatch('Powerpoint.Application')
app.Visible = True
mine = app.Presentations.Open(
    'C:/Users/amir/Documents/Slide-Transition/power.pptx')
mine.SlideShowSettings.Run()

# *Creating similarity object for refrernce.txt
file_docs = []
with open('reference.txt') as f:
    tokens = sent_tokenize(f.read())
    for item in tokens:
        file_docs.append(item)
gen_docs = [[w.lower() for w in word_tokenize(text)]
            for text in file_docs]
dictionary = gensim.corpora.Dictionary(gen_docs)
corpus = [dictionary.doc2bow(gen_doc) for gen_doc in gen_docs]
tf_idf = gensim.models.TfidfModel(corpus)
sims = gensim.similarities.Similarity(
    'workdir/', tf_idf[corpus], num_features=len(dictionary))

# *Converting speech to text
r = sr.Recognizer()
with sr.Microphone() as source:
    print("Start speaking :")
    suzo = 0
    ac_text = ''
    while suzo < 50:
        audio = r.listen(source)
        text = r.recognize_google(audio)
        ac_text = ac_text + text
        file2_docs = []
        tokens = sent_tokenize(ac_text)
        for line in tokens:
            file2_docs.append(line)
        avg_sims = []
        for line in file2_docs:
            query_doc = [w.lower() for w in word_tokenize(text)]
            query_doc_bow = dictionary.doc2bow(query_doc)
            query_doc_tf_idf = tf_idf[query_doc_bow]
            sum_of_sims = (np.sum(sims[query_doc_tf_idf], dtype=np.float32))
            avg = sum_of_sims/len(file_docs)
            avg_sims.append(avg)
            total_avg = np.sum(avg_sims, dtype=np.float)
            suzo = round(float(total_avg) * 100)
            if suzo >= 100:
                suzo = 100
    mine.SlideShowWindow.View.Next()
