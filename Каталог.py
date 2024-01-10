import openpyxl
import pandas as pd
from nltk.tokenize import word_tokenize

from nltk.stem.snowball import SnowballStemmer

import nltk

def tokenize(sentences):
    for sent in nltk.sent_tokenize(sentences.lower()):
        for word in nltk.word_tokenize(sent):
            yield word


text = nltk.Text(tkn for tkn in tokenize('mary had a little lamb.'))
print(text)
print(text.collocations(num=20))
collocations = [" ".join(el) for el in list(text._collocations)]
print(collocations)