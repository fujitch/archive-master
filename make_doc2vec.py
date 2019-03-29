# -*- coding: utf-8 -*-

from gensim.models.doc2vec import TaggedDocument
from gensim.utils import simple_preprocess as preprocess
from gensim.models import Doc2Vec

file = open('text/data_merge.txt', 'r', encoding='utf-8')

trainings = [TaggedDocument(words = data.split(),tags = [i]) for i,data in enumerate(file)]

model = Doc2Vec(documents= trainings, size=400, min_count=10, iter=100)

model.save("doc2vec_merge.model")