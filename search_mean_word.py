# -*- coding: utf-8 -*-

import pickle
from gensim.models import Doc2Vec
from title_master import title_master

title_master_class = title_master()
title_dict = title_master_class.get_dict()
title_dict_keys = list(title_dict.keys())
model = Doc2Vec.load('doc2vec_merge.model')
cluster_num = 17

all_list = pickle.load(open('all_list.pickle', 'rb'))
clusters = all_list[cluster_num - 2]
word_title_set_list = []
for key in clusters:
    cluster = clusters[key]
    title_list = []
    sum_vector = None
    for i in range(len(cluster)):
        if i == 0:
            sum_vector = cluster[i]
        else:
            sum_vector = sum_vector + cluster[i]
        paper_tuple_list = model.docvecs.most_similar([ cluster[i] ], [], 1)
        paper_tuple = paper_tuple_list[0]
        paper_title = title_dict[title_dict_keys[paper_tuple[0]]]
        title_list.append(paper_title)
    vector = sum_vector / float(len(cluster))
    out = model.most_similar([ vector ], [], 5)
    word_title_set_list.append((out, title_list))