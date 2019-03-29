# -*- coding: utf-8 -*-

import glob
import MeCab
import codecs
import pickle
from title_master import title_master
from eliminate_word_master import eliminate_word_master

criteria = 10

title_master_class = title_master()
title_dict = title_master_class.get_dict()
eliminate_word_list = eliminate_word_master()
eliminate_word_list.load_list("eliminate_word_list.pickle")

m_ocha = MeCab.Tagger(r'-Ochasen -d C:\Users\hori\workspace\encoder-decoder-sentence-chainer-master\mecab-ipadic-neologd')

text_path_list = glob.glob('text/*.txt')

archive_frequency_word = {}

for text_path in text_path_list:
    f = codecs.open(text_path, 'r', 'utf-8')
    text = f.read()
    text = text.replace(' ', '')
    text = text.replace('\n', '')
    node = m_ocha.parseToNode(text)
    word_count_dict = {}
    while node:
        fields = node.feature.split(",")
        word = node.surface
        if fields[0] == '名詞' or fields[0] == '動詞' or fields[0] == '形容詞' or fields[0] == '形容動詞':
            if not eliminate_word_list.is_include(word) and fields[1] != '数':
                if word in word_count_dict:
                    count = word_count_dict[word]
                    word_count_dict[word] = count + 1
                else:
                    word_count_dict[word] = 1
        node = node.next
    
    fname = text_path[5:-4]
    sorted_word_list = []
    sorted_num_list = []
    for k, v in sorted(word_count_dict.items(), key=lambda x: -x[1]):
        sorted_word_list.append(k)
        sorted_num_list.append(v)
    archive_frequency_word[fname] = sorted_word_list[:criteria]
    
pickle.dump(archive_frequency_word, open('archive_frequency_word_by' + str(criteria) + '.pickle', 'wb'))