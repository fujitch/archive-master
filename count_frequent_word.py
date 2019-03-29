# -*- coding: utf-8 -*-

import glob
import codecs
import pickle
import MeCab

m = MeCab.Tagger(r'-Ochasen -d C:\Users\hori\workspace\encoder-decoder-sentence-chainer-master\mecab-ipadic-neologd')
text_path_list = glob.glob('text/wakati_filtered/*.txt')
for text_path in text_path_list:
    f = codecs.open(text_path, 'r', 'utf-8')
    word_list = f.read().split(' ')
    word_dict = {}
    for word in word_list:
        try:
            category = m.parse(word).split('\t')[3]
        except:
            continue
        """
        #　名詞、動詞、形容詞、形容動詞のみの数を数える
        if category.find('名詞') == -1 and category.find('動詞') == -1 and category.find('形容詞') == -1 and category.find('形容動詞') == -1:
            continue
        """
        # 数は省く
        if category.find('数') != -1:
            continue
        if word in word_dict:
            count = word_dict[word]
            word_dict[word] = count + 1
        else:
            word_dict[word] = 1
    fname = 'dictionary/' + text_path[5:-4] + '.pickle'
    pickle.dump(word_dict, open(fname, 'wb'))