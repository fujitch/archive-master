# -*- coding: utf-8 -*-

from eliminate_word_master import eliminate_word_master
import glob
import MeCab
import codecs
import pickle

eliminate_word_list = eliminate_word_master()
m_ocha = MeCab.Tagger(r'-Ochasen -d C:\Users\hori\workspace\encoder-decoder-sentence-chainer-master\mecab-ipadic-neologd')

text_path_list = glob.glob('text/*.txt')

not_include_list = []

for text_path in text_path_list:
    print(text_path)
    f = codecs.open(text_path, 'r', 'utf-8')
    text = f.read()
    text = text.replace(' ', '')
    text = text.replace('\n', '')
    node = m_ocha.parseToNode(text)
    word_count_dict = {}
    word_category_dict = {}
    while node:
        fields = node.feature.split(",")
        if fields[0] == '名詞' or fields[0] == '動詞' or fields[0] == '形容詞' or fields[0] == '形容動詞':
            word = node.surface
            if word in word_count_dict:
                count = word_count_dict[word]
                word_count_dict[word] = count + 1
            else:
                word_count_dict[word] = 1
            if word not in word_category_dict:
                word_category_dict[word] = fields[0]
        node = node.next
    sorted_word_list = []
    for k, v in sorted(word_count_dict.items(), key=lambda x: -x[1]):
        sorted_word_list.append(k)
    
    
    for i in range(len(sorted_word_list)):
        word = sorted_word_list[i]
        if i > 50:
            break
        if word in not_include_list:
            continue
        if eliminate_word_list.is_include(word):
            continue
        else:
            flg = input(word)
            if flg == "z":
                eliminate_word_list.add_word(word)
            else:
                not_include_list.append(word)

pickle.dump(not_include_list, open("not_include_list.pickle", "wb"))
eliminate_word_list.save_list("eliminate_word_list.pickle")