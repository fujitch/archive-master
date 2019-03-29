# -*- coding: utf-8 -*-

import glob
import MeCab
import codecs

m = MeCab.Tagger(r'-Owakati -d C:\Users\hori\workspace\encoder-decoder-sentence-chainer-master\mecab-ipadic-neologd')
text_path_list = glob.glob('text/*.txt')
all_str = ""

for text_path in text_path_list:
    f = codecs.open(text_path, 'r', 'utf-8')
    text = f.read()
    text = text.replace(' ', '')
    text = text.replace('\n', '')
    text = m.parse(text)
    text = text.replace('\n', '')
    all_str = all_str + text + '\n'
    
all_str = all_str[:-1]

fout = open('text/data_merge.txt', 'w', encoding='utf-8')
fout.write(all_str)
fout.close()