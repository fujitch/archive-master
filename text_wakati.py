# -*- coding: utf-8 -*-

import MeCab
import glob
import codecs

m = MeCab.Tagger(r'-Owakati -d C:\Users\hori\workspace\encoder-decoder-sentence-chainer-master\mecab-ipadic-neologd')
m_ocha = MeCab.Tagger(r'-Ochasen -d C:\Users\hori\workspace\encoder-decoder-sentence-chainer-master\mecab-ipadic-neologd')

all_text = ""

text_path_list = glob.glob('text/*.txt')

for text_path in text_path_list:
    f = codecs.open(text_path, 'r', 'utf-8')
    text = f.read()
    text = text.replace(' ', '')
    text = text.replace('\n', '')
    
    fname = 'text/wakati/' + text_path[5:-4] + '.txt'
    f_out = codecs.open(fname, 'w', 'utf-8')
    f_out.write(m.parse(text))
    f_out.close()
    
    filtered_text = ""
    node = m_ocha.parseToNode(text)
    while node:
        fields = node.feature.split(",")
        """
        if fields[0] == '名詞' or fields[0] == '動詞' or fields[0] == '形容詞' or fields[0] == '形容動詞':
            filtered_text = filtered_text + node.surface + " "
        """
        if fields[0] == '名詞':
            filtered_text = filtered_text + node.surface + " "
        node = node.next
    fname_filtered = 'text/wakati_filtered/' + text_path[5:-4] + '.txt'
    f_out_filtered = codecs.open(fname_filtered, 'w', 'utf-8')
    f_out_filtered.write(filtered_text)
    f_out_filtered.close()
    
    all_text = all_text + m.parse(text)
    f.close()

"""   
f_out = codecs.open('text/merge_wakati.txt', 'w', 'utf-8')
f_out.write(all_text)
f_out.close()
"""