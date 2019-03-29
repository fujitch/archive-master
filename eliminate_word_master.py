# -*- coding: utf-8 -*-

import pickle

class eliminate_word_master():
    def __init__(self):
        self.eliminate_word_list = []
        
    def get_list(self):
        return self.eliminate_word_list
    
    def load_list(self, path):
        self.eliminate_word_list = pickle.load(open(path, "rb"))
    
    def is_include(self, word):
        return word in self.eliminate_word_list
        
    def add_word(self, word):
        self.eliminate_word_list.append(word)
        
    def save_list(self, path):
        pickle.dump(self.eliminate_word_list, open(path, "wb"))