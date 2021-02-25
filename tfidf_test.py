#!/usr/bin/env python
#-*- coding:utf-8 -*-
import jieba
import jieba.analyse
import jieba.posseg as pseg

import nltk
import math
import string
import openpyxl
import json
# from　nltk.corpus import stopwords
from collections import Counter




def get_tokens(text):
    # lower = text.lower()
    # remove_punctuation_map = dict((ord(char), None) for char in string.punctuation)
    # no_punctuation = lower.translate(remove_punctuation_map)
    tokens = jieba.lcut(text,cut_all=True)

    return tokens


def tf(word, count):
    return count[word] / sum(count.values())

def n_containing(word, count_list):
    return sum(1 for count in count_list if word in count)

def idf(word, count_list):
    return math.log(len(count_list)) / (1 + n_containing(word, count_list))

def tfidf(word, count, count_list):
    return tf(word, count) * idf(word, count_list)

def avgDl(text,texts):
    # 返回文档长度和平均长度的比值
    avgL = sum(len(i) for i in texts)/len(texts)
    return abs(len(text))/avgL

def bm25tf(word, count,l,b=0.75,k=1.2):
    """ 使用bm25算法得到的tf """
    return ((k+1)*tf(word, count))/(k*(1.0-b+b*l)+tf(word, count))

def bm25tfidf(word, count, count_list,l,b=0.75,k=1.2):
    """ 使用bm25算法得到的tf-idf 
    k：这个参数控制着词频结果在词频饱和度中的上升速度。默认值为1.2。
    值越小饱和度变化越快，值越大饱和度变化越慢。
    b：这个参数控制着字段长归一值所起的作用，
    0.0会禁用归一化，1.0会启用完全归一化。默认值为0.75
    """
    
    return bm25tf(word, count,l,b,k) * idf(word, count_list)

def count_term(text):
    """ 
    返回一个counter类,形如
    Counter({'你好': 2, '中国': 2, '再见': 2})
     """
    tokens = get_tokens(text)
    # filtered = [w for w in tokens if not w in stopwords.words('english')]
    # stemmer = PorterStemmer()
    # stemmed = stem_tokens(filtered, stemmer)
    # 词干抽取，删除
    count = Counter(tokens)
    return count

def load_json(filename):
    """ 
    加载json格式,处理成为可用的文本语料数组
    返回值包含所有文本语料的数组和一个对应的编号数组
     """
    with open('sort_cluster_k_300.json','r',encoding='utf-8') as fp:
        json_data = json.load(fp)
    # 导入测试语料
    texts = []
    keyindex = []
    for data in json_data:
        text = ''
        for key,values in data.items():
            if key != 'num':            
                for i in values:
                    text = text+i
                keyindex.append(key)
        texts.append(text)
    return texts,keyindex

def main():
    filename = 'sort_cluster_k_300.json'
    texts,keyindex = load_json(filename)        
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = 'Sheet'
    wb.get_sheet_by_name('Sheet')
    countlist = []
    L = []
    for text in texts:
        countlist.append(count_term(text))
        L.append(avgDl(text,texts))

    for i, count in enumerate(countlist):
        print("Top words in document {}".format(i + 1))
        l = L[i]
        # 取得对应的文档长度与平均文档长度的比值
        # scores = {word: tfidf(word, count, countlist) for word in count}
        scores = {word: bm25tfidf(word, count, countlist,l,b=0.75,k=1.2) for word in count}
        sorted_words = sorted(scores.items(), key = lambda x: x[1], reverse=True)
        # print(sorted_words)
        # for word, score in sorted_words[:5]:
        #     print("\tWord: {}, TF-IDF: {}".format(word, round(score, 5)))
        sheet['A'+str(i+1)]=keyindex[i]
        try:
            sheet['B'+str(i+1)]=str(sorted_words[0])
            sheet['C'+str(i+1)]=str(sorted_words[1])
            sheet['D'+str(i+1)]=str(sorted_words[2])
            sheet['E'+str(i+1)]=str(sorted_words[3])
            sheet['F'+str(i+1)]=str(sorted_words[4])
        except IndexError:
            continue

    wb.save("result1.xlsx")   

        

if __name__ == "__main__":
    main()