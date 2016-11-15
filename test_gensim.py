#!/usr/bin/python
# -*- coding: gbk -*-

import gensim
import jieba
import sys
import pandas as pd
import glob


prefix_path = 'C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem'


def item_similarity(item, item_pool):
    jieba.load_userdict(prefix_path + '/PG dict.txt')

    num_item = len(item_pool)

    class MyCorpus(object):
        def __iter__(self):
            for seg in item_pool:
                yield list(jieba.cut(seg, cut_all=False))

    corp = MyCorpus()
    dictionary = gensim.corpora.Dictionary(corp)
    corpus = [dictionary.doc2bow(text) for text in corp]
    lsi = gensim.models.LsiModel(corpus, id2word=dictionary, num_topics=num_item)
    query = list(jieba.cut(item, cut_all=False))
    vec_bow = dictionary.doc2bow(query)
    vec_lsi = lsi[vec_bow]
    index = gensim.similarities.MatrixSimilarity(lsi[corpus])
    sims = index[vec_lsi]
    sims = sorted(enumerate(sims), key=lambda x: -x[1])

    return list(sims[x][0] for x in range(0, num_item) if sims[x][1] > 0.9)


def similarity(lst1, lst2, total_lst):

    jieba.load_userdict(prefix_path + '/PG dict.txt')
    num_item = len(lst2)

    class MyCorpus(object):
        def __iter__(self):
            for seg_lst in lst2:
                yield list(jieba.cut(seg_lst, cut_all=False))

    corp = MyCorpus()
    dictionary = gensim.corpora.Dictionary(corp)
    # dictionary.filter_extremes(no_below=1, no_above=1, keep_n=None)
    corpus = [dictionary.doc2bow(text) for text in corp]

    lsi = gensim.models.LsiModel(corpus, id2word=dictionary, num_topics=num_item)

    tfidf = gensim.models.TfidfModel(corpus)
    corpus_tfidf = tfidf[corpus]

    print tfidf

    query = list(jieba.cut(lst1[9], cut_all=False))

    # error!
    vec_bow = dictionary.doc2bow(query)

    vec_lsi = lsi[vec_bow]
    vec_tfidf = tfidf[vec_bow]

    index = gensim.similarities.MatrixSimilarity(lsi[corpus])

    sims = index[vec_lsi]

    print '/'.join(query)

    sims = sorted(enumerate(sims), key=lambda x: -x[1])
    print (sims)
    print lst2[sims[0][0]].decode('gbk')


def gen_pg_dict(lst):
    jieba.load_userdict(prefix_path + '/PG dict.txt')
    seg_list = list()
    unique_seg = list()
    for string in lst:
        seg_item = jieba.cut(string, cut_all=False)
        for term1 in seg_item:
            # print term1
            seg_list.append(term1)

    # for item in seg_list:
    #     print ('/'.join(item))
    for term2 in seg_list:
        if term2 not in unique_seg:
            unique_seg.append(term2)

    with open(prefix_path + '/pre_seg.txt', 'wb') as f:
        for term3 in unique_seg:
            # print term3
            f.write(term3 + '\n')


def main():
    reload(sys)
    sys.setdefaultencoding('gbk')

    path1 = prefix_path + '/DM Scorecard/WM 1501*.csv'
    path2 = prefix_path + '/DM Dictionary/WM 1501*.csv'

    ignore_cols = list([u'Unnamed', u'SFA Code', u'Store Name', u'Project Task'])

    total_lst = list()
    lst1 = list()
    lst2 = list()

    for path in glob.glob(path1):
        df1 = pd.read_csv(path, encoding='gbk')
        for col_name in df1.columns.values:
            if not any(s in col_name for s in ignore_cols):
                total_lst.append(col_name)
                lst1.append(col_name)
        # one loop test
        # break

    for path in glob.glob(path2):
        df2 = pd.read_csv(path, encoding='gbk')
        for col_name in df2['Item ']:
            total_lst.append(col_name)
            lst2.append(str(col_name))

        # one loop test
        # break

    # needs manually clean
    # one time job
    # gen_pg_dict(total_lst)

    # call gensim
    # similarity(lst1, lst2, total_lst)
    for num in item_similarity(lst1[9], lst2):
        print lst2[num].decode('gbk')

    print 'Done'

if __name__ == '__main__':
    main()
