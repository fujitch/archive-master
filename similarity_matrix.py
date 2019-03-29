# -*- coding: utf-8 -*-

from gensim.models import Doc2Vec
import numpy as np
import random
import math
import pickle

model = Doc2Vec.load('doc2vec_merge.model')

# 類似度行列
similarity_matrix = np.zeros((92, 92))

for i in range(92):
    for k in range(92):
        similarity_matrix[i, k] = round(model.docvecs.similarity(i, k), 3)

# 類似順位行列
similarity_ranking_matrix = np.zeros((92, 92))
for i in range(92):
    sorted_index = np.argsort(-similarity_matrix[i, :])
    for k in range(len(sorted_index)):
        index = sorted_index[k]
        similarity_ranking_matrix[i, index] = k
        
# k平均法
def k_means(obj_list, cluster_num, iteration=1000):
    ind_samples = random.sample(range(len(obj_list)), cluster_num)
    # 中心点初期状態作成
    center_point_list = []
    for ind in ind_samples:
        center_point_list.append(obj_list[ind])
    
    # 対象となる点群をクラスタリング
    clustered_point_dict = {}
    
    for i in range(iteration):
        # クラスタ分類を初期化
        for k in range(cluster_num):
            clustered_point_dict[k] = []
        # obj一個ずつクラスタ分類
        for obj in obj_list:
            # 中心点との距離を保存
            distances = np.zeros((cluster_num))
            for num in range(cluster_num):
                distance = np.sqrt(sum(np.square(obj - center_point_list[num])))
                distances[num] = distance
            clustered_point_dict[np.argmin(distances)].append(obj)
        # 中心点を更新
        center_point_list = []
        for ind in range(cluster_num):
            clustered_point_list = clustered_point_dict[ind]
            center_point = np.zeros((400))
            for l in range(len(clustered_point_list)):
                center_point = center_point + clustered_point_list[l]
            center_point = center_point / (l + 1)
            center_point_list.append(center_point)
    
    """
    # 一様分布の点を作成
    min_coords = np.zeros((len(obj_list[0])))
    max_coords = np.zeros((len(obj_list[0])))
    for i in range(len(obj_list)):
        obj = obj_list[i]
        if i == 0:
            min_coords[:] = obj
            max_coords[:] = obj
        for k in range(len(obj)):
            if obj[k] < min_coords[k]:
                min_coords[k] = obj[k]
            if obj[k] > max_coords[k]:
                max_coords[k] = obj[k]
    random_point_list = []
    for i in range(len(obj_list)):
        point = np.zeros((len(obj_list[0])))
        for k in range(len(point)):
            point[k] = np.random.uniform(min_coords[k], max_coords[k])
        random_point_list.append(point)
    """
    
    # 正規分布の点を作成
    obj_matrix = np.zeros((len(obj_list), len(obj_list[0])))
    for i in range(len(obj_list)):
        obj = obj_list[i]
        obj_matrix[i, :] = obj
    mean_obj_matrix = np.sum(obj_matrix, 0) / float(len(obj_list))
    for i in range(len(obj_list)):
        obj_matrix[i, :] = pow(obj_matrix[i, :] - mean_obj_matrix[i], 2)
    dispersion_obj_matrix = np.sum(obj_matrix, 0) / float(len(obj_list))
    random_point_list = []
    for i in range(len(obj_list)):
        point = np.zeros((len(obj_list[0])))
        for k in range(len(point)):
            point[k] = np.random.normal(mean_obj_matrix[k], dispersion_obj_matrix[k])
        random_point_list.append(point)
    
    # ギャップ統計量を計算
    clustered_sum_distance = 0.0
    random_sum_distance = 0.0
    for i in range(len(obj_list)):
        obj = obj_list[i]
        random_obj = random_point_list[i]
        distances = np.zeros((cluster_num))
        random_distances = np.zeros((cluster_num))
        for ind in range(cluster_num):
            distances[ind] = np.sqrt(sum(np.square(obj - center_point_list[ind])))
            random_distances[ind] = np.sqrt(sum(np.square(random_obj - center_point_list[ind])))
        clustered_sum_distance += np.min(distances)
        random_sum_distance += np.min(random_distances)
    gap_score = math.log(random_sum_distance) - math.log(clustered_sum_distance)
    
    return clustered_point_dict, gap_score
    
vectors_list_org = []
for i in range(92):
    vectors_list_org.append(model.docvecs[i])

gap_scores = np.zeros((90))
all_list = []
for i in range(90):
    print(i)
    clustered_point_dict, gap_score = k_means(vectors_list_org, i+2)
    gap_scores[i] = gap_score
    all_list.append(clustered_point_dict)

pickle.dump(gap_scores, open('gap_scores2.pickle', 'wb'))
pickle.dump(all_list, open('all_list2.pickle', 'wb'))