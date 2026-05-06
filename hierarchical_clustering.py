# -*- coding: utf-8 -*-
"""
访问控制技术 — 凝聚层次聚类算法实现
包含：手写实现（单链/全链/组平均）、sklearn/scipy 实现、可视化、距离矩阵
"""

import numpy as np
import matplotlib
matplotlib.use('Agg')  # 非交互后端，用于生成图片
import matplotlib.pyplot as plt
from scipy.cluster.hierarchy import dendrogram, linkage, fcluster
from scipy.spatial.distance import pdist, squareform
from sklearn.cluster import AgglomerativeClustering
import warnings
warnings.filterwarnings('ignore')

# ======================== 1. 数据生成 ========================

np.random.seed(42)
N = 25  # 对象数量
# 生成 3 类聚类数据，使结果更直观
X = np.vstack([
    np.random.randn(8, 2) * 0.6 + np.array([1.0, 1.0]),
    np.random.randn(9, 2) * 0.6 + np.array([5.0, 1.0]),
    np.random.randn(8, 2) * 0.6 + np.array([3.0, 4.5]),
])
labels_true = np.array([0]*8 + [1]*9 + [2]*8)

# ======================== 2. 手写凝聚层次聚类 ========================

def euclidean_dist(a, b):
    """计算两个点之间的欧几里得距离"""
    return np.sqrt(np.sum((a - b) ** 2))

def compute_distance_matrix(points):
    """计算初始距离矩阵（N×N）"""
    n = len(points)
    dist_mat = np.zeros((n, n))
    for i in range(n):
        for j in range(i + 1, n):
            d = euclidean_dist(points[i], points[j])
            dist_mat[i][j] = d
            dist_mat[j][i] = d
    return dist_mat

def cluster_distance(cluster_i, cluster_j, dist_mat, method='single'):
    """
    计算两个簇之间的距离
    method: 'single' (单链/min), 'complete' (全链/max), 'average' (组平均)
    """
    dists = []
    for p in cluster_i:
        for q in cluster_j:
            dists.append(dist_mat[p][q])
    if method == 'single':
        return np.min(dists)
    elif method == 'complete':
        return np.max(dists)
    elif method == 'average':
        return np.mean(dists)
    else:
        raise ValueError(f"未知方法: {method}")

def agglomerative_clustering(points, n_clusters=3, method='single'):
    """
    凝聚层次聚类（自底向上）
    参数:
        points: 数据点数组 (N, dim)
        n_clusters: 目标聚类数
        method: 簇间距离计算方法
    返回:
        clusters_history: 每次合并的记录 [(合并前簇数, 最短距离, 合并后新簇ID), ...]
        final_labels: 每个点的最终簇标签
        linkage_matrix: 类似 scipy linkage 格式的矩阵
        dist_mat: 初始距离矩阵
    """
    n = len(points)
    dist_mat = compute_distance_matrix(points)

    # 初始化：每个点为一个簇，簇ID = 点的索引
    # clusters 格式: {簇ID: [成员点索引列表]}
    clusters = {i: [i] for i in range(n)}
    cluster_size = {i: 1 for i in range(n)}
    next_id = n  # 新簇从 n 开始编号

    clusters_history = []
    linkage_rows = []

    while len(clusters) > n_clusters:
        active_ids = list(clusters.keys())
        min_dist = float('inf')
        best_pair = None

        # 寻找距离最近的两个簇
        for idx_i, ci in enumerate(active_ids):
            for cj in active_ids[idx_i + 1:]:
                d = cluster_distance(clusters[ci], clusters[cj], dist_mat, method)
                if d < min_dist:
                    min_dist = d
                    best_pair = (ci, cj)

        ci, cj = best_pair
        size_i, size_j = cluster_size[ci], cluster_size[cj]

        # 合并：创建新簇
        new_members = clusters[ci] + clusters[cj]
        clusters[next_id] = new_members
        cluster_size[next_id] = size_i + size_j

        # 删除旧簇
        del clusters[ci], clusters[cj]
        del cluster_size[ci], cluster_size[cj]

        # 记录
        clusters_history.append((len(active_ids), min_dist, next_id))
        linkage_rows.append([float(ci), float(cj), min_dist, float(len(new_members))])

        next_id += 1

    # 生成标签：每个点映射到其所在簇的标签 (0, 1, 2)
    final_labels = np.zeros(n, dtype=int)
    for label, (cid, members) in enumerate(clusters.items()):
        for p in members:
            final_labels[p] = label

    linkage_matrix = np.array(linkage_rows) if linkage_rows else np.empty((0, 4))
    return clusters_history, final_labels, linkage_matrix, dist_mat


# ======================== 3. 可视化函数 ========================

def plot_dendrogram(linkage_mat, method_name, ax):
    """绘制树状图"""
    dendrogram(linkage_mat, ax=ax, labels=[f'P{i}' for i in range(N)])
    ax.set_title(f'{method_name} - 层次聚类树状图', fontproperties='SimHei')
    ax.set_xlabel('数据点索引')
    ax.set_ylabel('距离')

def plot_clusters(X, labels, method_name, ax):
    """绘制聚类结果的散点图"""
    unique_labels = np.unique(labels)
    colors = plt.cm.Set2(np.linspace(0, 1, len(unique_labels)))
    for k, col in zip(unique_labels, colors):
        ax.scatter(X[labels == k, 0], X[labels == k, 1],
                   c=[col], label=f'簇 {k}', s=60, edgecolors='k', linewidth=0.5)
    ax.set_title(f'{method_name} - 聚类结果', fontproperties='SimHei')
    ax.legend()
    ax.grid(True, alpha=0.3)

def plot_combined(points, results_dict, save_path='clustering_results.png'):
    """组合大图：三种方法的树状图 + 散点图"""
    n_methods = len(results_dict)
    fig, axes = plt.subplots(n_methods, 2, figsize=(14, 5 * n_methods))

    for row, (method_name, (linkage_mat, labels)) in enumerate(results_dict.items()):
        ax_dendro = axes[row, 0] if n_methods > 1 else axes[0]
        ax_scatter = axes[row, 1] if n_methods > 1 else axes[1]
        # 树状图用 scipy 的 linkage 重新计算
        from scipy.cluster.hierarchy import linkage as scipy_linkage
        link_scipy = scipy_linkage(points, method=method_name.split('-')[0].strip().lower())
        plot_dendrogram(link_scipy, method_name, ax_dendro)
        plot_clusters(points, labels, method_name, ax_scatter)

    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"[√] 组合图已保存: {save_path}")


# ======================== 4. 主程序 ========================

def print_section(title):
    print("\n" + "=" * 60)
    print(f"  {title}")
    print("=" * 60)

if __name__ == '__main__':
    # 配置中文字体
    plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'DejaVu Sans']
    plt.rcParams['axes.unicode_minus'] = False

    # 数据概览
    print_section("1. 数据概览")
    print(f"共生成 {N} 个二维数据点，分为 3 个自然簇")
    print(f"簇0: {sum(labels_true==0)} 个点  |  簇1: {sum(labels_true==1)} 个点  |  簇2: {sum(labels_true==2)} 个点")
    print("\n前 5 个数据点：")
    for i in range(5):
        print(f"  P{i}: ({X[i][0]:.4f}, {X[i][1]:.4f})  — 真实标签: {labels_true[i]}")

    # ---- 距离矩阵 ----
    print_section("2. 初始距离矩阵（前8×8子矩阵）")
    dist_mat_initial = compute_distance_matrix(X)
    np.set_printoptions(precision=3, suppress=True, linewidth=120)
    print(dist_mat_initial[:8, :8])

    # ---- 手写实现三种方法 ----
    methods = {
        'single':  '单链法 (Single-linkage)',
        'complete': '全链法 (Complete-linkage)',
        'average':  '组平均法 (Average-linkage)',
    }

    hand_results = {}
    for key, name in methods.items():
        print_section(f"3.{list(methods.keys()).index(key)+1} 手写凝聚层次聚类 — {name}")
        history, labels, link_mat, _ = agglomerative_clustering(X, n_clusters=3, method=key)

        print(f"\n合并过程（共 {len(history)} 步）：")
        print(f"{'步骤':<6}{'剩余簇数':<10}{'最短距离':<14}{'新簇ID':<10}")
        print("-" * 50)
        for step, (n_clust, dist, new_id) in enumerate(history):
            print(f"{step+1:<6}{n_clust:<10}{dist:<14.4f}{new_id:<10.0f}")

        print(f"\n最终分类结果：")
        for k in range(3):
            members = np.where(labels == k)[0]
            print(f"  簇 {k}: {list(members)}")

        # 检查聚类纯度
        from collections import Counter
        for k in range(3):
            members = np.where(labels == k)[0]
            true_in_cluster = labels_true[members]
            counts = Counter(true_in_cluster)
            print(f"  簇 {k} 成员: {sorted(members)}, 真实分布: {dict(counts)}")

        hand_results[name] = (link_mat, labels)

    # ---- 用 scipy 的 linkage 生成树状图 ----
    print_section("4. Scipy 层次聚类")
    for key, name in methods.items():
        link_scipy = linkage(X, method=key, metric='euclidean')
        cluster_labels = fcluster(link_scipy, t=3, criterion='maxclust') - 1
        print(f"\n{name} — Scipy fcluster 分类结果：")
        for k in range(3):
            members = np.where(cluster_labels == k)[0]
            print(f"  簇 {k}: {list(members)}")

    # ---- sklearn 实现 ----
    print_section("5. Sklearn 层次聚类")
    for key, name in methods.items():
        # Sklearn 只支持部分 linkage
        if key in ['single', 'complete', 'average']:
            linkage_map = {'single': 'single', 'complete': 'complete', 'average': 'average'}
            agg = AgglomerativeClustering(n_clusters=3, linkage=linkage_map[key])
            sk_labels = agg.fit_predict(X)
            print(f"\n{name} — Sklearn 分类结果：")
            for k in range(3):
                members = np.where(sk_labels == k)[0]
                print(f"  簇 {k}: {list(members)}")

    # ---- 思考题：距离矩阵 ----
    print_section("6. 最短距离值汇总")
    print(f"{'方法':<20}{'合并步数':<10}{'最短距离变化'}")
    print("-" * 60)
    for key, name in methods.items():
        history, _, _, _ = agglomerative_clustering(X, n_clusters=3, method=key)
        dists = [h[1] for h in history]
        print(f"{name:<20}{len(history):<10}{[f'{d:.4f}' for d in dists]}")

    # ---- 生成可视化图 ----
    print_section("7. 生成可视化图表")

    # 图1：三种方法的组合（树状图 + 散点）
    fig, axes = plt.subplots(3, 2, figsize=(14, 15))

    for row, (key, name) in enumerate(methods.items()):
        link = linkage(X, method=key, metric='euclidean')
        dendrogram(link, ax=axes[row, 0], labels=[f'P{i}' for i in range(N)])
        axes[row, 0].set_title(f'{name} — 树状图')
        axes[row, 0].set_xlabel('数据点')
        axes[row, 0].set_ylabel('距离')
        axes[row, 0].grid(axis='y', alpha=0.3)

        _, labels, _, _ = agglomerative_clustering(X, n_clusters=3, method=key)
        unique = np.unique(labels)
        for k in unique:
            axes[row, 1].scatter(X[labels == k, 0], X[labels == k, 1],
                                 s=60, edgecolors='k', linewidth=0.5, label=f'簇 {k}')
        axes[row, 1].set_title(f'{name} — 聚类结果')
        axes[row, 1].legend()
        axes[row, 1].grid(True, alpha=0.3)

    plt.tight_layout()
    plt.savefig('clustering_all_methods.png', dpi=150, bbox_inches='tight')
    plt.close()
    print("[√] clustering_all_methods.png — 三种方法树状图+散点图")

    # 图2：手写算法 vs sklearn 对比
    fig, axes = plt.subplots(2, 3, figsize=(18, 10))
    for col, (key, name) in enumerate(methods.items()):
        # 手写
        _, labels_h, _, _ = agglomerative_clustering(X, n_clusters=3, method=key)
        for k in np.unique(labels_h):
            axes[0, col].scatter(X[labels_h == k, 0], X[labels_h == k, 1],
                                 s=60, edgecolors='k', linewidth=0.5, label=f'C{k}')
        axes[0, col].set_title(f'手写 {name}')
        axes[0, col].legend(fontsize=8)
        axes[0, col].grid(alpha=0.3)

        # sklearn
        linkage_map = {'single': 'single', 'complete': 'complete', 'average': 'average'}
        agg = AgglomerativeClustering(n_clusters=3, linkage=linkage_map[key])
        sk_labels = agg.fit_predict(X)
        for k in np.unique(sk_labels):
            axes[1, col].scatter(X[sk_labels == k, 0], X[sk_labels == k, 1],
                                 s=60, edgecolors='k', linewidth=0.5, label=f'C{k}')
        axes[1, col].set_title(f'Sklearn {name}')
        axes[1, col].legend(fontsize=8)
        axes[1, col].grid(alpha=0.3)

    plt.tight_layout()
    plt.savefig('hand_vs_sklearn.png', dpi=150, bbox_inches='tight')
    plt.close()
    print("[√] hand_vs_sklearn.png — 手写 vs Sklearn 对比")

    # 图3：手写 vs Scipy 对比
    fig, axes = plt.subplots(2, 3, figsize=(18, 10))
    for col, (key, name) in enumerate(methods.items()):
        # 手写
        _, labels_h, _, _ = agglomerative_clustering(X, n_clusters=3, method=key)
        for k in np.unique(labels_h):
            axes[0, col].scatter(X[labels_h == k, 0], X[labels_h == k, 1],
                                 s=60, edgecolors='k', linewidth=0.5, label=f'C{k}')
        axes[0, col].set_title(f'手写 {name}')
        axes[0, col].legend(fontsize=8)
        axes[0, col].grid(alpha=0.3)

        link_s = linkage(X, method=key)
        sc_labels = fcluster(link_s, t=3, criterion='maxclust') - 1
        for k in np.unique(sc_labels):
            axes[1, col].scatter(X[sc_labels == k, 0], X[sc_labels == k, 1],
                                 s=60, edgecolors='k', linewidth=0.5, label=f'C{k}')
        axes[1, col].set_title(f'Scipy {name}')
        axes[1, col].legend(fontsize=8)
        axes[1, col].grid(alpha=0.3)

    plt.tight_layout()
    plt.savefig('hand_vs_scipy.png', dpi=150, bbox_inches='tight')
    plt.close()
    print("[√] hand_vs_scipy.png — 手写 vs Scipy 对比")

    # 图4：距离矩阵热力图
    fig, ax = plt.subplots(figsize=(10, 8))
    im = ax.imshow(dist_mat_initial, cmap='YlOrRd', aspect='auto')
    ax.set_title('数据点距离矩阵热力图')
    ax.set_xlabel('数据点索引')
    ax.set_ylabel('数据点索引')
    plt.colorbar(im, ax=ax, label='欧几里得距离')
    plt.tight_layout()
    plt.savefig('distance_matrix.png', dpi=150, bbox_inches='tight')
    plt.close()
    print("[√] distance_matrix.png — 距离矩阵热力图")

    print("\n" + "=" * 60)
    print("  实验完成！所有图表已生成。")
    print("=" * 60)
