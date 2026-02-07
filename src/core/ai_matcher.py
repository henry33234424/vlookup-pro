"""AI 语义匹配 - 使用 BAAI/bge-base-zh-v1.5 进行向量相似度匹配"""

import threading
import logging
import numpy as np

logger = logging.getLogger(__name__)

_model = None
_model_lock = threading.Lock()


def _get_model():
    """懒加载模型（线程安全，双重检查锁）"""
    global _model
    if _model is None:
        with _model_lock:
            if _model is None:
                from .model_manager import load_model
                _model = load_model()
    return _model


def encode_texts(texts):
    """将文本列表编码为归一化向量"""
    model = _get_model()
    vectors = model.encode(texts, normalize_embeddings=True, show_progress_bar=False)
    return np.array(vectors)


def compute_similarity_matrix(vectors_a, vectors_b):
    """计算余弦相似度矩阵（向量已归一化，直接点积）"""
    return np.dot(vectors_a, vectors_b.T)


def greedy_match(sim_matrix, threshold, a_items, b_items):
    """贪心一对一匹配：按相似度降序，每项最多匹配一次

    Args:
        sim_matrix: shape (len(a_items), len(b_items)) 的相似度矩阵
        threshold: 相似度阈值
        a_items: A 表未匹配项列表
        b_items: B 表未匹配项列表

    Returns:
        list of (a_idx, b_idx, similarity) 匹配结果
    """
    if sim_matrix.size == 0:
        return []

    # 找所有超过阈值的候选对
    rows, cols = np.where(sim_matrix >= threshold)
    scores = sim_matrix[rows, cols]

    # 按相似度降序排序
    order = np.argsort(-scores)
    rows = rows[order]
    cols = cols[order]
    scores = scores[order]

    matched_a = set()
    matched_b = set()
    matches = []

    for r, c, s in zip(rows, cols, scores):
        if r in matched_a or c in matched_b:
            continue
        matches.append((int(r), int(c), float(s)))
        matched_a.add(r)
        matched_b.add(c)

    return matches


def ai_match(a_texts, b_texts, threshold=0.85, progress_callback=None):
    """对两组文本进行 AI 语义匹配

    Args:
        a_texts: A 表待匹配文本列表
        b_texts: B 表待匹配文本列表
        threshold: 相似度阈值
        progress_callback: 进度回调 fn(message)

    Returns:
        list of (a_idx, b_idx, similarity)
    """
    if not a_texts or not b_texts:
        return []

    if progress_callback:
        progress_callback("正在加载 AI 模型...")

    if progress_callback:
        progress_callback(f"正在编码 A 表文本 ({len(a_texts)} 项)...")
    vectors_a = encode_texts(a_texts)

    if progress_callback:
        progress_callback(f"正在编码 B 表文本 ({len(b_texts)} 项)...")
    vectors_b = encode_texts(b_texts)

    if progress_callback:
        progress_callback("正在计算相似度矩阵...")
    sim_matrix = compute_similarity_matrix(vectors_a, vectors_b)

    if progress_callback:
        progress_callback("正在执行贪心匹配...")
    matches = greedy_match(sim_matrix, threshold, a_texts, b_texts)

    return matches
