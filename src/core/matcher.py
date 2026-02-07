"""匹配引擎 - 编排精确匹配 + AI 语义匹配"""

import logging
import pandas as pd

from .ai_matcher import ai_match

logger = logging.getLogger(__name__)


def load_first_column(filepath):
    """读取 Excel 文件第一列，返回字符串列表"""
    df = pd.read_excel(filepath, header=0)
    col = df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
    return col


def exact_match(a_items, b_items):
    """精确匹配：大小写不敏感 + strip，O(n) 哈希匹配

    Returns:
        matches: list of (a_idx, b_idx, similarity=1.0)
        unmatched_a_indices: list of int
        unmatched_b_indices: list of int
    """
    # 构建 B 表索引：normalized_text -> [original_indices]
    b_index = {}
    for i, text in enumerate(b_items):
        key = text.strip().lower()
        if key not in b_index:
            b_index[key] = []
        b_index[key].append(i)

    matches = []
    matched_a = set()
    matched_b = set()

    for a_idx, a_text in enumerate(a_items):
        key = a_text.strip().lower()
        if key in b_index:
            for b_idx in b_index[key]:
                if b_idx not in matched_b:
                    matches.append((a_idx, b_idx, 1.0))
                    matched_a.add(a_idx)
                    matched_b.add(b_idx)
                    break

    unmatched_a = [i for i in range(len(a_items)) if i not in matched_a]
    unmatched_b = [i for i in range(len(b_items)) if i not in matched_b]

    return matches, unmatched_a, unmatched_b


def run_match(file_a, file_b, threshold=0.85, progress_callback=None):
    """执行完整匹配流程

    Args:
        file_a: Excel A 文件路径
        file_b: Excel B 文件路径
        threshold: AI 匹配相似度阈值
        progress_callback: 进度回调 fn(message)

    Returns:
        dict with keys:
            matches: list of (a_text, b_text, similarity, status)
            a_items: 完整 A 表项目列表
            b_items: 完整 B 表项目列表
            unmatched_b: B 表未使用项的索引列表
    """
    if progress_callback:
        progress_callback("正在读取 Excel A ...")
    a_items = load_first_column(file_a)

    if progress_callback:
        progress_callback("正在读取 Excel B ...")
    b_items = load_first_column(file_b)

    if progress_callback:
        progress_callback(f"A 表 {len(a_items)} 项, B 表 {len(b_items)} 项")

    # Step 1: 精确匹配
    if progress_callback:
        progress_callback("正在执行精确匹配...")
    exact_matches, unmatched_a_indices, unmatched_b_indices = exact_match(a_items, b_items)

    if progress_callback:
        progress_callback(f"精确匹配: {len(exact_matches)} 对")

    # Step 2: AI 模糊匹配（仅未匹配项）
    ai_matches_raw = []
    if unmatched_a_indices and unmatched_b_indices:
        unmatched_a_texts = [a_items[i] for i in unmatched_a_indices]
        unmatched_b_texts = [b_items[i] for i in unmatched_b_indices]

        ai_matches_raw = ai_match(
            unmatched_a_texts,
            unmatched_b_texts,
            threshold=threshold,
            progress_callback=progress_callback
        )

    # 将 AI 匹配的局部索引映射回全局索引
    ai_matched_a = set()
    ai_matched_b = set()
    ai_matches = []
    for local_a, local_b, sim in ai_matches_raw:
        global_a = unmatched_a_indices[local_a]
        global_b = unmatched_b_indices[local_b]
        ai_matches.append((global_a, global_b, sim))
        ai_matched_a.add(global_a)
        ai_matched_b.add(global_b)

    if progress_callback:
        progress_callback(f"AI 匹配: {len(ai_matches)} 对")

    # 构建按 A 表原始顺序排列的结果
    # 先建索引映射
    a_match_info = {}
    for a_idx, b_idx, sim in exact_matches:
        a_match_info[a_idx] = (b_items[b_idx], sim, "精确匹配")
    for a_idx, b_idx, sim in ai_matches:
        a_match_info[a_idx] = (b_items[b_idx], sim, "模糊匹配")

    results = []
    for a_idx in range(len(a_items)):
        if a_idx in a_match_info:
            b_text, sim, status = a_match_info[a_idx]
            results.append((a_items[a_idx], b_text, sim, status))
        elif a_idx in set(unmatched_a_indices) - ai_matched_a:
            results.append((a_items[a_idx], "", 0.0, "未匹配"))

    # B 表未使用项
    still_unmatched_b = [i for i in unmatched_b_indices if i not in ai_matched_b]
    unmatched_b_items = [b_items[i] for i in still_unmatched_b]

    if progress_callback:
        progress_callback("匹配完成！")

    return {
        "matches": results,
        "a_items": a_items,
        "b_items": b_items,
        "unmatched_b": unmatched_b_items,
    }
