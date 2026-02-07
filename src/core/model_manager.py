"""模型缓存管理 - 管理 BAAI/bge-base-zh-v1.5 模型的下载和缓存"""

import os
import sys
import logging

logger = logging.getLogger(__name__)

MODEL_NAME = "BAAI/bge-base-zh-v1.5"
APP_NAME = "VLookupPro"


def get_cache_dir():
    """获取模型缓存目录，检查多个可能的缓存位置"""
    # 候选缓存路径
    candidates = []

    # 1. 应用专属缓存目录
    if sys.platform == "darwin":
        app_cache = os.path.join(os.path.expanduser("~"), "Library", "Caches", APP_NAME, "models")
    elif sys.platform == "win32":
        app_cache = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), APP_NAME, "models")
    else:
        app_cache = os.path.join(os.path.expanduser("~"), ".cache", APP_NAME, "models")
    candidates.append(app_cache)

    # 2. sentence-transformers 默认缓存
    st_cache = os.path.join(os.path.expanduser("~"), ".cache", "torch", "sentence_transformers")
    candidates.append(st_cache)

    # 3. huggingface hub 缓存
    hf_cache = os.path.join(os.path.expanduser("~"), ".cache", "huggingface", "hub")
    candidates.append(hf_cache)

    # 检查是否已有缓存的模型
    model_dir_name = MODEL_NAME.replace("/", "_")
    for cache_dir in candidates:
        model_path = os.path.join(cache_dir, model_dir_name)
        if os.path.exists(model_path):
            logger.info(f"找到已缓存模型: {model_path}")
            return cache_dir

    # 没有找到缓存，使用应用专属目录
    os.makedirs(app_cache, exist_ok=True)
    logger.info(f"使用缓存目录: {app_cache}")
    return app_cache


def load_model():
    """加载 sentence-transformers 模型"""
    from sentence_transformers import SentenceTransformer

    cache_dir = get_cache_dir()
    logger.info(f"正在加载模型 {MODEL_NAME} ...")
    model = SentenceTransformer(MODEL_NAME, cache_folder=cache_dir)
    logger.info("模型加载完成")
    return model
