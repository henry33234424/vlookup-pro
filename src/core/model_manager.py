"""模型缓存管理 - 管理 BAAI/bge-base-zh-v1.5 模型的下载和缓存"""

import os
import sys
import logging

logger = logging.getLogger(__name__)

MODEL_NAME = "BAAI/bge-base-zh-v1.5"
APP_NAME = "VLookupPro"


def _get_bundled_model_path():
    """检查是否有打包在 EXE 内的模型"""
    # PyInstaller 打包后 sys._MEIPASS 指向临时解压目录
    if hasattr(sys, '_MEIPASS'):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

    model_dir_name = MODEL_NAME.replace("/", "_")
    bundled_path = os.path.join(base, "models", model_dir_name)
    if os.path.exists(bundled_path) and os.path.isfile(os.path.join(bundled_path, "config.json")):
        logger.info(f"找到内置模型: {bundled_path}")
        return bundled_path
    return None


def get_cache_dir():
    """获取模型缓存目录，检查多个可能的缓存位置"""
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
    """加载 sentence-transformers 模型（优先使用内置模型）"""
    from sentence_transformers import SentenceTransformer

    # 优先使用打包在 EXE 内的模型
    bundled = _get_bundled_model_path()
    if bundled:
        logger.info(f"正在从内置路径加载模型: {bundled}")
        model = SentenceTransformer(bundled)
        logger.info("模型加载完成")
        return model

    # 回退到缓存/在线下载
    cache_dir = get_cache_dir()
    logger.info(f"正在加载模型 {MODEL_NAME} ...")
    model = SentenceTransformer(MODEL_NAME, cache_folder=cache_dir)
    logger.info("模型加载完成")
    return model
