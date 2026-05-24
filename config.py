import os
from pathlib import Path


def _load_env():
    env_path = Path(__file__).resolve().parent / ".env"
    if not env_path.exists():
        return
    with open(env_path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value


_load_env()

DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")
DEEPSEEK_BASE_URL = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com/v1")
DEEPSEEK_MODEL = os.getenv("DEEPSEEK_MODEL", "deepseek-v4-flash")

QWEN_API_KEY = os.getenv("QWEN_API_KEY")
QWEN_BASE_URL = os.getenv("QWEN_BASE_URL", "https://dashscope.aliyuncs.com/compatible-mode/v1")
QWEN_MODEL = os.getenv("QWEN_MODEL", "qwen3-vl-flash")

INPUT_DIR = "data/crawljson"
IMAGE_DIR = "data/images"
OUTPUT_DIR = "data/output"

DOWNLOAD_WORKERS = 5
DOWNLOAD_RETRIES = 3
VISION_RETRIES = 2

OUTPUT_COLUMNS = [
    "冷冻/冰鲜",
    "品类",
    "菜式",
    "烹饪方式",
    "场景",
    "建议产品名称",
    "产品形态\n(排，块，丝，粒，片，肉糜)",
    "使用原料",
    "原料最新价kg(元)",
    "经销商毛利",
    "建议厚度 (cm)",
    "主推季节",
    "营销场景",
    "营销卖点/备注",
    "竞品渠道",
    "类似产品品名",
    "国别",
    "品牌",
    "规格",
    "售价",
    "折合每kg(元)",
    "价格带",
]

IMAGE_EXTRACT_FIELDS = {
    "product_form": "产品形态(排/块/丝/粒/片/肉糜/末/整只/段/其他)",
    "raw_material": "使用的原料部位名称(如眼肉、板腱、牛腱等)",
    "cooking_method": "推荐烹饪方式(如煎/炒/涮/烤/煮/炖等)",
    "storage_hint": "储存条件提示(冷冻/冰鲜/常温，从图片包装上的储存说明判断)",
    "thickness_hint": "产品厚度提示(从图片比例估算，单位cm，如0.2、0.5、2等)",
}

INFERENCE_MODEL = os.getenv("INFERENCE_MODEL", "deepseek-v4-flash")
INFERENCE_BASE_URL = os.getenv("INFERENCE_BASE_URL", "https://api.deepseek.com/v1")
INFERENCE_API_KEY = os.getenv("INFERENCE_API_KEY")
