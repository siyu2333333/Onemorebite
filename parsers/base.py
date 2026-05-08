import re
from typing import List
from abc import ABC, abstractmethod
from models import ProductRecord


def safe_get(d, *keys, default=None):
    for k in keys:
        if not isinstance(d, dict):
            return default
        d = d.get(k, default)
        if d is None:
            return default
    return d


def extract_grams(spec_str):
    if not spec_str:
        return None
    s = str(spec_str)
    m = re.search(r"(\d+(?:\.\d+)?)\s*克", s)
    if m:
        return float(m.group(1))
    m = re.search(r"(\d+(?:\.\d+)?)\s*g(?:\b|/|\s|$)", s, re.IGNORECASE)
    if m:
        return float(m.group(1))
    m = re.search(r"(\d+(?:\.\d+)?)\s*(?:千克|kg)", s, re.IGNORECASE)
    if m:
        return float(m.group(1)) * 1000
    return None


def calc_per_kg(price, spec_str):
    if not price or not spec_str:
        return None
    grams = extract_grams(spec_str)
    if grams and grams > 0:
        return round(price / (grams / 1000), 2)
    return None


def infer_storage_type(*texts):
    combined = " ".join(str(t) for t in texts if t).lower()
    if not combined:
        return ""
    if any(kw in combined for kw in ["冷冻", "冻", "-18", "freeze"]):
        return "冷冻"
    if any(kw in combined for kw in ["冷藏", "冰鲜", "0-4", "chill", "fresh"]):
        return "冰鲜"
    if any(kw in combined for kw in ["常温"]):
        return "常温"
    return ""


def infer_country(text):
    if not text:
        return ""
    t = str(text)
    if any(kw in t for kw in ["澳大利亚", "澳洲", "澳"]):
        return "澳洲"
    if "新西兰" in t:
        return "新西兰"
    if "美国" in t:
        return "美国"
    if "日本" in t:
        return "日本"
    if "阿根廷" in t:
        return "阿根廷"
    if "巴西" in t:
        return "巴西"
    domestic = ["中国", "国产", "山东", "北京", "上海", "广东", "四川", "河南",
                "内蒙", "新疆", "云南", "湖南", "湖北", "浙江", "江苏", "福建"]
    if any(kw in t for kw in domestic):
        return "国产"
    if len(t) <= 10 and t != "见产品外包装":
        return t
    return ""


CATEGORY_KEYWORDS = {
    "牛": ["牛", "beef", "牛排", "牛肉", "牛腩", "牛腱", "肥牛"],
    "猪": ["猪", "pork", "猪肉", "猪排", "五花"],
    "鸡": ["鸡", "chicken", "鸡肉", "鸡胸", "鸡腿", "鸡翅"],
    "羊": ["羊", "lamb", "羊肉", "羊排"],
    "鸭": ["鸭", "duck", "鸭肉"],
    "海鲜": ["鱼", "虾", "蟹", "贝", "鱿", "鲍", "海", "fish", "shrimp"],
}


def infer_category(*texts):
    combined = " ".join(str(t) for t in texts if t)
    for cat, kws in CATEGORY_KEYWORDS.items():
        if any(kw in combined for kw in kws):
            return cat
    return ""


_REGISTRY = []


def register_parser(cls):
    _REGISTRY.append(cls)
    return cls


def detect_parser(data):
    for parser_cls in _REGISTRY:
        if _key_exists(data, parser_cls.detect_keys):
            return parser_cls()
    raise ValueError(f"无法识别的JSON格式，可用解析器: {[p.__name__ for p in _REGISTRY]}")


def _key_exists(data, keys):
    for k in keys:
        parts = k.split(".")
        d = data
        for p in parts:
            if isinstance(d, dict) and p in d:
                d = d[p]
            else:
                return False
    return True


class BaseParser(ABC):
    platform_name: str = ""
    page_type: str = ""
    detect_keys: List[str] = []

    @abstractmethod
    def parse(self, data: dict) -> List[ProductRecord]:
        ...
