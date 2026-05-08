#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
朴朴平台商品 JSON 数据解析器
=================================
功能：将朴朴 App 商品详情 API 返回的 JSON 解析为结构化字典，
      并可批量输出到 Excel（对接竞品规划表格式）。

字段对应说明：
  JSON路径                                          → 目标字段
  ─────────────────────────────────────────────────────────────
  store_product_model.name                          → 类似产品品名
  store_product_model.spec                          → 规格
  store_product_model.price / 100                   → 售价（元）
  attribute_area[品牌]                              → 品牌
  attribute_area[存储条件]  → 推断                  → 冷冻/冰鲜
  attribute_area[保质期]                            → 保质期
  attribute_area[产地]      → 推断                  → 国别
  规格中提取 g 数 + 售价计算                         → 折合每kg(元)
  selling_point_area                                → 营销卖点
  comment_area                                      → 用户评价
  head_picture_area.img_list                        → 商品主图URL列表
"""

import json
import re
import pandas as pd
from pathlib import Path


# ─────────────────────────────────────────────
# 辅助函数
# ─────────────────────────────────────────────

def safe_get(d: dict, *keys, default=None):
    """安全多级取值，任一层不存在则返回 default"""
    cur = d
    for k in keys:
        if not isinstance(cur, dict):
            return default
        cur = cur.get(k, default)
        if cur is None:
            return default
    return cur


def parse_price_yuan(price_fen: int) -> float:
    """将价格从分转换为元（保留2位小数）"""
    return round(price_fen / 100, 2)


def calc_per_kg(price_yuan: float, spec_str: str) -> float | None:
    """
    根据规格字符串（如 '180g/份'）和售价计算折合每kg价格。
    返回 None 表示无法解析。
    """
    if not spec_str or not price_yuan:
        return None
    match = re.search(r"(\d+(?:\.\d+)?)\s*g", str(spec_str), re.IGNORECASE)
    if match:
        grams = float(match.group(1))
        return round(price_yuan / (grams / 1000), 2)
    return None


def infer_freeze_type(storage_text: str, sub_title: str = "") -> str:
    """
    根据存储条件 + 副标题推断冷冻/冰鲜。
    优先级：属性字段 > 副标题关键词
    """
    combined = str(storage_text) + str(sub_title)
    if any(kw in combined for kw in ["冷冻", "-18", "冻"]):
        return "冷冻"
    if any(kw in combined for kw in ["冷藏", "冰鲜", "0-4"]):
        return "冰鲜"
    return "未知"


def infer_country(origin_text: str) -> str:
    """标准化产地/国别"""
    t = str(origin_text)
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
    # 国内省市关键词 → 国产
    domestic_kw = ["国产", "山东", "北京", "上海", "广东", "四川", "河南",
                   "内蒙", "新疆", "云南", "湖南", "湖北", "浙江"]
    if any(kw in t for kw in domestic_kw):
        return "国产"
    return t  # 无法推断时保留原始值


# ─────────────────────────────────────────────
# 核心解析函数
# ─────────────────────────────────────────────

def parse_pupu_product(json_data: dict, platform: str = "朴朴") -> dict:
    """
    解析单个朴朴商品 JSON，返回结构化字典。

    Parameters
    ----------
    json_data : dict
        已 json.loads 后的字典
    platform : str
        平台名称标签，默认 "朴朴"

    Returns
    -------
    dict  包含所有目标字段的扁平化字典
    """
    # ── 校验入参 ──────────────────────────
    assert json_data.get("errcode") == 0, (
        f"API 返回错误: errcode={json_data.get('errcode')}, "
        f"errmsg={json_data.get('errmsg')}"
    )

    d  = json_data["data"]
    sm = d["show_model"]          # 前端展示模型
    sku = d["store_product_model"] # SKU 基础信息

    # ── 1. 商品基础信息 ───────────────────
    product_name = safe_get(sm, "product_title_area", "title", "text", default="")
    sub_title    = safe_get(sm, "product_title_area", "sub_title", "text", default="")
    spec         = sku.get("spec", "")
    product_id   = sku.get("product_id", "")
    store_pid    = sku.get("id", "")
    main_image   = sku.get("main_image", "")

    # ── 2. 价格（最可靠：store_product_model.price，单位分）──
    price_yuan   = parse_price_yuan(sku.get("price", 0))
    market_yuan  = parse_price_yuan(sku.get("market_price", 0))

    # 新人优惠展示价（前端字符串拼接）
    former = safe_get(sm, "price_area", "price_element", "former_price", "text", default="")
    latter = safe_get(sm, "price_area", "price_element", "latter_price", "text", default="")
    promo_price_str = f"{former}.{latter}" if former and latter else ""

    # ── 3. 商品属性（动态 key，转字典）──
    attr_list = safe_get(sm, "attribute_area", "attribute_list", default=[])
    attr_dict = {}
    for item in attr_list:
        k = safe_get(item, "title", "text", default="")
        v = safe_get(item, "content", "text", default="")
        if k:
            attr_dict[k] = v

    brand    = attr_dict.get("品牌", "")
    storage  = attr_dict.get("存储条件", "")
    shelf_life = attr_dict.get("保质期", "")
    origin   = attr_dict.get("产地", "")

    # ── 4. 派生字段 ───────────────────────
    per_kg      = calc_per_kg(price_yuan, spec)
    freeze_type = infer_freeze_type(storage, sub_title)
    country     = infer_country(origin)

    # ── 5. 卖点标签 ───────────────────────
    sell_points = [
        safe_get(sp, "content", "text", default="")
        for sp in safe_get(sm, "selling_point_area", "sell_point_list", default=[])
    ]
    sell_points_str = " / ".join(filter(None, sell_points))

    # ── 6. 商品主图列表 ───────────────────
    img_urls = [
        img.get("image_url", "")
        for img in safe_get(sm, "head_picture_area", "img_list", default=[])
    ]

    # ── 7. 评论区（展示层仅含1条摘要） ──
    comments_raw = safe_get(sm, "comment_area", "detail_list", default=[])
    comment_count_text = safe_get(sm, "comment_area", "suffix_title", "text", default="")
    sample_comments = []
    for c in comments_raw:
        sample_comments.append({
            "user"     : safe_get(c, "user_name", "text", default=""),
            "score"    : c.get("score", ""),
            "location" : safe_get(c, "time_and_location", "text", default=""),
            "content"  : safe_get(c, "content", "text", default=""),
            "tags"     : [safe_get(t, "content", "text", default="")
                          for t in c.get("user_tag_list", [])],
        })

    # ── 8. 库存 & 上架状态 ───────────────
    stock    = sku.get("stock_quantity", 0)
    is_valid = sku.get("is_effective", False)
    on_shelf = safe_get(d, "td", "product_on_shelf_status", default="")

    # ── 汇总结果 ──────────────────────────
    result = {
        # 竞品规划表对应字段
        "竞品渠道"       : platform,
        "类似产品品名"   : product_name,
        "品牌"           : brand,
        "规格"           : spec,
        "售价"           : price_yuan,
        "市场价"         : market_yuan,
        "折合每kg(元)"   : per_kg,
        "冷冻/冰鲜"      : freeze_type,
        "国别"           : country,
        "产地"           : origin,
        "保质期"         : shelf_life,
        "存储条件"       : storage,
        "营销卖点"       : sell_points_str,
        "副标题/描述"    : sub_title,
        # 扩展字段
        "新人优惠展示价" : promo_price_str,
        "商品主图_第1张" : img_urls[0] if img_urls else "",
        "全部主图URLs"   : " | ".join(img_urls),
        "评论数"         : comment_count_text,
        "评论摘要"       : sample_comments[0]["content"] if sample_comments else "",
        "评论用户"       : sample_comments[0]["user"] if sample_comments else "",
        "评论评分"       : sample_comments[0]["score"] if sample_comments else "",
        "评论标签"       : ", ".join(sample_comments[0]["tags"]) if sample_comments else "",
        # 系统字段
        "product_id"     : product_id,
        "store_product_id": store_pid,
        "库存"           : stock,
        "是否有效"       : is_valid,
        "上架状态"       : on_shelf,
    }
    return result


# ─────────────────────────────────────────────
# 批量解析入口（支持单文件 / 文件夹）
# ─────────────────────────────────────────────

def parse_file(filepath: str, platform: str = "朴朴") -> dict:
    """解析单个 JSON 文件"""
    with open(filepath, "r", encoding="utf-8") as f:
        data = json.load(f)
    return parse_pupu_product(data, platform=platform)


def parse_folder(folder: str, platform: str = "朴朴") -> pd.DataFrame:
    """
    批量解析文件夹内所有 .txt/.json 文件，返回 DataFrame。
    文件名格式建议：产品名.json 或 产品名.txt
    """
    records = []
    for fp in Path(folder).glob("*.txt"):
        try:
            records.append(parse_file(str(fp), platform=platform))
        except Exception as e:
            print(f"  ⚠ 跳过 {fp.name}：{e}")
    for fp in Path(folder).glob("*.json"):
        try:
            records.append(parse_file(str(fp), platform=platform))
        except Exception as e:
            print(f"  ⚠ 跳过 {fp.name}：{e}")
    return pd.DataFrame(records)


# ─────────────────────────────────────────────
# 主程序示例
# ─────────────────────────────────────────────

if __name__ == "__main__":
    # ── 单文件解析 ──
    result = parse_file("朴朴-牛肉片源数据.txt", platform="朴朴")

    print("\n" + "=" * 60)
    print("✅ 解析结果（竞品规划表关键字段）")
    print("=" * 60)
    key_fields = [
        "竞品渠道", "类似产品品名", "品牌", "规格",
        "售价", "折合每kg(元)", "冷冻/冰鲜", "国别",
        "产地", "保质期", "存储条件", "营销卖点",
        "评论数", "评论摘要",
    ]
    for k in key_fields:
        print(f"  {k:<15}: {result[k]}")

    # ── 输出到 Excel ──
    df = pd.DataFrame([result])
    df.to_excel("朴朴_解析结果.xlsx", index=False)
    print("\n📁 已保存到 朴朴_解析结果.xlsx")
