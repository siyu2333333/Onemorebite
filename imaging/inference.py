import json
import re
import time

from openai import OpenAI

import config


def _build_client():
    return OpenAI(
        api_key=config.INFERENCE_API_KEY,
        base_url=config.INFERENCE_BASE_URL,
    )


INFERENCE_PROMPT = """你是生鲜商品规划专家。根据给定的商品信息，推断以下缺失字段，严格以JSON格式输出。

已知商品信息:
{context}

请推断以下字段（无法推断的填null）:
- cuisine_style: 菜式(中餐/西餐/日韩/东南亚/零食/其他)
- scene: 场景(涮/烤/煎/炒/炖/煮/蒸/炸/零食/主食/煲汤/烘焙/其他)
- suggested_name: 建议产品名称(简洁规范的商品名，去掉促销前缀，品牌+品类+部位+形态+规格)
- main_season: 主推季节(全年/冬季/夏季/春秋)
- marketing_scenario: 营销场景(如高端火锅/家庭餐桌/户外烧烤/办公室零食/健身代餐等)
- thickness: 建议厚度(根据产品形态推断，单位cm，如肉片0.2-0.3，肉块2-3，肉粒1-1.5，肉丝0.3-0.5)

输出要求:
1. 只输出纯JSON，不要markdown代码块标记
2. 所有字段值使用字符串类型
3. 无法推断的填null"""


def _safe_parse_json(raw):
    raw = raw.strip()
    m = re.search(r"\{.*\}", raw, re.DOTALL)
    if m:
        raw = m.group()

    # Try direct parse first
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        pass

    # Fix truncated JSON: close unclosed strings and braces
    # Remove trailing incomplete key-value pairs
    fixed = raw.rstrip()
    while fixed.endswith(","):
        fixed = fixed[:-1]
    # Remove trailing incomplete key
    colon_pos = fixed.rfind('":')
    if colon_pos > 0:
        # Check if what follows is a valid value
        after = fixed[colon_pos + 2:]
        if after.strip() and not after.strip().startswith('"') and not after.strip() in ("null", "true", "false") and not after.strip()[0].isdigit():
            fixed = fixed[:fixed.rfind('"')]
    # Count braces and close
    open_braces = fixed.count("{") - fixed.count("}")
    fixed += "}" * open_braces
    try:
        return json.loads(fixed)
    except json.JSONDecodeError:
        pass

    # Fallback: extract individual fields via regex
    result = {}
    for key in ["cuisine_style", "scene", "suggested_name", "main_season", "marketing_scenario", "thickness"]:
        m = re.search(rf'"{key}"\s*:\s*"([^"]*)"', raw)
        if m:
            result[key] = m.group(1)
    return result


def infer_fields(record):
    known = {
        "product_name": record.product_name,
        "brand": record.brand,
        "category": record.category,
        "spec": record.spec,
        "price": record.price,
        "price_per_kg": record.price_per_kg,
        "storage_type": record.storage_type,
        "country": record.country,
        "selling_points": record.selling_points,
        "product_form": record.product_form,
        "raw_material": record.raw_material,
    }
    context = json.dumps(known, ensure_ascii=False, indent=2)
    prompt = INFERENCE_PROMPT.format(context=context)

    client = _build_client()
    for attempt in range(2):
        try:
            resp = client.chat.completions.create(
                model=config.INFERENCE_MODEL,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=800,
                temperature=0,
            )
            raw = resp.choices[0].message.content.strip()
            data = _safe_parse_json(raw)

            record.cuisine_style = data.get("cuisine_style") or ""
            record.scene = data.get("scene") or ""
            record.suggested_name = data.get("suggested_name") or ""
            record.season = data.get("main_season") or ""
            record.marketing_scenario = data.get("marketing_scenario") or ""
            record.thickness = data.get("thickness") or ""
            return
        except Exception as e:
            if attempt == 0:
                print(f"    [推理重试] {record.product_name[:30]}: {e}")
                time.sleep(2)
            else:
                print(f"    [推理失败] {record.product_name[:30]}: {e}")
