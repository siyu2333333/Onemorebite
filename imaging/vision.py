import base64
import io
import json
import re
import time
from pathlib import Path

from openai import OpenAI
from PIL import Image

import config


def _build_client():
    return OpenAI(
        api_key=config.QWEN_API_KEY,
        base_url=config.QWEN_BASE_URL,
    )


def extract_batch(image_paths, fields_config=None):
    if not image_paths:
        return {}
    fields_config = fields_config or config.IMAGE_EXTRACT_FIELDS
    client = _build_client()

    for attempt in range(config.VISION_RETRIES + 1):
        try:
            b64_images = [_encode_image(p) for p in image_paths if Path(p).exists()]
            if not b64_images:
                return {}
            content = [{"type": "text", "text": _build_prompt(fields_config)}]
            for b64 in b64_images:
                content.append({
                    "type": "image_url",
                    "image_url": {"url": f"data:image/jpeg;base64,{b64}"}
                })

            response = client.chat.completions.create(
                model=config.QWEN_MODEL,
                messages=[{"role": "user", "content": content}],
                max_tokens=1000,
                temperature=0,
            )

            raw = response.choices[0].message.content.strip()
            m = re.search(r"\{.*\}", raw, re.DOTALL)
            return json.loads(m.group() if m else raw)

        except Exception:
            if attempt < config.VISION_RETRIES:
                time.sleep(2 ** attempt)

    return {}


def _encode_image(image_path):
    img = Image.open(image_path)
    if img.mode != "RGB":
        img = img.convert("RGB")

    w, h = img.size
    max_size = 1024
    if max(w, h) > max_size:
        scale = max_size / max(w, h)
        img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)

    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")


def _build_prompt(fields_config):
    fields_text = "\n".join(f"- {k}: {v}" for k, v in fields_config.items())
    return f"""请仔细分析图片中的商品信息，提取以下字段，严格以JSON格式输出：

{fields_text}

输出要求：
1. 只输出纯JSON，不要markdown代码块标记
2. 图片中找不到的字段，值填null
3. 所有字段值使用字符串类型"""
