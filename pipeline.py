import json
import sys
from pathlib import Path
from datetime import datetime

import config
from models import ProductRecord
from parsers.base import detect_parser, calc_per_kg, infer_storage_type
from writers.excel import write_excel
from imaging.downloader import download_batch
from imaging.vision import extract_batch
from imaging.inference import infer_fields


def run(input_dir=None, output_dir=None):
    input_dir = Path(input_dir or config.INPUT_DIR)
    output_dir = Path(output_dir or config.OUTPUT_DIR)

    if not input_dir.exists():
        print(f"错误: 输入目录不存在 {input_dir}")
        sys.exit(1)

    json_files = sorted(input_dir.glob("*.*"))
    json_files = [f for f in json_files if f.suffix.lower() in (".txt", ".json")]
    if not json_files:
        print(f"错误: {input_dir} 中没有 .txt 或 .json 文件")
        sys.exit(1)

    print(f"扫描到 {len(json_files)} 个JSON文件")

    raw_records = []
    for fp in json_files:
        try:
            data = json.loads(fp.read_text(encoding="utf-8"))
            parser = detect_parser(data)
            records = parser.parse(data)
            for r in records:
                r.source_file = fp.name
            raw_records.extend(records)
            print(f"  {fp.name}: {type(parser).__name__} → {len(records)} 条记录")
        except ValueError as e:
            print(f"  [跳过] {fp.name}: {e}")
        except Exception as e:
            print(f"  [错误] {fp.name}: {e}")

    merged = _merge_records(raw_records)
    records = list(merged.values())

    for r in records:
        if not r.price_per_kg:
            result = calc_per_kg(r.price, r.spec) or calc_per_kg(r.price, r.product_name)
            if result:
                r.price_per_kg = result
        if not r.storage_type:
            r.storage_type = infer_storage_type(
                r.selling_points, r.product_name, r.shelf_life
            )
        if not r.price_band:
            pk = r.price_per_kg
            if pk:
                if pk < 50:
                    r.price_band = "低端"
                elif pk <= 150:
                    r.price_band = "中端"
                else:
                    r.price_band = "高端"

    all_urls = list(set(u for r in records for u in r.image_urls if u))
    if all_urls:
        url_map = download_batch(all_urls)
        for r in records:
            r.local_images = [url_map[u] for u in r.image_urls if u in url_map]

        if config.QWEN_API_KEY:
            print("视觉提取中...")
            for i, r in enumerate(records):
                if r.local_images:
                    fields = extract_batch(r.local_images)
                    r.product_form = fields.get("product_form", "") or ""
                    r.raw_material = fields.get("raw_material", "") or ""
                    if not r.storage_type:
                        hint = fields.get("storage_hint", "") or ""
                        if hint:
                            r.storage_type = infer_storage_type(hint)
                    if not r.thickness:
                        r.thickness = fields.get("thickness_hint", "") or ""
                    if (i + 1) % 5 == 0:
                        print(f"  进度: {i + 1}/{len(records)}")
            print("  视觉提取完成")
        else:
            print("跳过视觉提取(QWEN_API_KEY未设置)")
    else:
        print("无图片URL，跳过图片处理")

    if config.INFERENCE_API_KEY:
        print("LLM推理中...")
        for i, r in enumerate(records):
            infer_fields(r)
            if (i + 1) % 5 == 0:
                print(f"  进度: {i + 1}/{len(records)}")
        print("  LLM推理完成")
    else:
        print("跳过LLM推理(INFERENCE_API_KEY未设置)")

    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = output_dir / f"全产品规划_{timestamp}.xlsx"
    write_excel(records, str(out_path))

    print(f"\n完成: {len(records)} 条商品记录 → {out_path}")
    _print_summary(records)
    return records


def _merge_records(raw_records):
    merged = {}
    for r in raw_records:
        key = r.merge_key()
        if key in merged:
            _merge_into(merged[key], r)
        else:
            merged[key] = r
    return merged


def _merge_into(target, source):
    field_names = ProductRecord.field_names()
    skip_fields = {"image_urls", "local_images", "source_file"}
    for fname in field_names:
        if fname in skip_fields:
            if fname == "image_urls":
                existing = set(getattr(target, fname))
                new_urls = set(getattr(source, fname))
                setattr(target, fname, list(existing | new_urls))
            continue
        sv = getattr(source, fname)
        tv = getattr(target, fname)
        if not sv:
            continue
        if fname == "spec":
            from parsers.base import extract_grams
            if extract_grams(sv) and not extract_grams(tv):
                setattr(target, fname, sv)
            elif not tv:
                setattr(target, fname, sv)
        elif sv and not tv:
            setattr(target, fname, sv)


def _print_summary(records):
    with_storage = sum(1 for r in records if r.storage_type)
    print(f"  冷冻/冰鲜: {with_storage}/{len(records)}")
    with_form = sum(1 for r in records if r.product_form)
    print(f"  产品形态: {with_form}/{len(records)}")
    with_material = sum(1 for r in records if r.raw_material)
    print(f"  使用原料: {with_material}/{len(records)}")
    with_cuisine = sum(1 for r in records if r.cuisine_style)
    print(f"  菜式: {with_cuisine}/{len(records)}")
    with_scene = sum(1 for r in records if r.scene)
    print(f"  场景: {with_scene}/{len(records)}")
    with_season = sum(1 for r in records if r.season)
    print(f"  主推季节: {with_season}/{len(records)}")
    with_marketing = sum(1 for r in records if r.marketing_scenario)
    print(f"  营销场景: {with_marketing}/{len(records)}")
    with_price_band = sum(1 for r in records if r.price_band)
    print(f"  价格带: {with_price_band}/{len(records)}")


if __name__ == "__main__":
    run()
