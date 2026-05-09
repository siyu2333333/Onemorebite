from playwright.sync_api import sync_playwright

from parsers.base import BaseParser, register_parser, safe_get, calc_per_kg, infer_storage_type, infer_category, infer_country
from models import ProductRecord


@register_parser
class RTMartDetailParser(BaseParser):
    platform_name = "大润发"
    page_type = "detail"
    detect_keys = ["body.productDetail"]

    def parse(self, data):
        detail = safe_get(data, "body", "productDetail", default={})
        if not detail:
            return []

        r = ProductRecord()
        r.platform = self.platform_name
        r.sku_id = str(detail.get("commodityNum", detail.get("goodsNo", "")))

        r.product_name = detail.get("itName", "").strip()
        r.price = detail.get("sm_price", detail.get("smPriceBg", 0))
        if isinstance(r.price, str):
            r.price = float(r.price) if r.price else 0.0

        unit = detail.get("priceUnit", detail.get("unit", ""))
        spec_desc = detail.get("specDesc", "")
        multi = detail.get("multiGoodsInfo", {})
        multi_desc = multi.get("rightDesc", "") if isinstance(multi, dict) else ""

        if spec_desc and unit and spec_desc != unit:
            r.spec = spec_desc
        else:
            r.spec = unit
        if multi_desc and "共" in multi_desc:
            r.spec = f"{r.spec}（{multi_desc}）" if r.spec else multi_desc

        props = detail.get("property", [])
        prop_map = {p.get("name", ""): p.get("value", "") for p in props if isinstance(p, dict)}
        r.brand = prop_map.get("品牌", "")
        r.storage_type = infer_storage_type(
            prop_map.get("保存条件", ""),
            detail.get("detailStorageFlag", ""),
            r.product_name
        )
        r.country = infer_country(prop_map.get("产地", ""))

        ai_points = detail.get("sellingAiPoint", [])
        if ai_points:
            r.selling_points = " | ".join(ai_points)

        left_tags = detail.get("goodsTitleLeftTag", [])
        if left_tags:
            tag_str = " / ".join(left_tags)
            if r.selling_points:
                r.selling_points = f"{tag_str} | {r.selling_points}"
            else:
                r.selling_points = tag_str

        selling_point = detail.get("sellingPoint", "").strip()
        if selling_point:
            if r.selling_points:
                r.selling_points = f"{r.selling_points} | {selling_point}"
            else:
                r.selling_points = selling_point

        good_detail_url = detail.get("goodDetailURL", "")
        if good_detail_url:
            r.image_urls = _fetch_detail_images(good_detail_url)

        valid_info = detail.get("validProductInfo", {})
        if isinstance(valid_info, dict):
            r.shelf_life = valid_info.get("expirationDate", "")

        r.category = infer_category(r.product_name, detail.get("cpSeq", ""))

        return [r]


def _fetch_detail_images(detail_url, timeout=30000):
    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            page = browser.new_page()
            page.goto(detail_url, wait_until="networkidle", timeout=timeout)
            page.wait_for_timeout(2000)

            img_urls = page.evaluate("""() => {
                const imgs = document.querySelectorAll('img');
                return Array.from(imgs).map(i => i.src).filter(s => s);
            }""")

            browser.close()

            seen = set()
            unique = []
            for u in img_urls:
                if u not in seen:
                    seen.add(u)
                    unique.append(u)
            return unique
    except Exception:
        return []
