import re
from parsers.base import BaseParser, register_parser, safe_get, calc_per_kg, infer_storage_type, infer_category, infer_country, extract_grams
from models import ProductRecord


@register_parser
class RTMartSearchParser(BaseParser):
    platform_name = "大润发"
    page_type = "search"
    detect_keys = ["body.MerchandiseList"]

    def parse(self, data):
        body = data.get("body", {})
        merch_list = body.get("MerchandiseList", [])
        cat_map = self._build_category_map(body.get("filter", []))

        records = []
        for item in merch_list:
            r = self._parse_item(item, cat_map)
            if r.sku_id:
                records.append(r)
        return records

    def _build_category_map(self, filters):
        cat_map = {}
        for f in filters:
            for child in safe_get(f, "children", default=[]):
                self._walk_cat(child, cat_map)
        return cat_map

    def _walk_cat(self, node, cat_map):
        cp_seq = node.get("cp_seq", "")
        cp_name = node.get("cp_name", "")
        if cp_seq and cp_name:
            cat_map[cp_seq] = cp_name
        for child in node.get("children", []):
            self._walk_cat(child, cat_map)

    def _parse_item(self, item, cat_map):
        r = ProductRecord()
        r.sku_id = str(item.get("sku_id", ""))
        r.platform = self.platform_name
        r.product_name = item.get("sm_name", "").strip()
        r.price = float(item.get("sm_price", 0))

        sale_unit = item.get("sale_unit", "")
        grams = extract_grams(r.product_name)
        if grams:
            r.spec = f"{int(grams) if grams == int(grams) else grams}g/{sale_unit}" if sale_unit else f"{int(grams) if grams == int(grams) else grams}g"
        else:
            r.spec = sale_unit

        r.selling_points = item.get("subtitle", "").strip()

        tags = [t.get("name", "") for t in item.get("items", []) if t.get("name")]
        tag_str = " / ".join(tags)
        if tag_str and r.selling_points:
            r.selling_points = f"{tag_str} | {r.selling_points}"
        elif tag_str:
            r.selling_points = tag_str

        pic = item.get("sm_pic", "")
        if pic:
            r.image_urls.append(pic)

        cp_seqs = item.get("cpSeqs", [])
        cat_names = []
        for seq in cp_seqs:
            name = cat_map.get(seq, "")
            if name:
                cat_names.append(name)
        r.category = infer_category(r.product_name, *cat_names)

        storage_url = item.get("storageConditionsFlag", "")
        if storage_url:
            r.storage_type = infer_storage_type(storage_url, r.product_name)

        r.shelf_life = ""
        return r
