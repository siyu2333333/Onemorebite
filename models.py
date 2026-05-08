from dataclasses import dataclass, field, fields
from typing import List


@dataclass
class ProductRecord:
    sku_id: str = ""
    platform: str = ""
    product_name: str = ""
    brand: str = ""
    spec: str = ""
    price: float = 0.0
    price_per_kg: float = 0.0
    storage_type: str = ""
    category: str = ""
    country: str = ""
    selling_points: str = ""
    product_form: str = ""
    raw_material: str = ""
    cuisine_style: str = ""
    scene: str = ""
    suggested_name: str = ""
    raw_material_price: str = ""
    distributor_margin: str = ""
    thickness: str = ""
    season: str = ""
    marketing_scenario: str = ""
    price_band: str = ""
    image_urls: list = field(default_factory=list)
    local_images: list = field(default_factory=list)
    source_file: str = ""
    shelf_life: str = ""

    def merge_key(self):
        return f"{self.platform}:{self.sku_id}"

    def to_row(self, columns: List[str]) -> list:
        mapping = {
            "冷冻/冰鲜": self.storage_type,
            "品类": self.category,
            "菜式": self.cuisine_style,
            "场景": self.scene,
            "建议产品名称": self.suggested_name,
            "产品形态\n(排，块，丝，粒，片，肉糜)": self.product_form,
            "使用原料": self.raw_material,
            "原料最新价kg(元)": self.raw_material_price,
            "经销商毛利": self.distributor_margin,
            "建议厚度 (cm)": self.thickness,
            "主推季节": self.season,
            "营销场景": self.marketing_scenario,
            "营销卖点/备注": self.selling_points,
            "竞品渠道": self.platform,
            "类似产品品名": self.product_name,
            "国别": self.country,
            "品牌": self.brand,
            "规格": self.spec,
            "售价": self.price if self.price else "",
            "折合每kg(元)": self.price_per_kg if self.price_per_kg else "",
            "价格带": self.price_band,
        }
        return [mapping.get(c, "") for c in columns]

    @classmethod
    def field_names(cls):
        return [f.name for f in fields(cls)]
