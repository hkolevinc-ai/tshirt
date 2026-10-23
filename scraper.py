#!/usr/bin/env python3
"""TShirt24 -> the supplied Temu template. Python 3.11+, no browser required.

The price authority is Products!O, never Specials, VAT-excluded columns or HTML.
XLSX input is read only. Output keeps the original OOXML package and replaces
only Template sheet data rows; reserved rows and all dropdown sheets survive.
"""
from __future__ import annotations

import argparse
import concurrent.futures as futures
import copy
import csv
import hashlib
import itertools
import json
import logging
import math
import re
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
from lxml import html

BASE = "https://tshirt24.bg/"
LOG = logging.getLogger("tshirt24")
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
QN = lambda name: f"{{{NS}}}{name}"
BAD = re.compile(r"достав|куриер|еконт|спиди|speedy|econt|наложен платеж|плащан|връщан|върне|замяна|рекламаци|гаранци|работни дни|безплатн|отстъпк|промоци|поръч|телефон|\+359|https?://|www\.|shipping|delivery|warranty|returns|payment|цената|цена[ :]|евро|\bлв\b|\bEUR\b", re.I)
SET_RE = re.compile(r"комплект|за двойки|семейн|семейство|мама.{0,25}татко|родители и дете", re.I)
COLORS = {"бял":"White", "бялo":"White", "бяло":"White", "черен":"Black", "черна":"Black", "черно":"Black", "сив":"Grey", "сива":"Grey", "сиво":"Grey", "червен":"Red", "червена":"Red", "червено":"Red", "розов":"Pink", "розова":"Pink", "син":"Blue", "синя":"Blue", "синьо":"Blue", "зелен":"Green", "зелена":"Green", "жълт":"Yellow", "жълта":"Yellow", "лилав":"Purple", "лилава":"Purple", "бордо":"Burgundy", "оранжев":"Orange", "тъмно-син":"Navy Blue", "тъмносин":"Navy Blue", "тъмно син":"Navy Blue", "тъмно-синя":"Navy Blue", "тъмно-синьо":"Navy Blue"}
FIBERS = {"памук":"Cotton", "полиестер":"Polyester", "еластан":"Elastane", "вискоза":"Viscose", "полиамид":"Polyamide", "лен":"Linen", "акрил":"Acrylic", "вълна":"Wool", "cotton":"Cotton", "polyester":"Polyester", "elastane":"Elastane", "viscose":"Viscose"}


def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def plain(markup):
    try:
        return norm(html.fragment_fromstring(str(markup or ""), create_parent=True).text_content())
    except (ValueError, ET.ParseError):
        return norm(markup)


def clean_description(markup):
    root = html.fragment_fromstring(str(markup or ""), create_parent=True)
    for e in root.xpath(".//table | .//script | .//style | .//iframe | .//form"):
        e.drop_tree()
    for e in root.xpath(".//br"):
        e.tail = "\n" + (e.tail or "")
    for e in root.xpath(".//p | .//div | .//li"):
        e.tail = "\n" + (e.tail or "")
    lines = []
    for line in re.split(r"[\n\r]+|(?<=[!?])\s+|(?<=\.)\s+(?=[А-ЯA-Z])", root.text_content()):
        line = norm(line)
        if line and not BAD.search(line) and not re.search(r"^таблица|^размери\s*:?$", line, re.I):
            if line not in lines:
                lines.append(line)
    return "\n".join(lines)[:9500]


def original_image(value):
    value = str(value or "").strip().replace("\\/", "/")
    if not value:
        return ""
    u = urllib.parse.urlsplit(urllib.parse.urljoin(BASE, value))
    if u.hostname not in {"tshirt24.bg", "www.tshirt24.bg"}:
        return ""
    path = u.path
    if path.startswith("/image/cache/"):
        path = "/image/" + path[len("/image/cache/"):]
        path = re.sub(r"-\d+x\d+[a-z]*(?=\.[^.]+$)", "", path, flags=re.I)
    elif path.startswith("/catalog/"):
        path = "/image" + path
    elif not path.startswith("/image/"):
        return ""
    return urllib.parse.urlunsplit(("https", "tshirt24.bg", urllib.parse.quote(urllib.parse.unquote(path), safe="/(),-_."), "", ""))


def sheet_rows(wb, name):
    ws = wb[name]
    if hasattr(ws, "reset_dimensions"):
        ws.reset_dimensions()  # Temu supplied XML incorrectly declares A1:A1.
    return list(ws.iter_rows(values_only=True))


def read_export(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = sheet_rows(wb, "Products")
    if rows[0][14] != "price":
        raise ValueError("Products!O1 must be price. Refusing another price column.")
    products = {}
    for excel_row, row in enumerate(rows[1:], 2):
        if row[0] is None:
            continue
        p = dict(zip(rows[0], row))
        pid = str(int(row[0]))
        if pid in products:
            raise ValueError(f"Duplicate product_id {pid}")
        p.update(id=pid, price_o=row[14], excel_row=excel_row, images=[], export_options=[], option_values=[])
        products[pid] = p
    for row in sheet_rows(wb, "AdditionalImages")[1:]:
        if row[0] is not None and str(int(row[0])) in products:
            products[str(int(row[0]))]["images"].append(row[1])
    for row in sheet_rows(wb, "ProductSEOKeywords")[1:]:
        if row[0] is not None and str(int(row[0])) in products and row[2]:
            products[str(int(row[0]))]["slug"] = row[2]
    for name, field in [("ProductOptions", "export_options"), ("ProductOptionValues", "option_values")]:
        rs = sheet_rows(wb, name)
        for row in rs[1:]:
            if row[0] is not None and str(int(row[0])) in products:
                products[str(int(row[0]))][field].append(dict(zip(rs[0], row)))
    wb.close()
    return products


class Template:
    def __init__(self, path):
        self.path = Path(path)
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        rs = sheet_rows(wb, "Template")
        self.headers, self.keys = list(rs[1]), list(rs[3])
        self.columns = defaultdict(list)
        for i, name in enumerate(self.headers):
            if name:
                self.columns[str(name)].append(i)
        self.categories = {str(r[0]): str(r[1]) for r in sheet_rows(wb, "Category Name") if r[0]}
        self.dropdowns = {str(r[0]): [str(v) for v in r[2:] if v is not None] for r in sheet_rows(wb, "Dropdown Lists") if r[0]}
        self.required, self.disabled = {}, {}
        for r in sheet_rows(wb, "GoodsLevelMode")[1:]:
            if not r[0]:
                continue
            cid, kind = str(r[0]).split("_", 1)
            if kind in {"require", "disabled"}:
                target = self.required if kind == "require" else self.disabled
                target[cid] = {i for i, v in enumerate(r) if v == kind and i > 0}
        wb.close()

    def allowed(self, cid, field, group=3):
        return self.dropdowns.get(f"t_{group}_{cid}_{field}", [])

    def set(self, row, name, value, cid=None, occurrence=0):
        cols = self.columns.get(name, [])
        if occurrence < len(cols) and value is not None and value != "":
            i = cols[occurrence]
            if cid is None or i not in self.disabled.get(cid, set()):
                row[i] = value

    def choose(self, cid, field, preferred):
        allowed = self.allowed(cid, field)
        return next((x for x in preferred if x in allowed), None)

    def size(self, cid, label):
        raw = norm(label)
        alpha = re.fullmatch(r"(?:[2-6]XL|X{0,5}[SML]|S|M|L)", raw.upper())
        values = [raw]
        if alpha:
            values = [{"2XL":"XXL", "3XL":"XXXL"}.get(raw.upper(), raw.upper()), raw.upper()]
        elif re.search(r"\b(?:cm|см)\b", raw, re.I) or re.search(r"тениска\s+\d", raw, re.I):
            m = re.search(r"(?:боди\s*-?\s*|тениска\s+)?(\d{2,3})(?:\s*(?:cm|см)|\s*-)", raw, re.I)
            if m:
                values.insert(0, m[1])
        for family in ["2 - Regular Size", "12 - US Size", "101 - Custom size"]:
            if family not in self.allowed(cid, "Size Family", 4):
                continue
            subs = self.dropdowns.get(f"t_4_{cid}_{family}_Sub-Size Family", [])
            for sub in sorted(subs, key=lambda s: ("Alpha" not in s if alpha else "Numeric" not in s)):
                allowed = self.dropdowns.get(f"t_4_{cid}_{family}_{sub}_Size", [])
                for v in values:
                    if v in allowed:
                        return family, sub, v
        return None

    def missing(self, cid, row):
        result, groups = [], defaultdict(list)
        for i in self.required.get(cid, set()):
            if i in self.disabled.get(cid, set()):
                continue
            field = str(self.headers[i] or "")
            if " - Composition:" in field or " - Lining Ingredients:" in field:
                groups[field.split(":")[0]].append(i)
            elif field in {"12 - Material", "Detail Images URL", "SKU Images URL"}:
                groups[field].append(i)
            elif field == "List Price - EUR":
                na = self.columns.get("Not available for List price", [])
                if not row.get(i) and not any(row.get(c) == "N/A" for c in na):
                    result.append(field)
            elif row.get(i) is None or row.get(i) == "":
                result.append(field)
        for field, indexes in groups.items():
            if not any(row.get(i) is not None and row.get(i) != "" for i in indexes):
                result.append(field)
        return sorted(set(result))


def extract_page(body, expected_id):
    root = html.fromstring(body)
    ids = root.xpath('//input[@name="product_id"]/@value')
    if str(expected_id) not in ids:
        raise ValueError("Page is not the requested product (404, redirect or verification)")
    titles = root.xpath('//h1[contains(@class,"page-title")]')
    if not titles:
        raise ValueError("Missing product title")
    blocks = root.xpath('//*[contains(@class,"product-extra-description")]//*[contains(concat(" ",normalize-space(@class)," ")," block-content ")]')
    description = html.tostring(blocks[0], encoding="unicode") if blocks else ""
    groups = []
    holders = root.xpath('//*[contains(concat(" ",normalize-space(@class)," ")," product-options ")]//*[contains(concat(" ",normalize-space(@class)," ")," form-group ")]')
    for g in holders:
        labels = g.xpath('.//label[contains(@class,"control-label")]')
        label = norm(labels[0].text_content()) if labels else "Option"
        choices = []
        for node in g.xpath('.//select/option[@value!=""] | .//input[@type="radio"] | .//input[@type="checkbox"]'):
            if "disabled" in node.attrib:
                continue
            text = norm(node.text_content()) if node.tag == "option" else norm(node.get("title") or node.get("aria-label"))
            if not text and node.tag == "input":
                parents = node.xpath('ancestor::label[1]')
                text = norm(parents[0].text_content()) if parents else ""
            text = re.sub(r"\s*\([+-]\s*[€$\d].*?\)\s*$", "", text)
            if not node.get("value") or not text:
                continue
            choices.append({"id":node.get("value"), "label":text})
        free = bool(g.xpath('.//input[not(@type="radio") and not(@type="checkbox") and not(@type="hidden")] | .//textarea | .//input[@type="checkbox"]'))
        groups.append({"name":label, "required":"required" in g.get("class", "").split(), "choices":choices, "custom":free})
    images = []
    for node in root.xpath('//*[contains(@class,"lightgallery-product-images") and @data-images]'):
        for item in json.loads(node.get("data-images")):
            if item.get("type") == "image":
                images.append(item.get("src"))
    return {"id":str(expected_id), "title":norm(titles[0].text_content()), "description":description, "groups":groups, "images":images, "breadcrumbs":[norm(e.text_content()) for e in root.xpath('//ul[contains(@class,"breadcrumb")]/li')], "fetched_at":time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())}


class Fetcher:
    def __init__(self, cache, refresh=False, offline=False):
        self.cache, self.refresh, self.offline = Path(cache), refresh, offline
        self.cache.mkdir(parents=True, exist_ok=True)

    def get(self, p):
        path = self.cache / f"{p['id']}.json"
        if path.exists() and not self.refresh:
            return json.loads(path.read_text(encoding="utf-8"))
        if self.offline:
            raise ValueError("No cached live product page")
        urls = []
        if p.get("slug"):
            urls.append(BASE + urllib.parse.quote(str(p["slug"]), safe="/,-_."))
        urls.append(BASE + "index.php?route=product/product&product_id=" + p["id"])
        last = None
        for url in urls:
            for attempt in range(3):
                try:
                    request = urllib.request.Request(url, headers={"User-Agent":"TShirt24-Catalog-Exporter/1.0", "Accept":"text/html"})
                    with urllib.request.urlopen(request, timeout=25) as response:
                        data = extract_page(response.read(), p["id"])
                    data["url"] = url
                    temp = path.with_suffix(".tmp")
                    temp.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
                    temp.replace(path)
                    return data
                except (ValueError, OSError, urllib.error.URLError, json.JSONDecodeError) as e:
                    last = e
                    if isinstance(e, urllib.error.HTTPError) and e.code == 404:
                        break
                    time.sleep(min(2 ** attempt, 4))
        raise ValueError(f"Could not verify live product: {last}")


def classify(p, group_name="", category_override=None):
    if category_override:
        return str(category_override), None
    name = norm(p["name(bg)"]).lower()
    text = name + " " + norm(p.get("tags(bg)")).lower()
    if SET_RE.search(name):
        return None, "Mixed/couple/family set: needs a suitable set category or individual item prices"
    hood = bool(re.search(r"суи[тчш]+[ъеа]?р|качулк|hoodie", text))
    sweatshirt = bool(re.search(r"блуз|sweatshirt", name))
    polo = "поло" in text
    tank = "потник" in text
    baby = "боди" in name
    # Explicit selected size group distinguishes age/gender in mixed-size single garments.
    audience = (group_name + " " + name).lower()
    if baby:
        return ("26325" if re.search(r"момич|girl", name) else "26402"), None
    if re.search(r"детски размер|детск|момич|момче", audience):
        girl = bool(re.search(r"момич|girl", name))
        return ({"hood":"29690", "sweat":"29689", "polo":"29554", "tank":"29555", "tee":"29553"} if girl else {"hood":"30846", "sweat":"30847", "polo":"30842", "tank":"30841", "tee":"30843"})["hood" if hood else "sweat" if sweatshirt else "polo" if polo else "tank" if tank else "tee"], None
    female = "дамск" in group_name.lower() or ("мъжк" not in group_name.lower() and bool(re.search(r"дамск|жени|бремен|мама|баба", name)))
    if "бремен" in name:
        return "29427" if tank else "29426", None
    if female:
        return "29065" if hood else "29066" if sweatshirt else "29070" if polo else "29071" if tank else "29069", None
    if re.search(r"мъжк|unisex|унисекс|татко|дядо", audience):
        return "30467" if hood else "30466" if sweatshirt else "30472" if polo else "30475" if tank else "30469", None
    return None, "Audience/category ambiguous; add product category override"


def color_name(label, title):
    text = norm(label or title).lower()
    for bg in sorted(COLORS, key=len, reverse=True):
        if re.search(r"(?<![а-я])" + re.escape(bg) + r"(?![а-я])", text):
            return COLORS[bg]
    return re.sub(r"\s*\([A-Z0-9-]+\)\s*$", "", norm(label)) or None


def composition(markup):
    text = plain(markup).lower()
    found = {}
    for bg, en in FIBERS.items():
        for m in re.finditer(r"(\d+(?:[.,]\d+)?)\s*%\s*" + bg, text):
            found[en] = float(m[1].replace(",", "."))
    if found and abs(sum(found.values()) - 100) > 0.1:
        return {}  # Do not invent remaining fibers.
    return found


def size_measurements(markup, size):
    result = {}
    root = html.fragment_fromstring(str(markup or ""), create_parent=True)
    for table in root.xpath('.//table'):
        rows = [[norm(td.text_content()) for td in tr.xpath('./td|./th')] for tr in table.xpath('.//tr')]
        if not rows:
            continue
        headings = rows[0]
        idx = next((i for i, x in enumerate(headings) if x.upper() == str(size).upper() or re.match(r'^' + re.escape(str(size)) + r'\s*[-–]',x)), None)
        if idx is None:
            continue
        for row in rows[1:]:
            if len(row) <= idx or not re.fullmatch(r"\d+(?:[.,]\d+)?", row[idx]):
                continue
            value = float(row[idx].replace(",", "."))
            key = row[0].lower().replace("a", "а")
            if re.search(r"шир(оч|ин)|1/2.*(гръд|chest)", key):
                result["Chest"] = value * 2  # flat garment width -> circumference
            elif re.search(r"гръдн.*обик|chest circumference", key):
                result["Chest"] = value
            elif re.search(r"височин|дължин|length", key) and not re.search(r"ръкав|sleeve|ръст", key):
                result["Length"] = value
            elif re.search(r"ръкав|sleeve", key):
                result["Sleeve length"] = value
    return result


def variants(p, live, max_variants):
    groups = live["groups"]
    if p["export_options"] and not groups:
        raise ValueError("Export has options but no live options were found")
    if any(g["custom"] for g in groups):
        raise ValueError("Free text / checkbox personalization requires review")
    if any(g["required"] and not g["choices"] for g in groups):
        raise ValueError("Required option has no available choices")
    groups = [g for g in groups if g["choices"]]
    if any(not re.search(r"размер|цвят|size|color", g["name"], re.I) for g in groups):
        raise ValueError("Unmapped option type: " + ", ".join(g["name"] for g in groups))
    size_groups = [g for g in groups if re.search(r"размер|size", g["name"], re.I)]
    if len(size_groups) > 1:
        raise ValueError("Multiple garment size selectors; bundle needs category/pricing review")
    count = math.prod(len(g["choices"]) for g in groups)
    if count > max_variants:
        raise ValueError(f"{count} combinations exceed max_variants={max_variants}; no truncation was performed")
    for combo in itertools.product(*(g["choices"] for g in groups)):
        yield [(g["name"], choice) for g, choice in zip(groups, combo)]


def build_product(p, live, schema, config):
    pid = p["id"]
    override = config.get("product_overrides", {}).get(pid, {})
    price = p["price_o"]
    if not isinstance(price, (int, float)) or isinstance(price, bool) or not math.isfinite(price) or price <= 0:
        raise ValueError("Invalid numeric Products!O price")
    if SET_RE.search(p["name(bg)"]):
        raise ValueError("Bundle price in O is for the entire set; supplied template has no mixed-family-set category")
    description_html = p.get("description(bg)") or live.get("description") or ""
    description = clean_description(description_html)
    if not description and live.get("description"):
        description_html = live["description"]
        description = clean_description(description_html)
    description = description or norm(p["name(bg)"])
    images = list(dict.fromkeys(filter(None, (original_image(v) for v in [p.get("image"), *p["images"], *live["images"]]))))
    if not images:
        raise ValueError("Missing original product images")
    fibers = composition(description_html)
    all_rows = []
    for combo in variants(p, live, config["max_variants_per_product"]):
        size_items = [(label, item) for label, item in combo if re.search(r"размер|size", label, re.I)]
        color_items = [item for label, item in combo if re.search(r"цвят|color", label, re.I)]
        group_name, size_item = size_items[0] if size_items else ("", {"label":"one-size", "id":"one"})
        cid, reason = classify(p, group_name, override.get("category"))
        if cid not in schema.categories:
            raise ValueError(reason or f"Category {cid} is not in this template")
        size = schema.size(cid, size_item["label"])
        color = color_name(color_items[0]["label"] if color_items else None, p["name(bg)"])
        row = {}
        def put(name, value, occurrence=0):
            schema.set(row, name, value, cid, occurrence)
        put("Category", cid)
        put("Category Name", schema.categories[cid])
        put("Product type", "Normal product")
        put("Product Name", norm(p["name(bg)"])[:500])
        goods = f"T24-{pid}-{cid}"
        signature = "|".join(f"{label}:{item['id']}" for label, item in combo) or "one"
        sku = goods + "-" + hashlib.sha256(signature.encode()).hexdigest()[:12]
        put("Contribution Goods", goods)
        put("Contribution SKU", sku)
        put("Update or Add", "Add")
        put("Product Description", description)
        for i, image in enumerate(images[:10]):
            put("SKU Images URL", image, i)
        for i, image in enumerate(images[:50]):
            put("Detail Images URL", image, i)
        for fiber, percent in fibers.items():
            put("15 - Composition:" + fiber, percent)
        if fibers:
            material_candidates = ["Cotton" if len(fibers) == 1 and "Cotton" in fibers else "Cotton Blend", *fibers]
            material = schema.choose(cid, "12 - Material", material_candidates)
            put("12 - Material", material)
        if cid in {"26325", "26402"}:
            age = "0-3 Years Old"
        elif cid in {"29553","29554","29555","29689","29690","30841","30842","30843","30844","30846","30847"}:
            age = "12 and under"
        else:
            age = None
        if age in schema.allowed(cid,"1117 - Applicable Age Group"):
            put("1117 - Applicable Age Group",age)
        theme = schema.allowed(cid, "Variation Theme", 4)
        put("Variation Theme", theme[0] if len(theme) == 1 else "Color × Size" if "Color × Size" in theme else None)
        if size:
            family, sub, canonical_size = size
            put("Size Family", family)
            put("Sub-Size Family", sub)
            put("Size", canonical_size)
            put("EU size", canonical_size)
        else:
            put("Size", size_item["label"])
        put("Color", color)
        dims = size_measurements(description_html, size[2] if size else size_item["label"])
        put("Size Chart Method", "Add size chart manually")
        put("Unit", "cm")
        for field, indexes in schema.columns.items():
            if " - Tops - Product - " in field or " - girls clothing - Product - " in field or " - Clothing - Product - " in field:
                measurement = field.rsplit(" - ", 1)[-1]
                for index in indexes:
                    if index in schema.required.get(cid, set()) and index not in schema.disabled.get(cid, set()) and measurement in dims:
                        row[index] = dims[measurement]
        quantity = max(0, int(p.get("quantity") or 0))
        # Option stock cannot be joined by product_option_value_id: HTML IDs differ
        # from export option_value_id. Never pretend these IDs are equivalent.
        put("Quantity", quantity)
        put("Base Price - EUR", float(price))
        put("Not available for List price", "N/A")
        put("Reference Link", live.get("url", BASE + "index.php?route=product/product&product_id=" + pid))
        weight = float(p.get("weight") or 0)
        if weight > 0:
            unit = str(p.get("weight_unit") or "").lower()
            put("Weight - g", weight * (1000 if unit == "kg" else 1 if unit in {"g", "gr", "гр"} else 0) or None)
        else:
            put("Weight - g", override.get("weight_g", config.get("default_weight_g")))
        for field, source in [("Length - cm", "length"), ("Width - cm", "width"), ("Height - cm", "height")]:
            value = p.get(source)
            if value and float(value) > 0:
                factor = {"cm":1, "mm":0.1, "m":100}.get(str(p.get("length_unit")).lower())
                value = float(value) * factor if factor else None
            else:
                value = override.get(source + "_cm", config.get("default_package_cm", {}).get(source))
            put(field, value)
        for field, value in [("SKU type","Single set"),("Individually packed",config.get("individually_packed")),("Total packaging quantity",1),("Packaging unit","piece")]:
            put(field,value)
        # Required operational template fields, never injected into descriptions.
        shipping = schema.dropdowns.get("t_7_Shipping Template", [])
        put("Shipping Template", config.get("shipping_template") or (shipping[0] if len(shipping)==1 else None))
        put("Handling Time", config.get("handling_time"))
        origin = override.get("country_of_origin", config.get("country_of_origin"))
        if not origin and re.search(r"произведен[аои]?\s+в\s+България|made\s+in\s+Bulgaria",plain(description_html),re.I):
            origin = "Bulgaria"
        put("Country/Region of Origin",origin)
        put("Manufacturer", override.get("manufacturer", config.get("manufacturer")))
        put("EU Responsible person", override.get("eu_responsible_person", config.get("eu_responsible_person")))
        put("Product Identification", override.get("product_identification", config.get("product_identification")))
        title = p["name(bg)"].lower()
        facts = {
            "29 - Sleeve Length": "Long Sleeve" if re.search(r"дъл[ъг]|блуз|суи[тчш]", title) else "Short Sleeve" if "тениск" in title else None,
            "21 - Collar Style": "Hooded" if re.search(r"суи[тчш]|качулк",title) else None,
        }
        for field, value in facts.items():
            if value in schema.allowed(cid, field):
                put(field,value)
        missing = schema.missing(cid, row)
        if not size:
            missing.append("Size is not in category dropdown: " + size_item["label"])
        if any(float(v.get("quantity") or 0) <= 0 and str(v.get("subtract")).lower() == "true" for v in p["option_values"]):
            missing.append("Option stock needs exact ID/name join; an export option has zero stock")
        if not fibers:
            missing.append("Verified composition adding to 100%")
        if not color:
            missing.append("Verified color")
        for index,value in row.items():
            field = schema.headers[index]
            for group in (3,6,7,8):
                allowed = schema.dropdowns.get(f"t_{group}_{cid}_{field}") or schema.dropdowns.get(f"t_{group}_{field}")
                if allowed and str(value) not in allowed:
                    missing.append(f"Invalid dropdown value for {field}: {value}")
                    break
        all_rows.append({"cells":row,"product_id":pid,"sku":sku,"category":cid,"price_o":price,"source_excel_row":p["excel_row"],"options":combo,"missing":sorted(set(missing))})
    return all_rows


def col_letter(zero_index):
    out, n = "", zero_index + 1
    while n:
        n, rem = divmod(n-1, 26)
        out = chr(65+rem) + out
    return out


def write_template(template_path, output_path, records):
    """External-run writer: preserve all other ZIP entries byte for byte.

    Only explicit data rows are replaced. This avoids openpyxl reserializing
    Temu's unsupported extensions, dropdowns and hidden control worksheets.
    """
    with zipfile.ZipFile(template_path) as src:
        workbook = ET.fromstring(src.read("xl/workbook.xml"))
        rid_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        rid = next(s.get(rid_ns) for s in workbook.find(QN("sheets")) if s.get("name")=="Template")
        rels = ET.fromstring(src.read("xl/_rels/workbook.xml.rels"))
        target = next(x.get("Target") for x in rels if x.get("Id")==rid)
        sheet_path = target.lstrip("/") if target.startswith("/") else "xl/" + target
        original = src.read(sheet_path).decode("utf-8")
        m = re.search(r"<sheetData\b[^>]*>(.*?)</sheetData>", original, re.S)
        if not m:
            raise ValueError("Template sheetData not found")
        preserved = []
        style_row = None
        for row_text in re.findall(r"<row\b[^>]*(?:/>|>.*?</row>)", m[1], re.S):
            number = re.search(r'\br="(\d+)"', row_text)
            if number and int(number[1]) <= 4:
                preserved.append(row_text)
            elif number and int(number[1]) == 5:
                style_row = row_text
        if len(preserved) != 4:
            raise ValueError("Temu reserved rows 1-4 missing")
        styles = {}
        if style_row:
            for cell in re.findall(r"<c\b[^>]*(?:/>|>.*?</c>)", style_row, re.S):
                addr, style = re.search(r'\br="([A-Z]+)\d+"',cell), re.search(r'\bs="(\d+)"',cell)
                if addr and style:
                    styles[addr[1]] = style[1]
        from xml.sax.saxutils import escape
        authored = []
        for rownum, record in enumerate(records, 5):
            cells = []
            for index, value in sorted(record["cells"].items(), key=lambda pair:int(pair[0])):
                index = int(index)
                if value is None or value == "":
                    continue
                letter = col_letter(index)
                style = f' s="{styles[letter]}"' if letter in styles else ""
                addr = f'{letter}{rownum}'
                if isinstance(value,(int,float)) and not isinstance(value,bool):
                    if not math.isfinite(value):
                        raise ValueError("Non-finite cell value")
                    cells.append(f'<c r="{addr}"{style}><v>{value}</v></c>')
                else:
                    value = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", str(value))
                    cells.append(f'<c r="{addr}"{style} t="inlineStr"><is><t xml:space="preserve">{escape(value)}</t></is></c>')
            authored.append(f'<row r="{rownum}">' + "".join(cells) + '</row>')
        replacement = "<sheetData>" + "".join(preserved) + "".join(authored) + "</sheetData>"
        result = original[:m.start()] + replacement + original[m.end():]
        result = re.sub(r'<dimension\b[^>]*/>', f'<dimension ref="A1:AWB{max(5,len(records)+4)}"/>', result, count=1)
        # 1,276 columns are AWB. Assert the template width instead of changing it.
        with zipfile.ZipFile(output_path,"w",zipfile.ZIP_DEFLATED) as dst:
            for info in src.infolist():
                dst.writestr(info,result.encode("utf-8") if info.filename==sheet_path else src.read(info.filename))


def csv_report(path, entries):
    fields = list(dict.fromkeys(k for entry in entries for k in entry)) or ["status"]
    with path.open("w",encoding="utf-8-sig",newline="") as f:
        writer = csv.DictWriter(f,fieldnames=fields)
        writer.writeheader()
        for item in entries:
            writer.writerow({k:json.dumps(v,ensure_ascii=False) if isinstance(v,(dict,list)) else v for k,v in item.items()})


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input",default="input")
    ap.add_argument("--output",default="output")
    ap.add_argument("--config",default="config.json")
    ap.add_argument("--workers",type=int,default=3)
    ap.add_argument("--limit",type=int,default=0)
    ap.add_argument("--ids",help="Comma-separated product IDs for a controlled check")
    ap.add_argument("--refresh",action="store_true",help="Refresh compact cached product pages")
    ap.add_argument("--offline",action="store_true")
    ap.add_argument("--dry-run",action="store_true",help="Validate/extract JSON and reports without authoring XLSX")
    args = ap.parse_args()
    logging.basicConfig(level=logging.INFO,format="%(asctime)s %(message)s")
    out = Path(args.output); out.mkdir(parents=True,exist_ok=True)
    config = json.loads(Path(args.config).read_text(encoding="utf-8"))
    if not 1 <= config["rows_per_file"] <= 4996:
        ap.error("rows_per_file must be 1..4996 to remain inside the template validation ranges")
    if config["max_variants_per_product"] < 1:
        ap.error("max_variants_per_product must be positive")
    if not args.dry_run:
        for pattern in ("TEMU_UPLOAD_*.xlsx", "TEMU_DRAFT_NOT_FOR_UPLOAD_*.xlsx"):
            for old in out.glob(pattern):
                old.unlink()
    inputs = list(Path(args.input).glob("*.xlsx"))
    template = next((p for p in inputs if p.name.upper().startswith("TEMU")),None)
    export = next((p for p in inputs if p.name.lower().startswith("export_24tshirt")),None)
    if not template or not export:
        ap.error("Put the TEMU template and export_24tshirt.xlsx in input/")
    schema = Template(template)
    products = read_export(export)
    selected = list(products.values())
    if args.ids:
        ids = set(args.ids.split(",")); selected = [p for p in selected if p["id"] in ids]
        if ids - products.keys():
            ap.error("Unknown product IDs: " + ",".join(ids-products.keys()))
    if args.limit:
        selected = selected[:args.limit]
    fetcher = Fetcher(Path("cache"),args.refresh,args.offline)
    rows, review, excluded, jobs = [], [], [], []
    for p in selected:
        enabled = str(p.get("status")).lower() in {"true","1","yes"}
        if not enabled or float(p.get("quantity") or 0)<=0:
            excluded.append({"product_id":p["id"],"title":p["name(bg)"],"reason":"Inactive or product quantity zero"})
        elif SET_RE.search(p["name(bg)"]):
            review.append({"product_id":p["id"],"title":p["name(bg)"],"price_o":p["price_o"],"issue":"Bundle has a set price; needs suitable category or individual prices"})
        else:
            jobs.append(p)
    done = 0
    def process(p):
        return build_product(p,fetcher.get(p),schema,config)
    with futures.ThreadPoolExecutor(max_workers=max(1,min(args.workers,6))) as pool:
        pending = {pool.submit(process,p):p for p in jobs}
        for future in futures.as_completed(pending):
            p = pending[future]
            try:
                records = future.result()
                if not records:
                    raise ValueError("No product variants produced")
                rows.extend(records)
                for r in records:
                    if r["missing"]:
                        review.append({"product_id":p["id"],"title":p["name(bg)"],"sku":r["sku"],"price_o":r["price_o"],"issue":r["missing"]})
            except Exception as e:
                review.append({"product_id":p["id"],"title":p["name(bg)"],"price_o":p["price_o"],"issue":str(e)})
                LOG.warning("%s: %s",p["id"],e)
            done += 1
            if done % 10 == 0 or done == len(jobs):
                LOG.info("Verified %s/%s products, %s rows",done,len(jobs),len(rows))
    rows.sort(key=lambda r:(int(r["product_id"]),r["sku"]))
    skus = [r["sku"] for r in rows]
    if len(set(skus)) != len(skus):
        raise ValueError("Duplicate SKUs; output stopped")
    for r in rows:
        if r["cells"][schema.columns["Base Price - EUR"][0]] != products[r["product_id"]]["price_o"]:
            raise ValueError("Price does not equal Products!O")
        if BAD.search(r["cells"].get(schema.columns["Product Description"][0],"")):
            raise ValueError("Prohibited logistics/commercial text in description")
    ready = [r for r in rows if not r["missing"]]
    draft = [r for r in rows if r["missing"]]
    if not args.dry_run:
        for prefix, collection in [("TEMU_UPLOAD",ready),("TEMU_DRAFT_NOT_FOR_UPLOAD",draft)]:
            for offset in range(0,len(collection),config["rows_per_file"]):
                part = offset//config["rows_per_file"]+1
                write_template(template,out/f"{prefix}_{part:03d}.xlsx",collection[offset:offset+config["rows_per_file"]])
    csv_report(out/"review_required.csv",review)
    csv_report(out/"excluded.csv",excluded)
    csv_report(out/"price_audit.csv",[{k:r[k] for k in ["product_id","sku","price_o","source_excel_row","category","options"]} for r in rows])
    (out/"records.json").write_text(json.dumps(rows,ensure_ascii=False),encoding="utf-8")
    summary = {"export_products":len(products),"selected_products":len(selected),"verified_products_with_rows":len({r["product_id"] for r in rows}),"variation_rows":len(rows),"upload_ready_rows":len(ready),"draft_rows":len(draft),"review_products":len({r["product_id"] for r in review}),"excluded_products":len(excluded),"dry_run":args.dry_run,"price_rule":"exact Products!O; no VAT change, no Specials, no option surcharges","categories":dict(Counter(r["category"] for r in rows))}
    accounted = {r["product_id"] for r in rows}|{r["product_id"] for r in review}|{r["product_id"] for r in excluded}
    if accounted != {p["id"] for p in selected}:
        raise ValueError("A selected product is missing from output/review/excluded")
    (out/"summary.json").write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding="utf-8")
    LOG.info("RESULT %s",json.dumps(summary,ensure_ascii=False))


if __name__ == "__main__":
    main()
