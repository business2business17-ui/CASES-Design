"""Сборка XLSX: ТЗ + промты для ИИ-дизайнера (A+ Basic, A+ Premium, Brand Story, Brand Store).
Запуск: python3 -I tools/build_design_tz.py
"""
import collections
import re
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
import load_data as L  # noqa: E402

OUT = L.ROOT / "MOBILIUS_design_TZ.xlsx"

# ----------------------------------------------------------------- размеры (из amazon-specs.md агента)
SIZE = {
    "Basic": "970 x 300 px",
    "Premium": "1464 x 600 px",
    "Comparison": "150 x 300 px (на каждый товар в колонке)",
    "StoryBG": "1464 x 625 px",
    "StoryCard": "362 x 453 px",
}
WH = {"Basic": (970, 300), "Premium": (1464, 600)}

SERIES = {
    "215": dict(label="Clear (прозрачный)", finish="Clear", en="Clear MagSafe-compatible case",
                features=["Clear back shows the artwork", "Built-in magnetic ring - MagSafe compatible",
                          "Raised camera surround", "Ribbed side grips", "Easy-press buttons"]),
    "223": dict(label="Matte (матовый, 4 цвета)", finish="Matte", en="Matte MagSafe-compatible case",
                features=["Matte finish", "TPU frame and hard PC back", "Built-in magnetic ring - MagSafe compatible",
                          "Easy-press buttons"]),
    "406": dict(label="AirPods Pro (чёрный)", finish="TBC", en="Case for AirPods Pro",
                features=["Black case", "Carabiner loop"]),
}

# ----------------------------------------------------------------- темы принтов (для сцены/фона)
THEMES = [
    ("sport", "Sports", ["hockey", "basketball", "soccer", "football", "rugby"],
     "Dynamic studio look: bold diagonal color blocks in the print's main colors, light motion streaks, sporty but clean"),
    ("usa", "USA & Americana", ["usa", "liberty", "eagle", "american", "flag", "patriotic", "stars stripes", "stripes stars", "stars and stripes", "land of the free"],
     "Clean Americana: soft cream backdrop with subtle red and blue accents, light paper texture; no flags or logos as props"),
    ("travel", "Travel & national parks", ["california", "yosemite", "yellowstone", "zion", "acadia", "canyon", "joshua", "desert", "highway", "road trip", "coast", "surf", "stamp", "national park", "scooter", "speedboat", "seaside", "summer", "mediterranean", "racing", "convertible", "picnic", "forest sunset", "mountain", "pool", "sunset"],
     "Vintage travel-poster mood: warm sunlit gradient, subtle paper grain, soft landscape silhouettes; small props such as sunglasses, a folded map, a small plant"),
    ("gothic", "Gothic & dark", ["skull", "gothic", "skeleton", "vampire", "witch", "snake", "crowned", "shadow", "bat", "monster", "red moon", "chrome tribal", "enchanted", "cowgirl skull"],
     "Dark moody studio: deep charcoal backdrop with soft rim light and a hint of dark red; minimal props (dried roses, velvet fabric)"),
    ("japan", "Japanese-inspired", ["dragon", "geisha", "samurai", "sakura", "cherry blossom", "katana", "tokyo", "lotus", "japan", "koi", "torii", "kitsune", "lucky cat", "matcha", "temple", "wave", "crane", "ink$", "calligraphy", "manga", "anime", "dumpling"],
     "Japanese-inspired mood: warm off-white washi-paper texture, soft ink-wash shapes, a single red sun-disc accent; calm minimal props (folded fan, bamboo, blossom petals)"),
    ("fruit", "Fruit & fresh", ["fruit", "citrus", "pear$", "orange pop$", "apple", "lemon", "pomegranate", "dragon fruit", "juicy", "popsicle"],
     "Fresh market mood: bright pastel backdrop, real-looking citrus and fruit slices, water droplets, sunny light"),
    ("floral", "Floral & botanical", ["flower", "blossom", "botanical", "daisy", "lavender", "poppy", "rose", "bouquet", "sunflower", "butterfl", "bow$", "berry", "berries", "strawberr", "cherries", "blueberr", "peony", "currant", "blackberr", "pressed"],
     "Soft botanical mood: pastel backdrop with real petals and leaves, window-light shadows, airy and fresh"),
    ("animal", "Animal prints", ["leopard", "zebra", "tiger", "cow$", "animal"],
     "Bold pattern-on-pattern studio: single saturated color-block backdrop that complements the print's accent color, crisp shadows, fashion-editorial"),
    ("cute", "Playful & cute", ["cat", "dog", "bear", "lamb", "frog", "goose", "corgi", "cute", "dachshund", "teddy", "pup", "hello", "vibes", "learn", "study", "class", "napping", "pop$", "sticker", "collage"],
     "Playful bright mood: pastel color-block backdrop with generic confetti and sticker-like shapes (no characters), soft shadows"),
    ("pattern", "Patterns & abstract", ["dot", "grid", "ring", "tile", "swirl", "stripe", "circle", "checker", "pattern", "abstract", "spots", "starburst", "web", "ripple", "heart", "sun$"],
     "Minimal color-block studio: seamless backdrop in a tint of the print's main color, simple geometric plinths, soft clean shadows"),
]
DEFAULT_THEME = ("graphic", "Graphic & art",
                 "Neutral modern studio: soft gray-to-white gradient with one accent color taken from the print, clean soft shadow")


def _hit(kws, text):
    for k in kws:   # совпадение с начала слова; 'xxx$' — слово целиком
        pat = r"\b" + re.escape(k[:-1]) + r"\b" if k.endswith("$") else r"\b" + re.escape(k)
        if re.search(pat, text):
            return True
    return False


def theme_of(name, desc):
    t = (name + " " + desc).lower()
    nm = name.lower()
    if "dragon fruit" in nm:
        return next((k, l, st) for k, l, _, st in THEMES if k == "fruit")
    for key, label, kws, style in THEMES:      # сначала по названию, потом по описанию
        if _hit(kws, nm):
            return key, label, style
    for key, label, kws, style in THEMES:
        if _hit(kws, t):
            return key, label, style
    return DEFAULT_THEME


NEG = ("no text, no letters, no numbers, no logos (no Apple logo, no Amazon logo), no watermark, no price tags, "
       "no badges such as 'best seller', no hands, no extra phone cases, no redrawn or altered case, no warped print, "
       "no extra cameras or buttons")

RULES_FORBIDDEN = ("shockproof / drop-proof / military-grade / impact or fall protection; 'protects'; "
                   "reinforced corners; Amazon's Choice / Best Seller / #1; prices, discounts; reviews and stars; "
                   "competitor names; QR codes and links; Apple and Amazon logos; unconfirmed certificates")


def families(models):
    nums = sorted({int(m) for mm in models for m in re.findall(r"iPhone (\d+)", mm)})
    if not nums:
        return ""
    return f"iPhone {nums[0]}-{nums[-1]}" if len(nums) > 1 else f"iPhone {nums[0]}"


def ru_list(items):
    return ", ".join(items)


# ----------------------------------------------------------------- построение наборов
def build_sets(products):
    sets = collections.OrderedDict()
    for p in products:
        sets.setdefault((p["series"], p["q"]), []).append(p)
    out = []
    for (series, q), ps in sets.items():
        name = ps[0]["name"]
        desc = ps[0]["desc"]
        models = sorted({p["model"] for p in ps}, key=L.model_sort_key)
        colors = [c for c in L.COLOR_ORDER if any(p["color"] == c for p in ps)] if series == "223" else \
                 ([ps[0]["color"]])
        hero_model = next((m for m in L.HERO_MODELS if m in models), models[0])
        hero_pool = [p for p in ps if p["model"] == hero_model]
        if series == "223":
            hero = next(p for c in L.COLOR_ORDER for p in hero_pool if p["color"] == c)
        else:
            hero = hero_pool[0]
        tk, tl, ts = theme_of(name, desc)
        out.append(dict(series=series, q=q, id=f"AP-{series}-q{q:03d}", name=name, desc=desc, models=models,
                        colors=colors, hero=hero, hero_model=hero_model, theme=tk, theme_label=tl,
                        theme_style=ts, skus=ps))
    return out


def hero_url(p, idx=0):
    if p["photos"]:
        return p["photos"][idx]
    return f"assets/406_prints/q{p['q']}.jpg"


def neighbors(s, all_sets, n=4):
    """Соседние принты той же серии: сначала той же темы, затем остальные (по q, по кругу)."""
    same = [x for x in all_sets if x["series"] == s["series"] and x["q"] != s["q"]]
    same.sort(key=lambda x: ((x["q"] - s["q"]) % 10000))
    pick = [x for x in same if x["theme"] == s["theme"]][:n]
    for x in same:
        if len(pick) >= n:
            break
        if x not in pick:
            pick.append(x)
    return pick


def compare_columns(s, all_sets):
    """[(label, sku, photo_url, design, color)] — только внутри серии."""
    cols = []
    if s["series"] == "223":
        for c in s["colors"]:
            cand = [p for p in s["skus"] if p["color"] == c]
            p = next((p for p in cand if p["model"] == s["hero_model"]),
                     min(cand, key=lambda p: L.model_sort_key(p["model"]) if p["model"] not in L.HERO_MODELS else -1))
            cols.append((f"{s['name']} - {c}", p["sku"], hero_url(p), s["name"], c))
        if len(cols) < 3:   # меньше 3 колонок — добираем соседей (цвет как у героя)
            for x in neighbors(s, all_sets, 4 - len(cols)):
                cols.append((f"{x['name']} - {x['hero']['color']}", x["hero"]["sku"], hero_url(x["hero"]),
                             x["name"], x["hero"]["color"]))
    else:
        cols.append((f"{s['name']} (this design)", s["hero"]["sku"], hero_url(s["hero"]), s["name"], s["hero"]["color"]))
        for x in neighbors(s, all_sets, 4):
            cols.append((x["name"], x["hero"]["sku"], hero_url(x["hero"]), x["name"], x["hero"]["color"]))
    return cols


def compare_rows(s, cols):
    if s["series"] == "406":
        return ["Design", "Color", "Fits", "Carabiner loop"]
    return ["Design", "Finish", "Color", "Fits", "MagSafe compatible"]


def compare_cells(s, cols):
    fam = families(s["models"]) or "AirPods Pro"
    rows = {}
    rows["Design"] = [c[3] for c in cols]
    rows["Finish"] = [SERIES[s["series"]]["finish"] for _ in cols]
    rows["Color"] = [c[4] for c in cols]
    rows["Fits"] = [("AirPods Pro" if s["series"] == "406" else fam) for _ in cols]
    rows["MagSafe compatible"] = ["Yes" for _ in cols]
    rows["Carabiner loop"] = ["Yes" for _ in cols]
    return rows


# ----------------------------------------------------------------- тексты слайдов
def slide_specs(s, tier, all_sets):
    ser = SERIES[s["series"]]
    fam = families(s["models"]) or "AirPods Pro"
    W, H = WH[tier]
    url = hero_url(s["hero"])
    is_406 = s["series"] == "406"
    feats = ser["features"]
    model_line = ("AirPods Pro" if is_406 else " | ".join(s["models"]))
    colors = ", ".join(s["colors"])
    common_in = (f"INPUT: use the attached real product photo ({url}) as an untouched layer: cut out the case and "
                 f"place/scale/rotate it only. Do NOT redraw, restyle, recolor or re-render the case, its print, "
                 f"camera cutout or magnetic ring. Generate only the background and props. Canvas {W}x{H} px, sRGB, "
                 f"flat A+ layout. Leave clean empty areas for text (text is added later in Figma). ")
    style = f"Style: {s['theme_style']}. "
    sl = []
    if tier == "Basic":
        sl.append(dict(
            module="Header Image with Text / banner", goal="Обещание: показать принт и серию",
            ru=f"Баннер {W}x{H}. Реальное фото чехла справа (~40% ширины, лёгкий наклон), слева пустая зона под заголовок и подзаголовок. Фон по теме принта «{s['theme_label']}». Принт на чехле не трогать.",
            text=f"H: {s['name']}\nSub: {ser['en']} for {fam}" if not is_406 else f"H: {s['name']}\nSub: {ser['en']}",
            prompt=common_in + f"Hero banner: case on the right (~40% of width) slightly tilted, soft contact shadow; left 50% empty for headline. " + style))
        sl.append(dict(
            module="Standard Image & Text / features", goal="Показать, что внутри дизайна: описание принта + 3 факта",
            ru=f"Фото чехла слева, справа 3 пустых слота-выноски (иконка + 2-4 слова). Описание принта идёт строкой под заголовком. Факты только из списка: {ru_list(feats[:3])}.",
            text="H: The design\nBody: " + s["desc"] + "\nCallouts:\n- " + "\n- ".join(feats[:3]),
            prompt=common_in + "Feature banner: case on the left, three empty callout slots on the right connected by thin hairline connectors (no text), clean simple line-icon placeholders. " + style))
        sl.append(dict(
            module="Standard Image & Text / fit", goal="Совместимость: модели" + ("" if is_406 else " и MagSafe"),
            ru=f"Узкий баннер совместимости: небольшое фото чехла слева, справа чипы моделей (список ниже) и строка про магнитное кольцо." + (f" Цвета в наличии для этого принта: {colors}." if s["series"] == "223" else ""),
            text=("H: Find your fit\nChips: " + model_line + ("" if is_406 else "\nLine: Built-in magnetic ring - MagSafe compatible") + (f"\nColors: {colors}" if s["series"] == "223" else "")),
            prompt=common_in + "Fit banner: small case on the left, plain clean backdrop on the right with generous empty space for model chips and one line of text. " + style))
    else:
        sl.append(dict(
            module="Full Image / hero lifestyle", goal="Эмоция: принт в сцене",
            ru=f"Широкая сцена {W}x{H}: реальное фото чехла лежит/опирается на поверхности среди реквизита по теме «{s['theme_label']}», без рук и без людей. Слева пустая зона под короткий слоган.",
            text=f"H: {s['name']}",
            prompt=common_in + f"Lifestyle hero: the case resting on a surface among props, shallow depth of field, no hands, no people. Left 35% empty for a short headline. " + style))
        sl.append(dict(
            module="Image with Text / print close-up", goal="Принт крупно + описание",
            ru="Крупный кадр принта: кроп с оригинального фото 2000x2000 (магнитное кольцо и принт). ИИ только достраивает фон по краям. Справа/слева блок текста: описание принта (дословно).",
            text="H: Look closer\nBody: " + s["desc"],
            prompt=common_in + "Close-up detail layout: the real photo is cropped and enlarged on the print, AI only extends the background at the edges; keep print pixels untouched. Empty text column on one side. " + style))
        if s["series"] == "223":
            third = dict(module="Image & Text / color lineup", goal="Выбор цвета",
                         ru=f"Ряд реальных фото этого принта во всех цветах ({colors}) одного размера на едином фоне; под каждым цвет подписью. Фото берутся из SKU сравнительной таблицы.",
                         text="H: Pick your color\nLabels: " + colors,
                         prompt=common_in + f"Lineup: {len(s['colors'])} real case photos of the same design in different colors ({colors}) arranged in an even row on one seamless backdrop, identical scale, empty strip under each for a label. " + style)
        elif is_406:
            third = dict(module="Image & Text / carry", goal="Петля под карабин",
                         ru="Реальное фото чехла с петлёй под карабин на спокойной сцене (рюкзак/ключи как реквизит из ИИ-фона). Проверить, что петля есть на фото источника.",
                         text="H: Carry it your way\nBody: Black case with carabiner loop",
                         prompt=common_in + "Carry scene: backdrop with generic backpack strap and keys as props, case photo placed in the foreground, no hands. " + style)
        else:
            third = dict(module="Image & Text / MagSafe", goal="MagSafe-совместимость",
                         ru="Реальное фото чехла + безымянный круглый беспроводной зарядник/магнитный аксессуар как ИИ-реквизит рядом (без логотипов). Текст: одна строка про магнитное кольцо.",
                         text="H: Made for MagSafe\nBody: Built-in magnetic ring - MagSafe compatible\nFits: " + fam,
                         prompt=common_in + "MagSafe scene: generic unbranded round magnetic wireless charger puck lying next to the case on a clean surface, no logos. " + style)
        sl.append(third)
    return sl


def comparison_row(s, tier, all_sets):
    cols = compare_columns(s, all_sets)
    rows = compare_rows(s, cols)
    cells = compare_cells(s, cols)
    col_lines = [f"{i+1}. {c[0]} | SKU {c[1]} | источник: {c[2]}" for i, c in enumerate(cols)]
    tbl = "\n".join(f"{r}: " + " | ".join(cells[r]) for r in rows)
    ru = (f"Таблица сравнения ({'Standard' if tier == 'Basic' else 'Enhanced'} comparison chart), колонок: {len(cols)}. "
          f"Мини-картинка товара в каждой колонке 150 x 300 px — это КРОП реального фото (не генерация): вырезать чехол на белом/прозрачном фоне, вертикально, "
          f"без теней-артефактов, читаемо на мобильном. Колонка 1 — текущий товар. ASIN подставить по SKU в Seller Central.")
    return dict(module="Comparison chart", goal="Сравнение с мини-картинками", ru=ru,
                text="Columns:\n" + "\n".join(col_lines) + "\nRows:\n" + tbl,
                prompt="NO GENERATION. Crop each product from the source photo, straight vertical orientation, 150x300 px, neutral white background, keep print pixels untouched.",
                url="; ".join(c[2] for c in cols))


# ----------------------------------------------------------------- Brand Story / Store
def brand_assets(sets, products):
    by_theme = collections.OrderedDict()
    for s in sets:
        by_theme.setdefault((s["theme"], s["theme_label"], s["theme_style"]), []).append(s)
    return by_theme


def sample_urls(sets, series=None, n=6):
    pool = [s for s in sets if series is None or s["series"] == series]
    step = max(1, len(pool) // n)
    return [hero_url(s["hero"]) for s in pool[::step][:n]]


def story_rows(sets, themes):
    n215 = sum(1 for s in sets if s["series"] == "215")
    n223 = sum(1 for s in sets if s["series"] == "223")
    n406 = sum(1 for s in sets if s["series"] == "406")
    common = ("INPUT: real product photos only (listed). Do NOT redraw or alter any case or print; composite only. "
              "Generate background/props only. No text, logos, prices, Apple/Amazon logos, no shockproof/drop claims. ")
    rows = [
        ("BG", "Background (весь блок)", SIZE["StoryBG"], "Единый фон бренда: коллаж реальных чехлов, мягкая ткань/бумага, тёмный нейтральный тон. Критичное — по центру: карточки и мобайл-кроп закроют края.",
         "Brand: MOBILIUS (logo to be supplied)", sample_urls(sets, None, 8),
         common + "Wide 1464x625 brand backdrop: 6-8 real case photos fanned out softly blurred on a dark neutral fabric/paper texture, calm color accents from the prints, large clean center."),
        ("Card 1", "Карточка бренда", SIZE["StoryCard"], "Вертикальная карточка: логотип MOBILIUS + 1 строка. Нужны логотип и разрешённый слоган.",
         "H: MOBILIUS\nBody: [TAGLINE_NEEDED]", sample_urls(sets, "223", 1),
         common + "Vertical 362x453 card: one real case photo, brand-neutral dark backdrop, clean space on top for the logo and at the bottom for one line."),
        ("Card 2", "Дизайны на любой вкус", SIZE["StoryCard"], f"Мозаика из реальных принтов по темам ({', '.join(t[1] for t in themes)}).",
         f"H: Designs for every mood\nBody: {n215 + n223 + n406} design sets across 3 series", sample_urls(sets, None, 6),
         common + "Vertical 362x453 card: grid of 6 real case photos of different themes on one backdrop, equal size."),
        ("Card 3", "Серия Clear (215)", SIZE["StoryCard"], "Прозрачный чехол: принт виден на прозрачной задней панели.",
         "H: Clear series\nBody: Clear back shows the artwork", sample_urls(sets, "215", 1),
         common + "Vertical 362x453 card: one real clear case photo, airy light backdrop with soft reflections."),
        ("Card 4", "Серия Matte (223), 4 цвета", SIZE["StoryCard"], "Матовый чехол в цветах Beige / Black / Gray / Orange.",
         "H: Matte series\nBody: Beige, Black, Gray, Orange", sample_urls(sets, "223", 1),
         common + "Vertical 362x453 card: four real photos of the same design in the four colors, stacked, one seamless backdrop."),
        ("Card 5", "MagSafe", SIZE["StoryCard"], "Встроенное магнитное кольцо, совместимо с MagSafe (без обещаний защиты).",
         "H: Made for MagSafe\nBody: Built-in magnetic ring - MagSafe compatible", sample_urls(sets, "215", 1),
         common + "Vertical 362x453 card: one real case photo with a generic unbranded magnetic charger puck beside it, clean surface."),
        ("Card 6", "Серия AirPods Pro (406)", SIZE["StoryCard"], "Чёрный чехол для AirPods Pro с петлёй под карабин — проверить факты по фото.",
         "H: For AirPods Pro\nBody: Black case with carabiner loop", sample_urls(sets, "406", 1),
         common + "Vertical 362x453 card: one real AirPods case photo on a dark neutral backdrop with soft rim light."),
    ]
    return rows


def store_rows(sets, themes, models):
    common = ("INPUT: only real product photos (listed) composited as untouched layers. No generated cases. "
              "No embedded text (Amazon discourages text inside Store images; texts are native in the Store builder). "
              "No logos of Apple/Amazon, no prices, no shockproof/drop claims. ")
    rows = []
    add = lambda *a: rows.append(a)
    # Home
    add("Home", "H1", "Header hero", "3000 x 600 px (safe zone: центр, до 15% обрезки с каждой стороны), <=5 MB",
        "Hero: веер реальных чехлов по трём сериям на тёмном/нейтральном фоне; ключевое — в центральных 70%.", "MOBILIUS - phone & AirPods cases", sample_urls(sets, None, 8),
        common + "Wide 3000x600 hero: fan of 8 real cases, soft blur and shadow, keep all vital content in the central 70%, large clean areas left/right.")
    add("Home", "H2", "Brand logo", "400 x 400 px, <=5 MB", "Логотип MOBILIUS (файл от заказчика, не генерировать).", "-", [], "NO GENERATION - supplied logo file.")
    for t, (s_id, lab, url, txt) in zip(["T1", "T2", "T3"], [
            ("215", "Clear series", sample_urls(sets, "215", 4), "Clear series - MagSafe-compatible"),
            ("223", "Matte series", sample_urls(sets, "223", 4), "Matte series - 4 colors"),
            ("406", "AirPods Pro series", sample_urls(sets, "406", 4), "AirPods Pro cases")]):
        add("Home", t, "Image tile - Large", "1500 x 1500 px, <=5 MB (низ ~12% может быть закрыт заголовком)",
            f"Плитка серии {lab}: 3-4 реальных чехла в композиции. Ссылка: страница серии {s_id}.", txt, url,
            common + f"Square 1500x1500 tile: 3-4 real case photos of the {lab} in a neat overlapping arrangement on a seamless backdrop, bottom 12% kept empty.")
    add("Home", "T4", "Image tile - Medium", "1500 x 750 px, <=5 MB (низ ~19% может быть закрыт)", "Плитка «Shop by iPhone model». Ссылка: страница моделей.",
        "Shop by iPhone model", sample_urls(sets, "223", 5),
        common + "1500x750 tile: row of 5 real case photos (different models) at equal scale, bottom 19% kept empty.")
    add("Home", "T5", "Image tile - Medium", "1500 x 750 px, <=5 MB", "Плитка «Shop by design». Ссылка: каталог принтов.", "Shop by design", sample_urls(sets, None, 6),
        common + "1500x750 tile: collage of 6 real case photos of different design themes, bottom 19% kept empty.")
    add("Home", "G1", "Gallery (до 8 изображений)", "1500 x 750 px минимум, до 8 шт., <=5 MB", "8 лайфстайл-кадров по темам принтов из таблицы «Темы» (по одному на тему).", "-", sample_urls(sets, None, 8),
        common + "Per image: 1500x750 lifestyle scene with one real case on a themed backdrop (use the style line of the theme from sheet 'Темы').")
    # Series pages
    for pid, series, lab in [("P-215", "215", "Clear series"), ("P-223", "223", "Matte series"), ("P-406", "406", "AirPods Pro series")]:
        add(pid, "H1", "Header hero", "3000 x 600 px, <=5 MB", f"Hero серии {lab}.", lab, sample_urls(sets, series, 6),
            common + f"Wide 3000x600 hero for the {lab}: 5-6 real cases fanned out, keep content in the central 70%.")
        if series == "223":
            for c in ["Beige", "Black", "Gray", "Orange"]:
                add(pid, f"C-{c}", "Image tile - Small", "750 x 750 px", f"Плитка цвета {c}: 2-3 реальных чехла этого цвета. Ссылка: фильтр по цвету.", f"Color {c}",
                    [hero_url(s["hero"]) for s in sets if s["series"] == "223" and c in s["colors"]][:3],
                    common + f"Square 750x750 tile: 2-3 real case photos in {c}, seamless backdrop.")
        for key, label, style in [(k, lab2, st) for k, lab2, st in [(t[0], t[1], t[3]) for t in themes]]:
            urls = [hero_url(s["hero"]) for s in sets if s["series"] == series and s["theme"] == key][:3]
            if not urls:
                continue
            add(pid, f"TH-{key}", "Image tile - Small", "750 x 750 px", f"Плитка темы «{label}» серии {series}. Ссылка: фильтр по теме.", f"Theme {label}", urls,
                common + f"Square 750x750 tile: 2-3 real case photos of the '{label}' theme. Mood: {style}.")
    # Models page
    add("P-MODELS", "H1", "Header hero", "3000 x 600 px", "Hero «Shop by iPhone model».", "Shop by iPhone model", sample_urls(sets, "223", 6),
        common + "Wide 3000x600 hero: row of real cases for different iPhone models at equal scale, central 70% safe.")
    for m in models:
        urls = [p["photos"][0] for s in sets if s["series"] in ("215", "223") for p in s["skus"] if p["model"] == m and p["photos"]][:2]
        add("P-MODELS", "M-" + re.sub(r"\W+", "-", m).strip("-"), "Image tile - Small", "750 x 750 px", f"Плитка модели {m}: 1-2 реальных чехла для этой модели. Ссылка: фильтр по модели.", m, urls,
            common + f"Square 750x750 tile: 1-2 real case photos made for {m}, seamless backdrop.")
    return rows


# ----------------------------------------------------------------- запись XLSX
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
WRAP = Alignment(wrap_text=True, vertical="top")


def write_sheet(wb, title, header, rows, widths, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = title
    ws.append(header)
    for r in rows:
        ws.append(list(r))
    for c in ws[1]:
        c.fill, c.font = HDR_FILL, HDR_FONT
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = WRAP
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    return ws


def main():
    products = L.load()
    sets = build_sets(products)
    themes = [(k, lab, kws, st) for k, lab, kws, st in THEMES] + [(DEFAULT_THEME[0], DEFAULT_THEME[1], [], DEFAULT_THEME[2])]
    set_by_key = {(s["series"], s["q"]): s for s in sets}
    models = sorted({p["model"] for p in products if p["series"] in ("215", "223")}, key=L.model_sort_key)
    models_ui = [m for m in models]

    wb = Workbook()

    # ---- Лист 1: Читать
    cnt = collections.Counter(s["series"] for s in sets)
    theme_cnt = collections.Counter(s["theme_label"] for s in sets)
    info = [
        ("MOBILIUS — ТЗ и промты для ИИ-дизайнера (A+ Basic, A+ Premium, Brand Story, Brand Store)",),
        ("",),
        ("СТРУКТУРА",),
        ("Лист «SKU»: все SKU (новое, старое, серия, модель, принт, цвет, ID набора A+).",),
        ("Лист «Принты (наборы A+)»: 1 строка = 1 набор A+ (принт внутри серии). Все модели и цвета этого принта делят один набор.",),
        ("Лист «A+ ТЗ+промты»: по каждому набору — Basic и Premium по 3 слайда + таблица сравнения (мини-картинки).",),
        ("Лист «Brand Story» и «Brand Store»: один комплект на все SKU.",),
        ("Лист «Темы»: стили фона для промтов и счётчики.",),
        ("",),
        ("ЦИФРЫ",),
        (f"SKU всего: {len(products)} (215: {sum(1 for p in products if p['series']=='215')}, 223: {sum(1 for p in products if p['series']=='223')}, 406: {sum(1 for p in products if p['series']=='406')})",),
        (f"Наборов A+ (принтов в серии): {len(sets)} (215: {cnt['215']}, 223: {cnt['223']}, 406: {cnt['406']}). Строк ТЗ: {len(sets)*8}.",),
        ("",),
        ("РЕШЕНИЯ ЗАКАЗЧИКА (учтены)",),
        ("1. Brand Store и Brand Story — один на все SKU.",),
        ("2. A+ — один на принт; модели и цвета внутри принта в один набор. 215 и 223 — разные наборы (прозрачный / матовый).",),
        ("3. Серия 406: одно актуальное SKU, как в каталоге PDF (406xxx-002Ru001qNNN). Серии 215/223 в таблицах — старое и новое SKU.",),
        ("4. Размеры из agent MD (amazon-specs.md): Basic 970x300, Premium 1464x600, сравнение 150x300 на товар, Brand Story фон 1464x625 + карточка 362x453, Store hero 3000x600; A+ и Story <=2 MB.",),
        ("5. A+: 3 слайда + таблица сравнения с мини-картинками (для Basic и для Premium).",),
        ("6. MagSafe писать можно; про ударопрочность и защиту не писать.",),
        ("",),
        ("ДОПУЩЕНИЯ (проверить)",),
        ("а) Рынок — США, язык текста на картинках — английский (американское написание, как в таблицах).",),
        ("б) Матовость серии 223 — со слов заказчика (в таблице написано TPU+PC без слова matte).",),
        ("в) Размеры слайдов Basic/Premium — общие значения из MD; точные боксы модуля — VERIFY_IN_UI.",),
        ("г) Источник товара для дизайнера — фото _2 выбранной модели (hero-модель: iPhone 17, иначе первая по списку). Фото _3.._9 — готовые инфографики, не использовать как фон.",),
        ("д) Таблица сравнения строится ТОЛЬКО внутри серии: 223 — этот принт во всех цветах; 215/406 — этот принт + 4 принта той же темы. Кросс-серийные ссылки не делаю: q-номера у 215 и 223 совпадают, но названия различаются в 123 случаях, а в q055, q150 описания разные. Есть 3 принта только в 215 (q043, q052, q053).",),
        ("е) 406: названия и описания принтов составлены по картинкам каталога (низкое разрешение 402x269) — проверить; поколение AirPods Pro, материал и матовость неизвестны (TBC).",),
        ("ж) Факты в ТЗ: только из описаний таблиц + заметные на фото. «Reinforced corners» из описаний 215 не использую (близко к защите).",),
        ("",),
        ("ЗАПРЕЩЕНО НА КАРТИНКАХ",),
        (RULES_FORBIDDEN,),
        ("",),
        ("ЧТО НУЖНО ОТ ЗАКАЗЧИКА",),
        ("1) Логотип MOBILIUS и фирменные цвета/шрифты, слоган для Brand Story.",),
        ("2) Hi-res фото принтов 406 (в PDF 402x269) и нового SKU/ASIN, если появятся.",),
        ("3) ASIN для таблицы сравнения (в таблицах только SKU).",),
        ("4) Подтвердить рынок/язык и размеры боксов модулей в вашем билдере A+.",),
        ("",),
        ("ПРАВИЛА ИИ-ДИЗАЙНЕРА",),
        ("Реальное фото товара не перерисовывать: ИИ делает только фон/реквизит. Текст на слайдах добавляется в Figma (ИИ портит буквы), в промтах текста нет. Проверка каждого слайда: принт совпадает с эталоном, нет лишних кнопок/камер, нет запрещённых слов.",),
    ]
    ws = wb.active
    ws.title = "Читать"
    for r in info:
        ws.append(list(r))
    ws.column_dimensions["A"].width = 160
    for row in ws.iter_rows():
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    for r in (1, 3, 10, 14, 21, 31, 33, 39):
        ws.cell(r, 1).font = Font(bold=True, size=12 if r > 1 else 14)

    # ---- Лист 2: SKU
    sku_rows = []
    for p in products:
        s = set_by_key[(p["series"], p["q"])]
        photos = ", ".join(p["photos"]) if p["photos"] else f"assets/406_prints/q{p['q']}.jpg (из PDF, 402x269)"
        sku_rows.append((p["sku"], p["sku_old"] or "n/a (одно актуальное SKU)" if p["series"] == "406" else p["sku_old"],
                         p["series"], SERIES[p["series"]]["label"], p["model"], f"q{p['q']:03d}", p["name"], p["color"],
                         p["finish"], s["id"], p["title"] or "(нет — серия 406)", hero_url(p), photos))
    write_sheet(wb, "SKU", ["SKU новое (актуальное)", "SKU старое (каталог PDF)", "Серия", "Тип серии", "Модель", "Print ID",
                            "Принт", "Цвет", "Покрытие", "ID набора A+", "Amazon title", "Фото-источник (_2)", "Все фото"],
                sku_rows, [30, 30, 8, 22, 28, 9, 28, 9, 10, 16, 60, 55, 80])

    # ---- Лист 3: Принты
    pr_rows = []
    for s in sets:
        pr_rows.append((s["id"], s["series"], f"q{s['q']:03d}", s["name"], s["desc"], s["theme_label"],
                        ru_list(s["colors"]) if s["series"] == "223" else s["colors"][0],
                        ru_list(s["models"]), len(s["skus"]), s["hero_model"], s["hero"]["sku"], hero_url(s["hero"])))
    write_sheet(wb, "Принты (наборы A+)", ["ID набора A+", "Серия", "Print ID", "Принт", "Описание принта (источник)", "Тема",
                                           "Цвета", "Модели", "Кол-во SKU", "Hero-модель", "Hero SKU", "Hero-фото"],
                pr_rows, [16, 8, 9, 28, 70, 20, 22, 60, 9, 14, 30, 55])

    # ---- Лист 4: A+
    ap_rows = []
    for s in sets:
        for tier in ("Basic", "Premium"):
            specs = slide_specs(s, tier, sets)
            for i, sp in enumerate(specs, 1):
                ap_rows.append((s["id"], s["series"], s["name"], tier, i, sp["module"], SIZE[tier], sp["goal"], sp["ru"],
                                sp["text"], hero_url(s["hero"]), sp["prompt"], NEG,
                                "Принт 1:1 с фото-источника; нет запрещённых слов; <=2 MB; RGB", "TODO"))
            cr = comparison_row(s, tier, sets)
            ap_rows.append((s["id"], s["series"], s["name"], tier, 4, cr["module"], SIZE["Comparison"], cr["goal"], cr["ru"],
                            cr["text"], cr["url"], cr["prompt"], "-", "Только реальные кропы; ASIN подставить", "TODO"))
    write_sheet(wb, "A+ ТЗ+промты", ["ID набора A+", "Серия", "Принт", "Тип A+", "Слайд №", "Модуль Amazon", "Размер", "Цель",
                                     "ТЗ дизайнеру (RU)", "Текст на слайде (EN, добавить в Figma)", "Фото-источник (URL)",
                                     "Промт для ИИ (EN)", "Negative prompt", "Проверка", "Статус"],
                ap_rows, [16, 7, 24, 9, 8, 26, 16, 26, 60, 50, 50, 90, 45, 30, 9])

    # ---- Лист 5: Brand Story
    st_rows = []
    for no, title, size, ru, text, urls, prompt in story_rows(sets, themes):
        st_rows.append((no, title, size, ru, text, "; ".join(urls), prompt, NEG, "TODO"))
    write_sheet(wb, "Brand Story", ["Элемент", "Название", "Размер", "ТЗ дизайнеру (RU)", "Текст (EN, в Figma)", "Фото-источники",
                                    "Промт для ИИ (EN)", "Negative prompt", "Статус"], st_rows, [10, 28, 14, 60, 45, 70, 80, 45, 9])

    # ---- Лист 6: Brand Store
    bs_rows = []
    for page, tid, ttype, size, ru, txt, urls, prompt in store_rows(sets, themes, models_ui):
        bs_rows.append((page, tid, ttype, size, ru, txt + " (native text в билдере)", "; ".join(urls), prompt, NEG, "TODO"))
    write_sheet(wb, "Brand Store", ["Страница", "ID плитки", "Тип", "Размер", "ТЗ дизайнеру (RU)", "Подпись (native text)",
                                    "Фото-источники", "Промт для ИИ (EN)", "Negative prompt", "Статус"],
                bs_rows, [10, 12, 22, 34, 60, 34, 70, 80, 45, 9])

    # ---- Лист 7: Темы
    th_rows = [(k, lab, st, theme_cnt.get(lab, 0)) for k, lab, kws, st in themes]
    write_sheet(wb, "Темы", ["Код", "Тема", "Стиль фона/реквизита (EN, подставляется в промты)", "Наборов A+"], th_rows, [10, 26, 120, 12])

    wb.save(OUT)
    print("saved", OUT, "sets", len(sets), "A+ rows", len(ap_rows), "story", len(st_rows), "store", len(bs_rows))
    print(theme_cnt)


if __name__ == "__main__":
    main()
