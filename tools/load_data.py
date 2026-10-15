"""Загрузка и нормализация данных серий 215 / 223 / 406 (MOBILIUS)."""
import re
import openpyxl
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

MODEL_ORDER = ["iPhone 13", "iPhone 13 Pro Max", "iPhone 14", "iPhone 14 Pro Max", "iPhone 15",
               "iPhone 15 Pro", "iPhone 15 Pro Max", "iPhone 16", "iPhone 16 Plus", "iPhone 16 Pro",
               "iPhone 16 Pro Max", "iPhone 16e / iPhone 17e", "iPhone 17", "iPhone 17 Pro",
               "iPhone 17 Pro / iPhone 18 Pro", "iPhone 17 Pro Max", "iPhone 17 Pro Max / iPhone 18 Pro Max"]
HERO_MODELS = ["iPhone 17", "iPhone 16 Pro", "iPhone 15", "iPhone 13"]
COLOR_ORDER = ["Beige", "Gray", "Black", "Orange"]

# 406: описания принтов по картинкам каталога (визуальный разбор; проверить владельцу)
P406 = {
 212: ("Climbing Stamps", "Three vintage-style stamp tiles with climbing gear, a starry night desert and a Joshua Tree."),
 213: ("Yellowstone Stamps", "Three vintage-style stamp tiles with a geyser pool, a ranger hat with binoculars and a Yellowstone bison."),
 214: ("Peony Raspberries", "A pink peony with raspberries, buds and green leaves."),
 215: ("Dark Bloom", "Black and white gothic motifs - roses, a cameo, a bow, a moon and a dagger - on a Dark Bloom banner."),
 216: ("Gothic Bat", "An ornate bat with red roses, a crescent moon and a goblet in dark red and gold."),
 217: ("Cherry Blossom Mix", "Red cherries with pink and blue blossoms and green leaves."),
 218: ("Pear Parade", "A blue oval badge with two cheerful walking pears and Pear Parade lettering."),
 219: ("Yosemite Stamps", "Three vintage-style stamp tiles with a tent and lantern, a giant sequoia and Yosemite Half Dome."),
 220: ("Acadia Stamps", "Three vintage-style stamp tiles with a lighthouse, a falcon and Acadia lettering."),
 221: ("Blueberry Blossoms", "Blueberries with small blue flowers and green leaves."),
 222: ("Orange Pop", "A pink Orange Pop badge with an orange popsicle and an orange slice."),
 223: ("Hello Citrus", "A smiling orange character in sneakers with bold Hello, Citrus! lettering."),
 224: ("Wild Luxe", "A leopard resting under a Wild Luxe label between palm leaves."),
 225: ("Mosaic Cat", "A mosaic-style cat in a red headscarf and boots among pineapples."),
 226: ("Just Coasting", "A black-and-white cat on a wave with Just Coasting lettering in a sticker style."),
 227: ("Splashing Dachshund", "A happy dachshund with water bubbles."),
 228: ("Bearly Awake", "A bear sleeping in a hammock between trees with Bearly Awake lettering."),
 229: ("Corgi Fish", "A corgi lying on a colorful fish."),
 230: ("Mad Scientist", "A wild-haired scientist in a lab coat holding a flask and a wrench."),
 231: ("Ready to Learn", "A smiling pink backpack and pencil with Ready to Learn! lettering."),
 232: ("Keep Going A+", "A smiling pencil with Keep Going! A+ lettering."),
 233: ("Koi Waves", "Black and red koi fish with cherry blossoms over blue Japanese waves."),
 234: ("Punk Mode", "A cassette with checkerboard, safety pins and lightning bolts and Punk Mode lettering."),
 235: ("Class Vibes", "A school collage with a notebook, smiley, backpack, rainbow and Class Vibes lettering."),
 236: ("No Rules", "A rock-hand gesture with checkerboard, safety pins and lightning bolts and No Rules lettering."),
 237: ("Butterflies", "A blue butterfly and a pink butterfly with green sprigs and small flowers."),
 238: ("Mine Now", "An orange cat running off with a pizza slice and Mine Now lettering."),
 239: ("No Risk No Story", "Bold red and pink No Risk No Story lettering with green sparkles."),
 240: ("Busy Napping", "A sleeping cat on a keyboard with Busy Napping lettering."),
 241: ("Good Vibes", "Bold pink and red Good Vibes lettering with a sun and sparkles."),
 242: ("USA Map", "A blue USA map with white stars and USA lettering."),
 243: ("Land of the Free", "A waving American flag with Land of the Free lettering."),
 244: ("Coast Collage", "A bridge, a sailboat, a vintage car and palm trees with a Coast banner."),
 245: ("Zero Regrets", "A pink snack bag with a black-and-white cat diving in and Zero Regrets lettering."),
 246: ("Perfect Fit", "A black cat in a pink box with a measuring tape and Perfect Fit lettering."),
 247: ("Sardines", "Blue and cream sardines with lemon slices and sparkles."),
 248: ("Cowgirl Skull", "A cow skull in a pink cowboy hat with horseshoes, cacti and flowers."),
 249: ("Enchanted Night", "A ghost, pumpkins, bats and flowers around Enchanted Night lettering in purple."),
 250: ("Witchy Mood", "A skeleton hand, pumpkins and flowers around Witchy Mood lettering in purple."),
 251: ("Study Break", "Books, a calculator, a rainbow and a smiley with Study Break lettering."),
 252: ("Daisy Bouquet", "Three colorful daisies with small flowers and leaves."),
 253: ("Blueberry Vine", "A vine of blue and purple blueberries with leaves and small flowers."),
 254: ("Leopard Lounge", "A leopard in sunglasses with a feather fan on a pink sofa."),
 255: ("Tiger Cocktail", "A tiger with a pink feather boa lounging in a cocktail glass."),
 256: ("Kitsune Mask", "A kitsune fox mask with fans, clouds and cherry blossoms."),
 257: ("Pink Butterfly", "A large pink and cream butterfly with small accents."),
 258: ("Ribbon Bows", "Three bows in blue, pink and lime with small flowers."),
 259: ("Fishy Business", "A black cat in a sardine tin with Fishy Business lettering."),
 260: ("Royal Mess", "A black cat wrapped in toilet paper with Royal Mess lettering."),
 261: ("Leopard Stretch", "A stretching leopard with yellow stars and pink flowers."),
 262: ("Leopard Bolt", "A leopard head in a pink star with lightning bolts."),
 263: ("You Got This", "An open smiling book with You Got This! lettering."),
 264: ("Origami Frog", "A green geometric origami-style frog."),
 265: ("Green Monster", "A one-eyed green monster with claws on a target-style panel."),
 266: ("Peace Rainbow", "A peace-sign hand, smiley sun, rainbow and flowers in neon colors."),
 267: ("Portal Cat", "A tabby cat walking through two colorful portals."),
 268: ("Burger Cat", "A fluffy white cat with a glowing burger on its belly."),
 269: ("Winging It", "A goose with a map and Winging It lettering in a sticker style."),
 270: ("Neon Mushrooms", "Colorful neon mushrooms with moons and stars."),
 271: ("Working My Magic", "A wizard at a retro computer with Working My Magic lettering."),
 272: ("Cat on Fish", "A black-and-white cat lying on a colorful fish."),
 273: ("Robot Cowboy", "A skull-faced robot cowboy with two revolvers."),
 274: ("Lucky Cat Torii", "A lucky cat under a torii gate with waves, fans and blossoms."),
 275: ("Vampire", "A vampire with a high collar and clawed hands."),
 276: ("Confetti Dots", "Scattered red, white and blue dots."),
}


def _rows(path):
    wb = openpyxl.load_workbook(ROOT / path, read_only=True)
    return [r for r in wb.worksheets[0].iter_rows(values_only=True)][1:]


def _qnum(sku):
    return int(re.search(r"q(\d+)$", sku).group(1))


def _short_model(m):
    return re.sub(r"^Apple ", "", m).strip()


def _print_sentences_215(desc):
    out = []
    for s in re.split(r"(?<=\.)\s+", desc):
        sl = s.lower()
        if sl.startswith("clear magnetic phone case") or "magnetic ring" in sl or "reinforced corners" in sl:
            continue
        out.append(s)
    return " ".join(out).strip()


def _name_215(title):
    m = re.search(r"MagSafe, (.+)$", title)
    if m:
        return m.group(1).strip()
    m = re.match(r"MOBILIUS (.+?) Case for", title)
    return m.group(1).strip() if m else None


def load():
    products = []
    # ---- 215
    a = {r[1]: r for r in _rows("215 Amazon 29.09 — модели исправлено.xlsx")}
    for r in _rows("215 SKU, принты, описания.xlsx"):
        sku = r[0]
        old = a[sku][0]
        photos = [u for u in r[6:13] if u]
        products.append(dict(series="215", sku=sku, sku_old=old, model=_short_model(r[3]), color="Clear",
                             finish="Clear", q=_qnum(sku), title=r[1], name=_name_215(r[1]),
                             desc=_print_sentences_215(r[2]), photos=photos))
    # ---- 223
    for r in _rows("Таблица 223 Амазон.xlsx"):
        sku = r[0]
        photos = [u for u in r[5:13] if u]
        old = re.search(r"/([^/]+?)_\d+\.jpg$", photos[0]).group(1) if photos else ""
        m = re.search(r"MagSafe, (.*?) Design,", r[1])
        d = r[2].split("MagSafe. ", 1)[-1]
        products.append(dict(series="223", sku=sku, sku_old=old, model=r[4], color=r[3], finish="Matte",
                             q=_qnum(sku), title=r[1], name=m.group(1), desc=d, photos=photos))
    # ---- 406 (одно актуальное SKU — как в каталоге)
    for q, (n, d) in P406.items():
        products.append(dict(series="406", sku=f"406xxx-002Ru001q{q}", sku_old="", model="AirPods Pro",
                             color="Black", finish="TBC", q=q, title="", name=n, desc=d,
                             photos=[]))
    return products


def model_sort_key(m):
    return MODEL_ORDER.index(m) if m in MODEL_ORDER else 99


if __name__ == "__main__":
    ps = load()
    import collections
    print(len(ps), collections.Counter(p["series"] for p in ps))
    bad = [p for p in ps if not p["name"] or not p["desc"]]
    print("no name/desc:", len(bad), [(p["sku"], p["name"]) for p in bad[:5]])
    # 215 name vs 223 name
    n223 = {p["q"]: p["name"] for p in ps if p["series"] == "223"}
    diff = {}
    for p in ps:
        if p["series"] == "215":
            diff.setdefault(p["q"], set()).add(p["name"])
    multi = {q: v for q, v in diff.items() if len(v) > 1}
    print("215 q with >1 names:", len(multi), list(multi.items())[:3])
    print("215 not in 223:", sorted(set(diff) - set(n223)))
    print(sorted((q, sorted(v)) for q, v in diff.items() if q not in n223))
    nomatch = [(q, v, n223[q]) for q, v in diff.items() if q in n223 and n223[q] not in v]
    print("name mismatch", len(nomatch), nomatch[:5])
    for p in ps:
        if p["series"] == "215" and p["q"] in (1, 4, 13):
            print(p["model"], "|", p["name"], "|", p["desc"]);
    print(sum(1 for p in ps if p["series"] != "406" and not p["photos"]))
