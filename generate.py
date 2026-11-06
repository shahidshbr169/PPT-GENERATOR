#!/usr/bin/env python3
"""
Medical PPT generator.

1. Asks gpt-oss-120b (via Groq) for a structured slide outline on a medical topic.
2. Finds an image for each slide from: SMART, Open-i, MedPix, CDC PHIL, NIH BioArt.
3. Builds a .pptx with python-pptx and registers it in docs/decks.json
   so the GitHub Pages site can list it.

Usage:
    GROQ_API_KEY=... python generate.py --topic "Myocardial infarction" --slides 10 --audience "medical students"
"""
import argparse
import datetime as dt
import io
import json
import os
import pathlib
import re
import sys
import time

import requests
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

ROOT = pathlib.Path(__file__).parent
DOCS = ROOT / "docs"
DECKS = DOCS / "decks"
INDEX = DOCS / "decks.json"

MODEL = "openai/gpt-oss-120b"
UA = {"User-Agent": "medical-ppt-generator/1.0 (educational, non-commercial)"}
IMG_EXT = re.compile(r"\.(png|jpe?g|gif|webp)(\?.*)?$", re.I)

NAVY = RGBColor(0x12, 0x2B, 0x4A)
TEAL = RGBColor(0x0E, 0x8A, 0x8A)
GREY = RGBColor(0x55, 0x5B, 0x66)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


# --------------------------------------------------------------------------
# 1. Slide text from Groq
# --------------------------------------------------------------------------
SYSTEM = (
    "You are a medical educator who writes accurate, concise teaching slides. "
    "Reply with a single JSON object and nothing else."
)

PROMPT = """Create a slide deck on: "{topic}"
Audience: {audience}
Number of content slides: {n} (do not count the title or credits slide).

Return JSON with exactly this shape:
{{
  "title": "deck title",
  "subtitle": "one-line subtitle",
  "slides": [
    {{
      "title": "slide title",
      "bullets": ["3-5 short bullets, max 18 words each"],
      "notes": "2-4 sentences of speaker notes",
      "image_query": "2-4 word search term for a suitable image",
      "image_source": "smart | openi | phil | medpix | bioart | none"
    }}
  ]
}}

Guidance for image_source:
- smart: anatomy / physiology / pathway diagrams, cartoon-style illustrations
- bioart: cells, organelles, viruses, molecules, lab-style vectors
- openi: figures, charts, histology, radiology, published-paper images
- medpix: real clinical cases, CT / MRI / X-ray teaching images
- phil: microbes, infectious disease, public-health photographs
- none: slides where an image would not help (e.g. summary)
Cover definition, epidemiology, pathophysiology, clinical features, diagnosis,
management, complications, and key takeaways where relevant. Be factually careful."""


def groq_outline(topic: str, n: int, audience: str) -> dict:
    key = os.environ.get("GROQ_API_KEY")
    if not key:
        sys.exit("GROQ_API_KEY is not set")
    body = {
        "model": MODEL,
        "temperature": 0.4,
        "messages": [
            {"role": "system", "content": SYSTEM},
            {"role": "user", "content": PROMPT.format(topic=topic, n=n, audience=audience)},
        ],
        "response_format": {"type": "json_object"},
    }
    last = None
    for attempt in range(4):
        try:
            r = requests.post(
                "https://api.groq.com/openai/v1/chat/completions",
                headers={"Authorization": f"Bearer {key}"},
                json=body,
                timeout=180,
            )
            if r.status_code == 429 or r.status_code >= 500:
                raise RuntimeError(f"Groq HTTP {r.status_code}: {r.text[:200]}")
            r.raise_for_status()
            text = r.json()["choices"][0]["message"]["content"]
            m = re.search(r"\{.*\}", text, re.S)
            data = json.loads(m.group(0) if m else text)
            assert data.get("slides"), "no slides in response"
            return data
        except Exception as e:  # noqa: BLE001
            last = e
            print(f"[groq] attempt {attempt + 1} failed: {e}", file=sys.stderr)
            time.sleep(3 * (attempt + 1))
    sys.exit(f"Groq failed: {last}")


# --------------------------------------------------------------------------
# 2. Image search adapters. Each yields dicts: url, source, credit, license
# --------------------------------------------------------------------------
def _get(url, **kw):
    kw.setdefault("headers", UA)
    kw.setdefault("timeout", 30)
    return requests.get(url, **kw)


def _walk_image_urls(obj, base=""):
    """Recursively pull anything that looks like an image URL out of JSON."""
    if isinstance(obj, dict):
        for v in obj.values():
            yield from _walk_image_urls(v, base)
    elif isinstance(obj, list):
        for v in obj:
            yield from _walk_image_urls(v, base)
    elif isinstance(obj, str) and IMG_EXT.search(obj):
        yield obj if obj.startswith("http") else base + obj


def search_openi(q, coll=None, label="Open-i (NLM)"):
    """Open-i JSON API. coll='mpx' restricts to the MedPix collection."""
    params = {"query": q, "m": 1, "n": 10}
    if coll:
        params["coll"] = coll
    r = _get("https://openi.nlm.nih.gov/api/search", params=params)
    r.raise_for_status()
    for it in r.json().get("list", []):
        path = it.get("imgLarge") or it.get("imgThumb")
        if not path:
            continue
        yield {
            "url": "https://openi.nlm.nih.gov" + path,
            "source": label,
            "credit": (it.get("title") or q)[:110],
            "license": "See source article; open access",
        }


def search_medpix(q):
    # MedPix cases are indexed inside Open-i (collection "mpx"), which has a stable API.
    yield from search_openi(q, coll="mpx", label="MedPix (NLM)")


def search_smart(q):
    """SMART has no public API; scrape the WordPress search page (best effort)."""
    r = _get("https://smart.servier.com/", params={"s": q})
    r.raise_for_status()
    seen = set()
    for m in re.finditer(r'https://smart\.servier\.com/wp-content/uploads/[^"\'\s)]+?\.(?:png|jpe?g)', r.text, re.I):
        url = re.sub(r"-\d+x\d+(?=\.\w+$)", "", m.group(0))  # strip thumbnail size suffix
        if url in seen:
            continue
        seen.add(url)
        yield {
            "url": url,
            "source": "Servier Medical Art",
            "credit": "Servier Medical Art (smart.servier.com), CC BY 4.0",
            "license": "CC BY 4.0",
        }


def search_bioart(q):
    """NIH BioArt: best-effort JSON endpoint; any image URL found is used."""
    r = _get("https://bioart.niaid.nih.gov/api/bioarts", params={"search": q})
    r.raise_for_status()
    for url in _walk_image_urls(r.json(), base="https://bioart.niaid.nih.gov"):
        yield {
            "url": url,
            "source": "NIH BioArt Source",
            "credit": "NIH BioArt Source (bioart.niaid.nih.gov)",
            "license": "Public domain / NIH BioArt terms",
        }


def search_phil(q):
    """CDC PHIL: best-effort via the CDC content syndication media API."""
    r = _get(
        "https://tools.cdc.gov/api/v2/resources/media",
        params={"q": q, "mediatype": "Image", "max": 10},
    )
    r.raise_for_status()
    for url in _walk_image_urls(r.json()):
        yield {
            "url": url,
            "source": "CDC Public Health Image Library",
            "credit": "CDC / Public Health Image Library (phil.cdc.gov)",
            "license": "Public domain (check individual PHIL record for credits)",
        }


ADAPTERS = {
    "smart": search_smart,
    "openi": search_openi,
    "medpix": search_medpix,
    "phil": search_phil,
    "bioart": search_bioart,
}


def fetch_image(url):
    """Download and validate an image; return PNG bytes or None."""
    try:
        r = _get(url, timeout=40)
        r.raise_for_status()
        im = Image.open(io.BytesIO(r.content))
        im.load()
        if im.width < 200 or im.height < 150:
            return None
        if im.mode not in ("RGB", "RGBA"):
            im = im.convert("RGB")
        im.thumbnail((1600, 1600))
        out = io.BytesIO()
        im.save(out, "PNG")
        return out.getvalue()
    except Exception:  # noqa: BLE001
        return None


def find_image(query, preferred, topic, used):
    order = [preferred] if preferred in ADAPTERS else []
    order += [k for k in ADAPTERS if k not in order]
    for q in dict.fromkeys([query, topic]):  # try the slide query, then the deck topic
        for name in order:
            try:
                for cand in ADAPTERS[name](q):
                    if cand["url"] in used:
                        continue
                    data = fetch_image(cand["url"])
                    if data:
                        used.add(cand["url"])
                        print(f"  image [{name}] {q!r} -> {cand['url']}")
                        return data, cand
            except Exception as e:  # noqa: BLE001
                print(f"  [{name}] failed for {q!r}: {e}", file=sys.stderr)
    return None, None


# --------------------------------------------------------------------------
# 3. Build the PPTX
# --------------------------------------------------------------------------
def _fill(slide, rgb):
    f = slide.background.fill
    f.solid()
    f.fore_color.rgb = rgb


def _text(slide, x, y, w, h, text, size, color, bold=False):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    return tb


def build_pptx(outline, images, path):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]

    # Title slide
    s = prs.slides.add_slide(blank)
    _fill(s, NAVY)
    _text(s, 0.9, 2.4, 11.5, 1.6, outline.get("title", "Medical Topic"), 44, WHITE, True)
    _text(s, 0.9, 4.1, 11.5, 1, outline.get("subtitle", ""), 22, RGBColor(0xBF, 0xDD, 0xE8))

    credits = []
    for sd, (img, meta) in zip(outline["slides"], images):
        s = prs.slides.add_slide(blank)
        bar = s.shapes.add_shape(1, 0, 0, prs.slide_width, Inches(1.1))
        bar.fill.solid()
        bar.fill.fore_color.rgb = NAVY
        bar.line.fill.background()
        _text(s, 0.5, 0.18, 12.3, 0.8, sd.get("title", ""), 30, WHITE, True)

        has_img = img is not None
        text_w = 6.9 if has_img else 12.3
        tb = s.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(text_w), Inches(5.3))
        tf = tb.text_frame
        tf.word_wrap = True
        for i, b in enumerate(sd.get("bullets", [])[:6]):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            r = p.add_run()
            r.text = "\u2022  " + b
            r.font.size = Pt(20)
            r.font.color.rgb = RGBColor(0x22, 0x26, 0x2E)
            p.space_after = Pt(12)

        if has_img:
            box_x, box_y, box_w, box_h = 7.7, 1.5, 5.1, 4.9
            im = Image.open(io.BytesIO(img))
            scale = min(box_w / im.width, box_h / im.height)
            w, h = im.width * scale, im.height * scale
            x = box_x + (box_w - w) / 2
            s.shapes.add_picture(io.BytesIO(img), Inches(x), Inches(box_y), Inches(w), Inches(h))
            _text(s, box_x, 6.5, box_w, 0.6, f"Image: {meta['source']}", 10, GREY)
            credits.append((sd.get("title", ""), meta))

        if sd.get("notes"):
            s.notes_slide.notes_text_frame.text = sd["notes"]

    # Credits slide
    s = prs.slides.add_slide(blank)
    _text(s, 0.6, 0.4, 12, 0.8, "Image credits", 30, NAVY, True)
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.3), Inches(12.1), Inches(5.6))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, (title, m) in enumerate(credits or [("", {"credit": "No images used", "license": "", "url": ""})]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run()
        r.text = f"{title}: {m['credit']} [{m['license']}] {m['url']}".strip()
        r.font.size = Pt(10)
        r.font.color.rgb = GREY
    note = tf.add_paragraph()
    nr = note.add_run()
    nr.text = (
        "\nAI-generated content for education only. Verify clinical facts against "
        "current guidelines. Check each image's licence before reuse."
    )
    nr.font.size = Pt(11)
    nr.font.bold = True
    nr.font.color.rgb = TEAL

    prs.save(path)


# --------------------------------------------------------------------------
# 4. Register in decks.json
# --------------------------------------------------------------------------
def register(topic, outline, filename, n_images):
    decks = []
    if INDEX.exists():
        try:
            decks = json.loads(INDEX.read_text())
        except json.JSONDecodeError:
            decks = []
    decks.insert(
        0,
        {
            "topic": topic,
            "title": outline.get("title", topic),
            "file": f"decks/{filename}",
            "slides": len(outline["slides"]),
            "images": n_images,
            "created": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
        },
    )
    INDEX.write_text(json.dumps(decks[:100], indent=2))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--topic", required=True)
    ap.add_argument("--slides", type=int, default=10)
    ap.add_argument("--audience", default="medical students")
    a = ap.parse_args()

    topic = a.topic.strip()[:200]
    n = max(4, min(a.slides, 20))
    print(f"Topic: {topic} | slides: {n} | audience: {a.audience}")

    outline = groq_outline(topic, n, a.audience)
    used, images = set(), []
    for sd in outline["slides"]:
        print(f"- {sd.get('title')}")
        if sd.get("image_source") == "none":
            images.append((None, None))
            continue
        images.append(find_image(sd.get("image_query") or sd.get("title", topic), sd.get("image_source"), topic, used))

    DECKS.mkdir(parents=True, exist_ok=True)
    slug = re.sub(r"[^a-z0-9]+", "-", topic.lower()).strip("-")[:50] or "deck"
    filename = f"{slug}-{dt.datetime.now(dt.timezone.utc):%Y%m%d-%H%M%S}.pptx"
    build_pptx(outline, images, DECKS / filename)
    register(topic, outline, filename, sum(1 for i, _ in images if i))
    print(f"Saved docs/decks/{filename}")


if __name__ == "__main__":
    main()
