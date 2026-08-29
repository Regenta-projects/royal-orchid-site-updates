"""
ROHL Site-Update Builder engine.
Opens a hotel's template .pptx and applies a site submission:
  - current photo -> CURRENT MONTH box (last month auto-rotates to PREVIOUS)
  - render image  -> RENDERED box (only on 3-column space slides)
  - up to 3 tappable video links (walkthrough / detail / snag)
  - feedback text + site-visit date footer
  - optional room-type variant -> clones the space slide on first sight
Auto-detects 3-column vs 2-column layouts. Locates zones by label/geometry,
never by fragile shape ids.
"""
from pptx import Presentation
from pptx.util import Emu
from PIL import Image
import copy

IN = 914400

# form space value -> keyword found in the slide title "04 · ..."
SPACE_KEY = {
    "facade": "façade", "reception": "reception", "dining": "dining",
    "banquet": "banquet", "guest room": "guest room", "bathroom": "bathroom",
    "corridor": "corridor", "kitchen": "kitchen", "back of house": "back of house",
    "electrical": "electrical", "stp": "stp", "fire": "fire",
}

def _txt(sh):
    return sh.text_frame.text.strip() if sh.has_text_frame else ""

def _cx(sh):
    return sh.left + sh.width / 2

def _set_text(sh, txt, hyperlink=None):
    p = sh.text_frame.paragraphs[0]
    if p.runs:
        r = p.runs[0]; r.text = txt
        for extra in p.runs[1:]:
            extra.text = ""
    else:
        r = p.add_run(); r.text = txt
    if hyperlink:
        r.hyperlink.address = hyperlink
    return r

def _img_size(path):
    with Image.open(path) as im:
        return im.size

# ---- slide / column discovery -------------------------------------------------

def _title(slide):
    for sh in slide.shapes:
        if _txt(sh).startswith("04 ·"):
            return sh
    return None

def _subtitle(slide):
    # the descriptive line under the title (Text 2), first small text below title
    t = _title(slide)
    if not t:
        return None
    cands = [sh for sh in slide.shapes
             if sh.has_text_frame and sh is not t and sh.top > t.top
             and abs(sh.left - t.left) < 0.5 * IN and sh.top < 1.5 * IN]
    return min(cands, key=lambda s: s.top) if cands else None

def find_columns(slide):
    """Return {'RENDER'|'PREV'|'CUR': box_shape} by matching labels to tall rects."""
    label_cx = {}
    for sh in slide.shapes:
        t = _txt(sh)
        if t == "RENDERED — DESIGN INTENT": label_cx["RENDER"] = _cx(sh)
        elif t == "PREVIOUS MONTH": label_cx["PREV"] = _cx(sh)
        elif t == "CURRENT MONTH": label_cx["CUR"] = _cx(sh)
    rects = [sh for sh in slide.shapes
             if sh.width and sh.height and sh.height > 4 * IN and sh.width > 3 * IN
             and not _txt(sh)]
    cols = {}
    for key, cx in label_cx.items():
        cols[key] = min(rects, key=lambda r: abs(_cx(r) - cx))
    return cols

def _placeholder_in(slide, box):
    """The 'Photograph awaited' text centred inside a box (may be blank after fill)."""
    for sh in slide.shapes:
        if not sh.has_text_frame or sh.height > 1.2 * IN:
            continue
        c = _cx(sh); cy = sh.top + sh.height / 2
        if box.left < c < box.left + box.width and box.top < cy < box.top + box.height:
            if _txt(sh) in ("Photograph awaited", ""):
                return sh
    return None

def _fit(box, iw, ih, inset=0.06 * IN):
    bw, bh = box.width - 2 * inset, box.height - 2 * inset
    s = min(bw / iw, bh / ih)
    nw, nh = iw * s, ih * s
    x = box.left + inset + (bw - nw) / 2
    y = box.top + inset + (bh - nh) / 2
    return int(x), int(y), int(nw), int(nh)

def _find_pic(slide, name):
    for sh in slide.shapes:
        if sh.name == name:
            return sh
    return None

def _place(slide, box, image_path, name):
    iw, ih = _img_size(image_path)
    x, y, w, h = _fit(box, iw, ih)
    pic = slide.shapes.add_picture(image_path, Emu(x), Emu(y), Emu(w), Emu(h))
    pic.name = name
    ph = _placeholder_in(slide, box)
    if ph is not None:
        _set_text(ph, "")
    return pic

# ---- the update actions -------------------------------------------------------

def set_current_photo(slide, cols, photo_path):
    # rotate existing current -> previous
    if "PREV" in cols:
        old_prev = _find_pic(slide, "PREVIMG")
        if old_prev is not None:
            old_prev._element.getparent().remove(old_prev._element)
        cur = _find_pic(slide, "CURIMG")
        if cur is not None:
            box = cols["PREV"]
            x, y, w, h = _fit(box, cur.width, cur.height)
            cur.left, cur.top, cur.width, cur.height = Emu(x), Emu(y), Emu(w), Emu(h)
            cur.name = "PREVIMG"
            ph = _placeholder_in(slide, box)
            if ph is not None:
                _set_text(ph, "")
    _place(slide, cols["CUR"], photo_path, "CURIMG")

def set_render(slide, cols, render_path):
    if "RENDER" not in cols or not render_path:
        return
    old = _find_pic(slide, "RENDERIMG")
    if old is not None:
        old._element.getparent().remove(old._element)
    _place(slide, cols["RENDER"], render_path, "RENDERIMG")

def set_videos(slide, videos):
    cells = [sh for sh in slide.shapes
             if sh.has_text_frame and abs(sh.left - 2.87 * IN) < 0.35 * IN
             and 8.2 * IN < sh.top < 9.7 * IN]
    cells.sort(key=lambda s: s.top)  # walkthrough, detail, snag
    for cell, key in zip(cells, ("walkthrough", "detail", "snag")):
        url = videos.get(key)
        if url:
            _set_text(cell, "Watch ▶", hyperlink=url)

def set_feedback(slide, text):
    lab = next((sh for sh in slide.shapes if _txt(sh) == "FEEDBACK FOR THIS SPACE"), None)
    if not lab:
        return
    cands = [sh for sh in slide.shapes if sh.has_text_frame and sh is not lab
             and abs(sh.left - lab.left) < 1 * IN and sh.top > lab.top]
    if cands:
        _set_text(max(cands, key=lambda s: s.width * s.height), text)

def set_footer(slide, date_str):
    for sh in slide.shapes:
        if _txt(sh).startswith("Last site visit"):
            _set_text(sh, f"Last site visit: {date_str}   ·   Updated automatically from site form")
            return

# ---- room-type variants -------------------------------------------------------

def _reset_media(slide, cols):
    for pic_name in ("CURIMG", "PREVIMG", "RENDERIMG"):
        p = _find_pic(slide, pic_name)
        if p is not None:
            p._element.getparent().remove(p._element)
    for key, box in cols.items():
        ph = _placeholder_in(slide, box)
        if ph is not None:
            _set_text(ph, "Photograph awaited")

def clone_variant(prs, base_idx, variant):
    src = prs.slides[base_idx]
    new = prs.slides.add_slide(src.slide_layout)
    for sh in list(new.shapes):
        sh._element.getparent().remove(sh._element)
    for sh in src.shapes:
        new.shapes._spTree.append(copy.deepcopy(sh._element))
    # relabel: title base keeps the space, subtitle carries the variant
    t = _title(new)
    if t:
        base = _txt(t).split("—")[0].rstrip(" -")   # "04 · Guest room"
        _set_text(t, f"{base} — {variant}")
    sub = _subtitle(new)
    if sub:
        _set_text(sub, f"{variant}  ·  site update")
    _reset_media(new, find_columns(new))
    # move newly-added (last) slide to just after the base slide's section
    ids = prs.slides._sldIdLst
    moved = list(ids)[-1]
    ids.remove(moved)
    ids.insert(base_idx + 1, moved)
    return new

def find_space_slide(prs, space_value, variant=None):
    key = SPACE_KEY.get(space_value.strip().lower(), space_value.strip().lower())
    matches = [(i, s) for i, s in enumerate(prs.slides)
               if _title(s) and key in _txt(_title(s)).lower()]
    if not matches:
        return None
    base_idx, base = matches[0]
    if not variant:
        return base
    for i, s in matches:
        sub = _subtitle(s)
        if (sub and variant.lower() in _txt(sub).lower()) or variant.lower() in _txt(_title(s)).lower():
            return s
    return clone_variant(prs, base_idx, variant)

# ---- public entry point -------------------------------------------------------

def apply_submission(prs, sub):
    """sub keys: space, variant?, photo?, render?, videos?{}, feedback?, date?"""
    slide = find_space_slide(prs, sub["space"], sub.get("variant"))
    if slide is None:
        raise ValueError(f"No slide for space {sub['space']!r}")
    cols = find_columns(slide)
    if sub.get("render"):
        set_render(slide, cols, sub["render"])
    if sub.get("photo"):
        set_current_photo(slide, cols, sub["photo"])
    if sub.get("videos"):
        set_videos(slide, sub["videos"])
    if sub.get("feedback"):
        set_feedback(slide, sub["feedback"])
    if sub.get("date"):
        set_footer(slide, sub["date"])
    return slide
