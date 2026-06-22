#!/usr/bin/env python3
"""Generate the Josef Sudek presentation (Option A) as a .pptx file.

Theme: "The World in One Hand: How Limitation Became Sudek's Style".
Layout: one photograph per slide (20 slides total), with an image placeholder
box on each photo slide and the English speaker notes attached to every slide.

The 15 photographs, the four trait sections, the slide order, and the speaker
notes are kept in sync with `presentation-plan.md` and `presentation-script.md`.

Usage:
    python build_pptx.py
    python build_pptx.py --output path/to/file.pptx

Replace the placeholder boxes with real, public-domain or licensed images
before presenting, and fill in the credits slide.
"""
from __future__ import annotations

import argparse
import os

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

# 16:9 canvas
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# Folder that holds the real photographs. Drop files here and re-run the
# script: any matching file is embedded automatically, otherwise the slide
# falls back to a labelled placeholder box.
#   - portrait.(jpg|jpeg|png|gif|bmp|tif|tiff|webp)   -> photographer portrait
#   - photo01.(...) ... photo15.(...)                 -> the 15 photographs
DEFAULT_IMAGE_DIR = "images"
IMAGE_DIR = DEFAULT_IMAGE_DIR
IMAGE_EXTENSIONS = (".jpg", ".jpeg", ".png", ".gif", ".bmp",
                    ".tif", ".tiff", ".webp")

FONT = "Calibri"

# Quiet, low-saturation palette echoing Sudek's mood
INK = RGBColor(0x20, 0x20, 0x20)        # near-black text
MUTED = RGBColor(0x6E, 0x6A, 0x63)      # warm grey
PAPER = RGBColor(0xF4, 0xF2, 0xED)      # warm off-white background
PAPER_DEEP = RGBColor(0xE5, 0xDF, 0xD2)  # slightly deeper warm tone for gradient
PLACEHOLDER = RGBColor(0xDF, 0xDB, 0xD2)  # image placeholder fill
TAG = RGBColor(0x8A, 0x6D, 0x3B)        # warm brown accent for trait tags

DEFAULT_OUTPUT = "Josef-Sudek-Presentation.pptx"

# Each photo slide: (number, title, year, trait tag, speaker notes).
# Speaker notes are a list of (text, kind) segments where kind is either
# "must" (core lines you should always say) or "opt" (optional lines you can
# add to fill time, or skip if you are running long). _set_photo_notes renders
# "must" segments in dark text and "opt" segments in grey italic, with a legend.
PHOTOS = [
    (1, "Saint Vitus Cathedral, interior", "c. 1924–1928", "SLOW", [
        ("First, SLOW. The view camera demanded long exposures, so Sudek learned "
         "to wait. A long exposure gathers soft, diffused light the eye barely "
         "notices \u2014 light becomes almost solid. This patience is the "
         "foundation of everything that follows.", "must"),
        ("Look how there are almost no hard shadows \u2014 the light seems to "
         "settle into the stone. Exposures like this could run for many minutes.",
         "opt"),
    ]),
    (2, "Prague Panorama, misty city", "c. 1950s–1960s", "SLOW", [
        ("Here, slowness turns fog and damp air into something dreamlike. This is "
         "from his panoramic work \u2014 a wide, demanding format. The mist isn't "
         "a problem to fix; it is the subject.", "must"),
        ("That long horizontal frame was unusual for its time, and it makes the "
         "city feel like it goes on forever. The damp Prague air was almost a "
         "collaborator for him.", "opt"),
    ]),
    (3, "Prague at Night", "c. 1950s", "SLOW", [
        ("This near-darkness was only possible with a long exposure on a tripod. "
         "The empty, glowing street shows how Sudek used time itself as a tool "
         "\u2014 letting faint light slowly build into a quiet scene.", "must"),
        ("If you tried to shoot this by hand, you'd get nothing but a black "
         "frame. The few points of light almost feel like they're breathing.",
         "opt"),
    ]),
    (4, "Mionší Forest, mist in the woods", "c. 1950s", "SLOW", [
        ("In a misty forest, the stillness becomes visible. The soft grey light "
         "and motionless trees are the trace of a long, patient exposure. "
         "Slowness was never a weakness \u2014 it was how he found poetry in "
         "ordinary light.", "must"),
        ("There's no wind, no movement, nothing dramatic \u2014 and that's "
         "exactly the point. He's photographing the quiet itself.", "opt"),
    ]),
    (5, "The Window of My Studio", "c. 1940–1948", "CLOSE", [
        ("Second, CLOSE. Because movement was difficult, Sudek worked within "
         "arm's reach. His most famous motif: the window of his studio, covered "
         "in mist and droplets. A single pane of glass became an entire "
         "universe.", "must"),
        ("He made dozens of versions of this window over the years, often just "
         "looking out at his garden. He didn't need to travel to find a world "
         "\u2014 it was right in front of him.", "opt"),
    ]),
    (6, "The Last Rose", "c. 1956", "CLOSE", [
        ("On his table, a single fading rose. Seen this closely, an ordinary "
         "flower becomes a meditation on beauty and time \u2014 one of his most "
         "poetic still lifes.", "must"),
        ("It's not a fresh, perfect bloom \u2014 it's wilting, and that's the "
         "whole point. He found beauty in the moment things begin to fade.",
         "opt"),
    ]),
    (7, "Glass and Egg (Labyrinths)", "c. 1950s", "CLOSE", [
        ("In his still lifes, simple objects hold tiny, glowing worlds of light. "
         "A glass, an egg \u2014 close up, their surfaces become mysterious. "
         "Closeness was his limitation, but also his way of seeing the "
         "extraordinary in the plain.", "must"),
        ("He called some of these his 'Labyrinths.' Spend a moment on the "
         "reflections \u2014 there's a whole little landscape inside that glass.",
         "opt"),
    ]),
    (8, "Still life with Bread and Glass", "c. 1950s", "CLOSE", [
        ("Bread and a glass on a dark table. Humble, everyday things \u2014 yet "
         "lit so gently they feel almost sacred. Intimacy as a method: the closer "
         "he looked, the more meaning he found.", "must"),
        ("This is the kind of food you'd find in any modest Prague kitchen. Sudek "
         "treats it with the same care a painter might give a portrait.", "opt"),
    ]),
    (9, "Window \u2014 view to the garden, summer", "c. 1940s", "REPEATED", [
        ("Third, REPEATED. Sudek returned to the same subjects for years. Here is "
         "the view from his studio window onto the garden in one season, full and "
         "green.", "must"),
        ("Keep this exact frame in your memory for a second \u2014 because the "
         "next image is the very same window. He photographed it over and over "
         "across the seasons.", "opt"),
    ]),
    (10, "Window \u2014 frost, winter", "c. 1940s–1950s", "REPEATED", [
        ("The same window in winter, the glass laced with frost. Same frame, "
         "completely different mood. Repetition let him discover endless "
         "variation inside a single point of view.", "must"),
        ("Same window, same spot \u2014 only the season has changed, and suddenly "
         "it's a different world. This is what patience with one subject can "
         "reveal.", "opt"),
    ]),
    (11, "A Walk in the Magic Garden I", "c. 1954", "REPEATED", [
        ("He gave the same devotion to his small, wild garden \u2014 his 'magic "
         "garden.' He photographed it again and again, finding life in its "
         "overgrown corners.", "must"),
        ("It was just a modest, tangled little garden in Prague, nothing grand. "
         "But to him it was endless.", "opt"),
    ]),
    (12, "A Walk in the Magic Garden II", "c. 1960s", "REPEATED", [
        ("The garden again, at a different hour and light. Repetition was not a "
         "lack of ideas \u2014 it was depth. By doing less, again and again, he "
         "saw more than photographers who do everything once.", "must"),
        ("Same garden, different day, different light \u2014 and a completely "
         "different feeling. For Sudek, returning was a form of looking deeper.",
         "opt"),
    ]),
    (13, "Prague Panorama \u2014 empty street/square", "c. 1950s–1960s", "SOLITARY", [
        ("Fourth, SOLITARY. Sudek's images are almost always empty of people. "
         "This empty square feels quiet, inward, a little melancholy \u2014 a "
         "whole city holding its breath.", "must"),
        ("Notice there's not a single person in the frame. That emptiness isn't "
         "lonely exactly \u2014 it leaves room for you, the viewer.", "opt"),
    ]),
    (14, "Veteran from the Invalidovna", "c. 1922–1927", "SOLITARY", [
        ("This earlier, documentary work shows a war veteran at the Invalidovna "
         "in Prague. It connects directly to Sudek's own war experience and his "
         "lost arm. Here, solitude is personal \u2014 where his whole sensibility "
         "begins.", "must"),
        ("Remember, Sudek lost his own arm in that same war. When he photographs "
         "this man, he's also, in a way, photographing himself.", "opt"),
    ]),
    (15, "A Chair in the Magic Garden / Remembrance", "c. 1950s", "SOLITARY", [
        ("I'll end with this: a single empty chair in the garden. No people, just "
         "pure silence and memory. For me, this is the emotional heart of his "
         "entire body of work.", "must"),
        ("An empty chair always suggests someone who was there and is now gone. "
         "Let it sit in silence for a second before we move on.", "opt"),
    ]),
]


def _set_background(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _set_gradient_background(slide, top_color=PAPER, bottom_color=PAPER_DEEP):
    """Apply a soft, low-contrast vertical gradient so slides feel less flat."""
    fill = slide.background.fill
    fill.gradient()
    stops = fill.gradient_stops
    stops[0].position = 0.0
    stops[0].color.rgb = top_color
    stops[1].position = 1.0
    stops[1].color.rgb = bottom_color
    try:
        fill.gradient_angle = 90.0  # top -> bottom
    except (AttributeError, ValueError):
        pass


def _add_accent_rule(slide, top, left=Inches(1.0), width=Inches(2.4)):
    """Thin warm-brown rule used as a quiet decorative accent."""
    rule = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left, top, width, Pt(3))
    rule.fill.solid()
    rule.fill.fore_color.rgb = TAG
    rule.line.fill.background()
    rule.shadow.inherit = False
    return rule


def _add_textbox(slide, left, top, width, height, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    return tf


def _style_run(run, size, color, bold=False, italic=False):
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = FONT


def _set_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


# Colours used to mark speaker notes on the photo slides.
NOTE_MUST = INK            # core lines to always say (dark)
NOTE_OPT = RGBColor(0x9A, 0x6A, 0x2E)  # optional / time-filler lines (warm grey-brown)


def _set_photo_notes(slide, segments):
    """Render photo speaker notes, colour-coding must-say vs optional lines.

    `segments` is a list of (text, kind) where kind is "must" or "opt".
    A short legend explains the colour code so the presenter can tell at a
    glance which sentences are essential and which are time-fillers.
    """
    tf = slide.notes_slide.notes_text_frame
    tf.clear()

    legend = tf.paragraphs[0]
    lr = legend.add_run()
    lr.text = ("Legend \u2014 dark = must say   \u00b7   "
               "brown italic = optional (say only if you need to fill time)")
    lr.font.name = FONT
    lr.font.size = Pt(11)
    lr.font.bold = True
    lr.font.color.rgb = NOTE_OPT

    body = tf.add_paragraph()
    for text, kind in segments:
        run = body.add_run()
        if kind == "opt":
            run.text = "  (optional) " + text + " "
            run.font.color.rgb = NOTE_OPT
            run.font.italic = True
        else:
            run.text = text + " "
            run.font.color.rgb = NOTE_MUST
        run.font.name = FONT
        run.font.size = Pt(14)



def _blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _find_image(stem):
    """Return the path to images/<stem>.<ext> if one exists, else None."""
    for ext in IMAGE_EXTENSIONS:
        path = os.path.join(IMAGE_DIR, stem + ext)
        if os.path.isfile(path):
            return path
    return None


def _add_fitted_picture(slide, path, left, top, width, height):
    """Embed an image centered inside the box, preserving its aspect ratio.

    The picture is scaled to fit entirely within (width, height) ("contain"),
    so nothing is cropped, and then centered within the box.
    """
    # Insert at native size first to read the image's real proportions.
    pic = slide.shapes.add_picture(path, left, top)
    native_w, native_h = pic.width, pic.height
    scale = min(width / native_w, height / native_h)
    new_w = int(native_w * scale)
    new_h = int(native_h * scale)
    pic.width = new_w
    pic.height = new_h
    pic.left = int(left + (width - new_w) / 2)
    pic.top = int(top + (height - new_h) / 2)
    pic.shadow.inherit = False
    return pic


def _add_image_or_placeholder(slide, stem, left, top, width, height,
                              placeholder_text, placeholder_size=18):
    """Embed images/<stem>.* if present; otherwise draw a placeholder box."""
    path = _find_image(stem)
    if path is not None:
        return _add_fitted_picture(slide, path, left, top, width, height)

    ph = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    ph.fill.solid()
    ph.fill.fore_color.rgb = PLACEHOLDER
    ph.line.color.rgb = MUTED
    ph.line.width = Pt(1)
    ph.shadow.inherit = False
    ptf = ph.text_frame
    ptf.word_wrap = True
    ptf.vertical_anchor = MSO_ANCHOR.MIDDLE
    pr = ptf.paragraphs[0]
    pr.alignment = PP_ALIGN.CENTER
    run = pr.add_run()
    run.text = placeholder_text
    _style_run(run, placeholder_size, MUTED, italic=True)
    return ph


def add_cover(prs):
    slide = _blank(prs)
    _set_gradient_background(slide)
    tf = _add_textbox(slide, Inches(1), Inches(2.4), Inches(11.33), Inches(2.7),
                      anchor=MSO_ANCHOR.MIDDLE)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r1 = p.add_run()
    r1.text = "The World in One Hand"
    _style_run(r1, 40, INK, bold=True)

    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run()
    r2.text = "How Limitation Became Sudek's Style"
    _style_run(r2, 24, MUTED, italic=True)

    p3 = tf.add_paragraph()
    p3.alignment = PP_ALIGN.CENTER
    r3 = p3.add_run()
    r3.text = "Josef Sudek \u2014 the Poet of Prague"
    _style_run(r3, 16, TAG)

    # Centered accent rule under the title block
    _add_accent_rule(slide, Inches(5.3), left=Inches(5.67), width=Inches(2.0))

    _set_notes(slide,
               "Good morning, everyone \u2014 thank you all for being here. Before "
               "I start, let me ask you something: have you ever felt that having "
               "less actually made you more creative? Hold that thought. Today I "
               "want to introduce a photographer who saw the whole world through a "
               "single window \u2014 and who made every one of his images with only "
               "one hand. His name is Josef Sudek, often called 'the Poet of "
               "Prague.' My talk is The World in One Hand, and honestly, the whole "
               "thing comes down to one question I kept asking myself: how did a "
               "physical limitation become an artistic style? Let's find out "
               "together. (0:30)")


def add_portrait_slide(prs):
    slide = _blank(prs)
    _set_gradient_background(slide)

    # Portrait image placeholder on the left
    _add_image_or_placeholder(
        slide, "portrait",
        Inches(0.9), Inches(1.1), Inches(4.6), Inches(5.3),
        "[ Portrait of Josef Sudek \u2014 insert image here ]",
        placeholder_size=16)

    # Title + short intro on the right
    tf = _add_textbox(slide, Inches(6.0), Inches(1.4), Inches(6.4), Inches(1.4))
    r = tf.paragraphs[0].add_run()
    r.text = "Josef Sudek"
    _style_run(r, 36, INK, bold=True)
    psub = tf.add_paragraph()
    rsub = psub.add_run()
    rsub.text = "1896\u20131976 \u00b7 the Poet of Prague"
    _style_run(rsub, 18, TAG, italic=True)

    _add_accent_rule(slide, Inches(3.05), left=Inches(6.05), width=Inches(2.0))

    body = _add_textbox(slide, Inches(6.0), Inches(3.4), Inches(6.4), Inches(3.2))
    lines = [
        "Czech photographer who spent his life in Prague.",
        "Lost his right arm in the First World War.",
        "Worked one-handed with a large-format view camera.",
        "Turned that limitation into a slow, quiet, poetic style.",
    ]
    first = True
    for line in lines:
        p = body.paragraphs[0] if first else body.add_paragraph()
        first = False
        rl = p.add_run()
        rl.text = line
        _style_run(rl, 18, INK)
        p.space_after = Pt(10)

    _set_notes(slide,
               "So this is the man himself \u2014 take a good look at that face. "
               "Josef Sudek, born in 1896, who lived and worked almost his entire "
               "life in Prague. Now here is the detail that always stops people: "
               "he made every single one of his photographs with just one hand, "
               "hauling around a heavy large-format camera. Try to imagine doing "
               "that on a tripod, by yourself, for fifty years. Keep his face in "
               "mind, because we're about to see how he turned a wartime injury "
               "into one of the most poetic styles in all of photography.")


def add_text_slide(prs, title, body_lines, notes):
    slide = _blank(prs)
    _set_gradient_background(slide)
    tf = _add_textbox(slide, Inches(1), Inches(0.9), Inches(11.33), Inches(1.2))
    r = tf.paragraphs[0].add_run()
    r.text = title
    _style_run(r, 32, INK, bold=True)

    _add_accent_rule(slide, Inches(1.95))

    body = _add_textbox(slide, Inches(1), Inches(2.3), Inches(11.33), Inches(4.4))
    first = True
    for line in body_lines:
        p = body.paragraphs[0] if first else body.add_paragraph()
        first = False
        run = p.add_run()
        run.text = line
        _style_run(run, 20, INK)
        p.space_after = Pt(10)
    _set_notes(slide, notes)


def add_photo_slide(prs, number, title, year, tag, notes):
    slide = _blank(prs)
    _set_gradient_background(slide)

    # Large image (or labelled placeholder if the file isn't present yet)
    _add_image_or_placeholder(
        slide, f"photo{number:02d}",
        Inches(0.7), Inches(0.7), Inches(8.4), Inches(6.1),
        f"[ Photo {number} \u2014 insert image here ]",
        placeholder_size=18)

    # Right-hand caption column
    cap = _add_textbox(slide, Inches(9.4), Inches(1.0), Inches(3.4), Inches(5.5))
    ptag = cap.paragraphs[0]
    rtag = ptag.add_run()
    rtag.text = tag
    _style_run(rtag, 18, TAG, bold=True)

    pnum = cap.add_paragraph()
    pnum.space_before = Pt(14)
    rnum = pnum.add_run()
    rnum.text = f"Photo {number} of 15"
    _style_run(rnum, 12, MUTED)

    ptitle = cap.add_paragraph()
    ptitle.space_before = Pt(10)
    rtitle = ptitle.add_run()
    rtitle.text = title
    _style_run(rtitle, 22, INK, bold=True)

    pyear = cap.add_paragraph()
    ryear = pyear.add_run()
    ryear.text = year
    _style_run(ryear, 14, MUTED, italic=True)

    _set_photo_notes(slide, notes)


# Suggested holding institution / archive for each photograph. Sudek's work is
# still under copyright (he died in 1976), so these are reference sources for the
# reproductions, not public-domain releases. Verify the exact provenance of the
# file you used against the institution before presenting.
CREDITS = {
    1: "National Gallery of Canada / Nasjonalmuseet",
    2: "Sudek panorama series \u2014 various public collections",
    3: "Sudek panorama series \u2014 various public collections",
    4: "Sudek Project archive (sudekproject.cz)",
    5: "Museum of Modern Art (MoMA) / National Gallery of Canada",
    6: "National Gallery of Canada",
    7: "Sudek Project archive (sudekproject.cz)",
    8: "Sudek Project archive (sudekproject.cz)",
    9: "Museum of Modern Art (MoMA)",
    10: "Museum of Modern Art (MoMA)",
    11: "Museum of Modern Art (MoMA)",
    12: "National Gallery of Canada",
    13: "Sudek panorama series \u2014 various public collections",
    14: "Sudek Project archive (sudekproject.cz)",
    15: "Cleveland Museum of Art / National Gallery of Canada",
}


def add_credits_slide(prs):
    slide = _blank(prs)
    _set_gradient_background(slide)
    tf = _add_textbox(slide, Inches(1), Inches(0.6), Inches(11.33), Inches(0.9))
    r = tf.paragraphs[0].add_run()
    r.text = "Image Credits"
    _style_run(r, 30, INK, bold=True)

    _add_accent_rule(slide, Inches(1.5))

    note = _add_textbox(slide, Inches(1), Inches(1.6), Inches(11.33), Inches(0.6))
    rn = note.paragraphs[0].add_run()
    rn.text = ("All photographs \u00a9 Estate of Josef Sudek. Reproduced here for "
               "educational, non-commercial use only.")
    _style_run(rn, 13, MUTED, italic=True)

    body = _add_textbox(slide, Inches(1), Inches(2.25), Inches(11.33), Inches(4.6))
    first = True
    for number, title, year, _tag, _notes in PHOTOS:
        p = body.paragraphs[0] if first else body.add_paragraph()
        first = False
        run = p.add_run()
        source = CREDITS.get(number, "see images/sources.md")
        run.text = f"{number}. {title} ({year}) \u2014 {source}"
        _style_run(run, 11, INK)
        p.space_after = Pt(3)
    _set_notes(slide,
               "All images are reproductions of works by Josef Sudek, whose "
               "estate still holds copyright (he died in 1976), shown here for "
               "educational, non-commercial use. The listed institutions are the "
               "suggested reference sources; verify the exact provenance and the "
               "title and year of each file you used against the holding "
               "institution before presenting. Full search links are in "
               "images/sources.md.")


def build(output):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    # 1. Cover
    add_cover(prs)

    # 2. Portrait of the photographer
    add_portrait_slide(prs)

    # 3. Who was Josef Sudek?
    add_text_slide(
        prs, "Who Was Josef Sudek?",
        ["Born in Bohemia, 1896.",
         "Lost his right arm in the First World War.",
         "Chose the large-format view camera \u2014 slow, deliberate work on a tripod.",
         "Spent his life in Prague: 'the Poet of Prague.'"],
        "Josef Sudek was born in Bohemia in 1896. Then the First World War came, "
        "and he lost his right arm. Just think about that for a second \u2014 for "
        "most people that's the end of a photographic career; a camera is hard "
        "enough to handle with two hands. But here's what fascinates me about "
        "Sudek: instead of fighting the limitation, he leaned right into it. He "
        "gave up fast, handheld photography and committed to the large-format "
        "view camera, the kind of slow, deliberate instrument that forces you to "
        "slow down. And he spent almost his whole life right here in Prague \u2014 "
        "which is exactly why they call him the Poet of Prague. (1:30)")

    # 3. Main idea
    add_text_slide(
        prs, "The Main Idea",
        ["Sudek's limitation pushed him toward one way of working: slow, close, patient.",
         "That became his signature style.",
         "Four qualities organize this talk:",
         "SLOW  \u00b7  CLOSE  \u00b7  REPEATED  \u00b7  SOLITARY",
         "Every photograph is evidence for one of them."],
        "Okay, so here's the whole argument of my talk in a nutshell. Sudek's "
        "limitation quietly pushed him toward one particular way of working "
        "\u2014 slow, close, and patient \u2014 and over time that became his "
        "signature style. So instead of organizing his photographs by subject, "
        "the usual way, I've grouped them by four qualities that all grow out of "
        "that single limitation: Slow, Close, Repeated, and Solitary. Keep these "
        "four words in the back of your mind \u2014 every photograph you're about "
        "to see is really evidence for one of them. (1:00)")

    # 4..18 One photo per slide
    for number, title, year, tag, notes in PHOTOS:
        add_photo_slide(prs, number, title, year, tag, notes)

    # 19. Bringing it together
    add_text_slide(
        prs, "Bringing It Together",
        ["He lost his arm, so he worked slowly.",
         "He worked slowly, so he stayed close.",
         "He stayed close, so he repeated intimate subjects for years.",
         "The result: stillness and solitude.",
         "The constraint didn't shrink his world \u2014 it concentrated it."],
        "So let's circle back to the question I asked at the very beginning. How "
        "did a limitation become a style? Follow the chain with me: he lost his "
        "arm, so he worked slowly. He worked slowly, so he stayed close. He "
        "stayed close, so he ended up repeating the same intimate subjects for "
        "years. And out of all of that came a body of work defined by stillness "
        "and solitude. This is the part I find genuinely moving \u2014 the "
        "constraint didn't shrink his world, it concentrated it. He found the "
        "infinite inside one small studio, one window, one garden. (1:30)")

    # 20. Conclusion & Q&A
    add_text_slide(
        prs, "Conclusion & Q&A",
        ["Creativity isn't always about having more.",
         "Sometimes it's about going deeper into less.",
         "One hand, one window \u2014 some of the most poetic photographs of the 20th century.",
         "Thank you. Questions?"],
        "So if you take just one thing away today, let it be this. Josef Sudek "
        "shows us that creativity isn't really about having more \u2014 more "
        "mobility, more subjects, more expensive equipment. Sometimes it's about "
        "going deeper into less. With his single hand and his single window he "
        "gave us some of the most poetic photographs of the entire twentieth "
        "century. And maybe that's worth remembering the next time we feel "
        "limited by something ourselves. Thank you so much for listening \u2014 "
        "now I'd love to hear your questions. (1:00)")

    # Credits
    add_credits_slide(prs)

    prs.save(output)
    return len(prs.slides._sldIdLst)


def main():
    global IMAGE_DIR
    parser = argparse.ArgumentParser(description="Build the Josef Sudek PPTX.")
    parser.add_argument("--output", default=DEFAULT_OUTPUT,
                        help=f"Output .pptx path (default: {DEFAULT_OUTPUT})")
    parser.add_argument("--images-dir", default=DEFAULT_IMAGE_DIR,
                        help="Folder holding portrait.* and photo01..photo15.* "
                             f"(default: {DEFAULT_IMAGE_DIR})")
    args = parser.parse_args()
    IMAGE_DIR = args.images_dir

    embedded = sum(1 for stem in ["portrait"] +
                   [f"photo{n:02d}" for n in range(1, 16)]
                   if _find_image(stem) is not None)
    count = build(args.output)
    print(f"Wrote {args.output} with {count} slides "
          f"({embedded}/16 images embedded from '{IMAGE_DIR}/').")


if __name__ == "__main__":
    main()
