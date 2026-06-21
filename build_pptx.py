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

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

# 16:9 canvas
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

FONT = "Calibri"

# Quiet, low-saturation palette echoing Sudek's mood
INK = RGBColor(0x20, 0x20, 0x20)        # near-black text
MUTED = RGBColor(0x6E, 0x6A, 0x63)      # warm grey
PAPER = RGBColor(0xF4, 0xF2, 0xED)      # warm off-white background
PAPER_DEEP = RGBColor(0xE5, 0xDF, 0xD2)  # slightly deeper warm tone for gradient
PLACEHOLDER = RGBColor(0xDF, 0xDB, 0xD2)  # image placeholder fill
TAG = RGBColor(0x8A, 0x6D, 0x3B)        # warm brown accent for trait tags

DEFAULT_OUTPUT = "Josef-Sudek-Presentation.pptx"

# Each photo slide: (number, title, year, trait tag, speaker notes)
PHOTOS = [
    (1, "Saint Vitus Cathedral, interior", "c. 1924–1928", "SLOW",
     "First, SLOW. The view camera demanded long exposures, so Sudek learned to "
     "wait. A long exposure gathers soft, diffused light the eye barely notices "
     "\u2014 light becomes almost solid. This patience is the foundation of "
     "everything that follows."),
    (2, "Prague Panorama, misty city", "c. 1950s–1960s", "SLOW",
     "Here, slowness turns fog and damp air into something dreamlike. This is "
     "from his panoramic work \u2014 a wide, demanding format. The mist isn't a "
     "problem to fix; it is the subject."),
    (3, "Prague at Night", "c. 1950s", "SLOW",
     "This near-darkness was only possible with a long exposure on a tripod. The "
     "empty, glowing street shows how Sudek used time itself as a tool \u2014 "
     "letting faint light slowly build into a quiet scene."),
    (4, "Mionší Forest, mist in the woods", "c. 1950s", "SLOW",
     "In a misty forest, the stillness becomes visible. The soft grey light and "
     "motionless trees are the trace of a long, patient exposure. Slowness was "
     "never a weakness \u2014 it was how he found poetry in ordinary light."),
    (5, "The Window of My Studio", "c. 1940–1948", "CLOSE",
     "Second, CLOSE. Because movement was difficult, Sudek worked within arm's "
     "reach. His most famous motif: the window of his studio, covered in mist and "
     "droplets. A single pane of glass became an entire universe."),
    (6, "The Last Rose", "c. 1956", "CLOSE",
     "On his table, a single fading rose. Seen this closely, an ordinary flower "
     "becomes a meditation on beauty and time \u2014 one of his most poetic still "
     "lifes."),
    (7, "Glass and Egg (Labyrinths)", "c. 1950s", "CLOSE",
     "In his still lifes, simple objects hold tiny, glowing worlds of light. A "
     "glass, an egg \u2014 close up, their surfaces become mysterious. Closeness "
     "was his limitation, but also his way of seeing the extraordinary in the "
     "plain."),
    (8, "Still life with Bread and Glass", "c. 1950s", "CLOSE",
     "Bread and a glass on a dark table. Humble, everyday things \u2014 yet lit "
     "so gently they feel almost sacred. Intimacy as a method: the closer he "
     "looked, the more meaning he found."),
    (9, "Window \u2014 view to the garden, summer", "c. 1940s", "REPEATED",
     "Third, REPEATED. Sudek returned to the same subjects for years. Here is the "
     "view from his studio window onto the garden in one season, full and green."),
    (10, "Window \u2014 frost, winter", "c. 1940s–1950s", "REPEATED",
     "The same window in winter, the glass laced with frost. Same frame, "
     "completely different mood. Repetition let him discover endless variation "
     "inside a single point of view."),
    (11, "A Walk in the Magic Garden I", "c. 1954", "REPEATED",
     "He gave the same devotion to his small, wild garden \u2014 his 'magic "
     "garden.' He photographed it again and again, finding life in its overgrown "
     "corners."),
    (12, "A Walk in the Magic Garden II", "c. 1960s", "REPEATED",
     "The garden again, at a different hour and light. Repetition was not a lack "
     "of ideas \u2014 it was depth. By doing less, again and again, he saw more "
     "than photographers who do everything once."),
    (13, "Prague Panorama \u2014 empty street/square", "c. 1950s–1960s", "SOLITARY",
     "Fourth, SOLITARY. Sudek's images are almost always empty of people. This "
     "empty square feels quiet, inward, a little melancholy \u2014 a whole city "
     "holding its breath."),
    (14, "Veteran from the Invalidovna", "c. 1922–1927", "SOLITARY",
     "This earlier, documentary work shows a war veteran at the Invalidovna in "
     "Prague. It connects directly to Sudek's own war experience and his lost "
     "arm. Here, solitude is personal \u2014 where his whole sensibility begins."),
    (15, "A Chair in the Magic Garden / Remembrance", "c. 1950s", "SOLITARY",
     "I'll end with this: a single empty chair in the garden. No people, just "
     "pure silence and memory. For me, this is the emotional heart of his entire "
     "body of work."),
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


def _blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


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
               "Good morning, everyone. Today I want to introduce a photographer "
               "who saw the whole world through a single window \u2014 and who "
               "made every one of his images with only one hand. His name is "
               "Josef Sudek, often called 'the Poet of Prague.' My talk is The "
               "World in One Hand, and my single question is: how did a physical "
               "limitation become an artistic style? (0:30)")


def add_portrait_slide(prs):
    slide = _blank(prs)
    _set_gradient_background(slide)

    # Portrait image placeholder on the left
    ph = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.9), Inches(1.1), Inches(4.6), Inches(5.3))
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
    run.text = "[ Portrait of Josef Sudek \u2014 insert image here ]"
    _style_run(run, 16, MUTED, italic=True)

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
               "This is the man himself: Josef Sudek, born in 1896, who lived "
               "and worked almost his entire life in Prague. He made every one "
               "of his photographs with a single hand, using a heavy "
               "large-format camera. Keep his face in mind as we look at how he "
               "turned a wartime injury into one of the most poetic styles in "
               "photography.")


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

    # Large image placeholder
    ph = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.7), Inches(0.7), Inches(8.4), Inches(6.1))
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
    run.text = f"[ Photo {number} \u2014 insert image here ]"
    _style_run(run, 18, MUTED, italic=True)

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

    _set_notes(slide, notes)


def add_credits_slide(prs):
    slide = _blank(prs)
    _set_gradient_background(slide)
    tf = _add_textbox(slide, Inches(1), Inches(0.9), Inches(11.33), Inches(1.0))
    r = tf.paragraphs[0].add_run()
    r.text = "Image Credits"
    _style_run(r, 30, INK, bold=True)

    _add_accent_rule(slide, Inches(1.85))

    body = _add_textbox(slide, Inches(1), Inches(2.1), Inches(11.33), Inches(4.6))
    first = True
    for number, title, year, _tag, _notes in PHOTOS:
        p = body.paragraphs[0] if first else body.add_paragraph()
        first = False
        run = p.add_run()
        run.text = f"{number}. {title} ({year}) \u2014 source / license: ____"
        _style_run(run, 12, INK)
        p.space_after = Pt(4)
    _set_notes(slide,
               "List the source and licensing for each image here. Use "
               "public-domain or properly licensed reproductions only, and verify "
               "each title and year against the holding institution before "
               "presenting.")


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
        "Josef Sudek was born in Bohemia in 1896. In the First World War he lost "
        "his right arm. For most people that would end a photographic career. But "
        "Sudek leaned into the limitation: he gave up fast, handheld photography "
        "and committed to the large-format view camera, which forces you to slow "
        "down. He spent almost his whole life in Prague. (1:30)")

    # 3. Main idea
    add_text_slide(
        prs, "The Main Idea",
        ["Sudek's limitation pushed him toward one way of working: slow, close, patient.",
         "That became his signature style.",
         "Four qualities organize this talk:",
         "SLOW  \u00b7  CLOSE  \u00b7  REPEATED  \u00b7  SOLITARY",
         "Every photograph is evidence for one of them."],
        "Here is the argument of this talk. Sudek's limitation pushed him toward "
        "one way of working \u2014 slow, close, and patient \u2014 and that became "
        "his signature style. I've organized his photographs not by subject, but "
        "by four qualities that grow from that single limitation: Slow, Close, "
        "Repeated, and Solitary. (1:00)")

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
        "So how did a limitation become a style? He lost his arm, so he worked "
        "slowly. He worked slowly, so he stayed close. He stayed close, so he "
        "repeated the same intimate subjects for years. Out of that came a body "
        "of work defined by stillness and solitude. The constraint didn't shrink "
        "his world \u2014 it concentrated it. (1:30)")

    # 20. Conclusion & Q&A
    add_text_slide(
        prs, "Conclusion & Q&A",
        ["Creativity isn't always about having more.",
         "Sometimes it's about going deeper into less.",
         "One hand, one window \u2014 some of the most poetic photographs of the 20th century.",
         "Thank you. Questions?"],
        "Josef Sudek shows us that creativity isn't about having more \u2014 more "
        "mobility, more subjects, more equipment. Sometimes it's about going "
        "deeper into less. His single hand, his single window, gave us some of "
        "the most poetic photographs of the twentieth century. Thank you for "
        "listening \u2014 I'd love to hear your questions. (1:00)")

    # Credits
    add_credits_slide(prs)

    prs.save(output)
    return len(prs.slides._sldIdLst)


def main():
    parser = argparse.ArgumentParser(description="Build the Josef Sudek PPTX.")
    parser.add_argument("--output", default=DEFAULT_OUTPUT,
                        help=f"Output .pptx path (default: {DEFAULT_OUTPUT})")
    args = parser.parse_args()
    count = build(args.output)
    print(f"Wrote {args.output} with {count} slides.")


if __name__ == "__main__":
    main()
