# photography-presentation

Materials for a photography course presentation on **Josef Sudek** —
*The World in One Hand: How Limitation Became Sudek's Style* (Option A theme).

## Contents

- `presentation-plan.md`: theme, the 15 selected photographs, one-photo-per-slide structure, timing, and PPT checklist
- `presentation-script.md`: English speaker script (one section per photo slide)
- `build_pptx.py`: generator that builds the slide deck (one photo per slide, with speaker notes)
- `Josef-Sudek-Presentation.pptx`: generated PowerPoint deck (a Sudek portrait slide near the start, one-photo-per-slide content, soft gradient backgrounds, and an image-credits slide)

## Rebuild the deck

```bash
pip install python-pptx
python build_pptx.py            # writes Josef-Sudek-Presentation.pptx
python build_pptx.py --output my-deck.pptx
```

## Adding the real photographs (automatic)

Drop the images into the `images/` folder using these exact base names and
re-run `python build_pptx.py`:

- `portrait.jpg` — the photographer portrait near the start
- `photo01.jpg` … `photo15.jpg` — the 15 photographs (note the two-digit numbering)

Accepted extensions: `.jpg`, `.jpeg`, `.png`, `.gif`, `.bmp`, `.tif`, `.tiff`,
`.webp`. Each image is scaled to fit its slot without cropping and centered,
so portrait or landscape photos both work. Any image that is missing falls
back to a labelled placeholder box, and the script prints how many of the 16
images were embedded. See `images/README.md` for the full name-to-slide map.
Use a different folder with `python build_pptx.py --images-dir path/to/folder`.

Use public-domain or properly licensed reproductions, and complete the
image-credits slide before presenting.
