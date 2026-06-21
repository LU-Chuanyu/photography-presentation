# images/ — drop your photos here

`build_pptx.py` looks in this folder and automatically embeds any matching
file into the deck. If a file is missing, that slide simply falls back to a
labelled placeholder box, so the script always runs.

## File names (this is the important part)

Use these exact base names — the extension can be `.jpg`, `.jpeg`, `.png`,
`.gif`, `.bmp`, `.tif`, `.tiff`, or `.webp`:

| File name      | Where it goes                                   |
|----------------|-------------------------------------------------|
| `portrait.jpg` | Portrait slide near the start (Josef Sudek)     |
| `photo01.jpg`  | Photo 1 — Saint Vitus Cathedral, interior       |
| `photo02.jpg`  | Photo 2 — Prague Panorama (misty city)          |
| `photo03.jpg`  | Photo 3 — Prague at Night                       |
| `photo04.jpg`  | Photo 4 — Mionší Forest, mist in the woods      |
| `photo05.jpg`  | Photo 5 — The Window of My Studio               |
| `photo06.jpg`  | Photo 6 — The Last Rose                         |
| `photo07.jpg`  | Photo 7 — Glass and Egg (Labyrinths)            |
| `photo08.jpg`  | Photo 8 — Still life with Bread and Glass       |
| `photo09.jpg`  | Photo 9 — Window, view to the garden (summer)   |
| `photo10.jpg`  | Photo 10 — Window, frost (winter)               |
| `photo11.jpg`  | Photo 11 — A Walk in the Magic Garden I         |
| `photo12.jpg`  | Photo 12 — A Walk in the Magic Garden II        |
| `photo13.jpg`  | Photo 13 — Prague Panorama, empty street/square |
| `photo14.jpg`  | Photo 14 — Veteran from the Invalidovna         |
| `photo15.jpg`  | Photo 15 — A Chair in the Magic Garden          |

Note the two-digit numbering: `photo01`, not `photo1`.

## Where to find the photos

Sudek's work is still under copyright, so museums show low-res previews only.
See [`sources.md`](sources.md) for per-image museum/archive links and search
keywords (Sudek Project, National Gallery of Canada, MoMA, Wikimedia Commons,
Internet Archive).

## Rebuild

```bash
python build_pptx.py
```

Images are scaled to fit each slot without cropping and centered, so any
shape of photo (portrait or landscape) works. The terminal prints how many
of the 16 images were embedded.
