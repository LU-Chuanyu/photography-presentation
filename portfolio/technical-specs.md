# Shared Technical & Export Specs + Submission Checklist

These apply to whichever track you choose. They come directly from the assignment
brief — match them exactly, because malformed files are an easy way to lose marks.

## File specifications (from the brief)

| Spec | Requirement | Notes |
| --- | --- | --- |
| **Format** | JPG | Export as `.jpg`. |
| **Resolution** | 72 pixels per inch (PPI) | Set on export. |
| **Long-edge dimension** | **1200–1300 px** | Aim for ~**1250 px** on the long edge; let the short edge follow the aspect ratio. |
| **Count** | **10** photographs | Exactly ten; order is graded. |
| **Naming** | Sequential | e.g. `01.jpg … 10.jpg`. The sequence carries significant weight. |
| **Delivery** | Single ZIP or RAR | Package all 10 JPGs into one archive and attach that one file to the submission email. |
| **Artist Statement** | One page | Submit alongside the images (PDF recommended). |

## Suggested export workflow

1. **Edit at full resolution** — do all tonal/color/crop work on the originals.
2. **Finalize the 10 images and their order** before exporting anything.
3. **Batch-resize** the long edge to **1250 px** at **72 PPI**.
4. **Export JPG** at quality ~80–90 (high quality, reasonable file size).
5. **Rename** sequentially `01`–`10` in the intended viewing order.
6. **Zip** all ten into one archive.
7. **Final check:** confirm no file's long edge falls below 1200 px or above 1300 px,
   that there are exactly 10 files, and that the order is correct.

### Example: batch-resize with ImageMagick

> Reference only — use your normal editor (Lightroom/Capture One/Photoshop "Export As")
> if you prefer. Run on a *copy* of your finished, ordered, full-res exports.

```bash
# Resize the long edge to 1250 px and set 72 PPI, writing JPGs to ./export
mkdir -p export
i=1
for f in $(ls *.tif *.png *.jpg 2>/dev/null | sort); do
  out=$(printf "export/%02d.jpg" "$i")
  magick "$f" -resize 1250x1250\> -units PixelsPerInch -density 72 -quality 88 "$out"
  i=$((i+1))
done

# Verify every long edge is within 1200–1300 px
for f in export/*.jpg; do
  identify -format "%f %wx%h\n" "$f"
done

# Package for submission
cd export && zip ../prague-portfolio.zip *.jpg && cd ..
```

The `1250x1250\>` geometry resizes so the **longer** edge becomes 1250 px while
preserving aspect ratio (the `\>` only shrinks, never enlarges) — so confirm your
full-res exports are larger than 1250 px on the long edge first.

## Submission checklist

- [ ] Track chosen and thesis sentence locked.
- [ ] (Track 3 only) Instructor emailed and approval received **before** shooting.
- [ ] Exactly **10** images selected.
- [ ] Sequence/order deliberately authored.
- [ ] All files **JPG**, **72 PPI**, long edge **1200–1300 px**.
- [ ] Files named `01.jpg … 10.jpg` in viewing order.
- [ ] One-page **Artist Statement** written (≈250–350 words), metaphorical not literal,
      justifying color vs. B&W, crops, time of day, and sequence.
- [ ] All 10 JPGs packaged into a **single ZIP/RAR**.
- [ ] Submission email drafted with the archive + statement attached.
