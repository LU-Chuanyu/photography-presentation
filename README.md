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

Each photo slide contains a placeholder box. Replace the placeholders with
public-domain or properly licensed reproductions, and complete the image-credits
slide before presenting.
