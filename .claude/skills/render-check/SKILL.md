---
name: render-check
description: Deterministically renders every slide layout, visual format, and background mode this repo supports, without calling the AI, and flags any real text box that overflows the slide bounds. Use after any change to pptx_builder.py, models.py, or theme.py to catch geometry regressions fast, before spending time or API quota on a live AI generation.
---

Run `python scripts/render_check.py` from the repo root.

It builds one synthetic `PresentationData` covering every `type`/`layout` combination
this repo renders (including the chart-insufficient-data and
comparison-insufficient-data downgrade paths), then calls
`pptx_builder.build_presentation` three times — no custom background, a light custom
background, and a dark one — and once more per `visual_format` to check the
bullet/body length clamps.

It prints one line per build (`OK ...` or a traceback) and, for the first build, a
bounding-box scan listing any shape with non-empty text whose box falls outside
`[0, slide_width] x [0, slide_height]`. Decorative bleed shapes (e.g. the title
slide's off-canvas circles) are excluded automatically since they carry no text — an
overflowing *empty* shape is intentional design, not a bug.

This does not call OpenRouter or Pexels — image-backed layouts render without an
image if none is supplied. It exists specifically so rendering/geometry regressions
can be caught without spending API quota or waiting on a slow free-tier model. Pair
it with the `prompt-tester` agent when the change is to `ai_service.py`'s prompt text
rather than to the renderer — this skill cannot catch a prompt regression, since it
never talks to the model.

Generated `.pptx` files land in `output/` (gitignored) for manual visual inspection in
PowerPoint/LibreOffice when the automated bounding-box check isn't enough to judge
whether something actually looks right.
