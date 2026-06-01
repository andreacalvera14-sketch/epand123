# Epsilon CSA Deck

This repository auto-generates a **fully editable PowerPoint (.pptx)** presentation for the Epsilon "Client Success Associate" interview task using [`python-pptx`](https://python-pptx.readthedocs.io/). Every push to `main` triggers a GitHub Actions workflow that builds the deck and makes it available to download — no local tools required.

Every text box, table, and bullet in the output is a native PowerPoint element — fully clickable and editable in PowerPoint, Keynote, and Google Slides.

> **Note:** All figures in the deck are illustrative.

---

## Download

**Direct link:**
```
https://github.com/andreacalvera14-sketch/epand123/releases/latest/download/Epsilon_CSA_Deck.pptx
```

Or:
1. Go to the [**Releases page**](https://github.com/andreacalvera14-sketch/epand123/releases/latest).
2. Under **Assets**, click **`Epsilon_CSA_Deck.pptx`** to download.

---

## How to edit the deck

1. Edit **`build_deck.py`** to update content, layout, or styling  
   (or use `epsilon-deck.md` as a content reference — it is kept in the repo but not used in the build).
2. Commit and push to `main`.
3. The workflow runs automatically and a new `.pptx` is ready in ~30 seconds.
4. Download the updated file from the Releases page or the Actions artifacts.

---

## How it works

| File | Purpose |
|---|---|
| `build_deck.py` | Python script (python-pptx) — generates the fully editable `.pptx` |
| `epsilon-deck.md` | Content reference in Marp markdown — not used in the build |
| `.github/workflows/build-deck.yml` | GitHub Actions workflow — installs python-pptx, runs the script, publishes `.pptx` to the `latest-deck` Release |

The workflow installs `python-pptx`, runs `build_deck.py`, uploads the result as a workflow artifact (retained 30 days), and attaches it to the `latest-deck` GitHub Release for a permanent one-click download.

---

## Brand colours

The deck is styled in Epsilon's black / white / turquoise brand identity (Publicis Groupe):

| Colour | Hex | Usage |
|---|---|---|
| Epsilon Black | `#0B0B0B` | Headings, body text |
| White | `#FFFFFF` | Slide background |
| Publicis Turquoise | `#00B0A3` | Accent rules, callout bars (sparingly) |
| Secondary Grey | `#6B6B6B` | Captions, footers, secondary text |
| Light Grey | `#E7E7E7` | Table borders |
