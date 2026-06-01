# Epsilon CSA Deck

This repository auto-generates a finished **PowerPoint (.pptx)** presentation for the Epsilon "Client Success Associate" interview task using [Marp](https://marp.app/). Every push to `main` triggers a GitHub Actions workflow that builds the deck and makes it available to download — no local tools required.

> **Note:** All performance figures quoted in the deck are illustrative.

---

## ⬇️ How to download the finished .pptx

### Easiest — GitHub Releases (stable URL)

1. Go to the [**Releases page**](https://github.com/andreacalvera14-sketch/epand123/releases/latest).
2. Under **Assets**, click **`Epsilon_CSA_Deck.pptx`** to download.

Direct link (works after the first workflow run on `main`):
```
https://github.com/andreacalvera14-sketch/epand123/releases/latest/download/Epsilon_CSA_Deck.pptx
```

### Alternative — Actions artifacts

1. Go to the [**Actions tab**](https://github.com/andreacalvera14-sketch/epand123/actions).
2. Click the latest **"Build Epsilon Deck"** workflow run.
3. Scroll to the **Artifacts** section at the bottom.
4. Download **`Epsilon_CSA_Deck.zip`** (contains the `.pptx`).

---

## ✏️ How to edit the deck

1. Edit **`epsilon-deck.md`** in this repository (the GitHub web editor works fine).
2. Commit and push to `main`.
3. The workflow runs automatically and a new `.pptx` is generated within a minute or two.
4. Download the updated file from the Releases page or the Actions artifacts.

---

## 🔧 How it works

| File | Purpose |
|---|---|
| `epsilon-deck.md` | Marp markdown source — all slide content lives here |
| `.github/workflows/build-deck.yml` | GitHub Actions workflow — builds `.pptx` on every push to `main` |

The workflow uses the official [`marpteam/marp-cli`](https://github.com/marp-team/marp-cli) Docker image to convert the Marp markdown into a `.pptx`, uploads it as a workflow artifact (retained 30 days), and also attaches it to the `latest-deck` GitHub Release for a permanent download link.
