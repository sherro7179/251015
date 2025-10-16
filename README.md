# Workspace Overview

This repository currently contains two independent front-end experiences:

1. **World Pulse Timer** – a redesigned world clock dashboard with language/time zone switching and impact animations.
2. **Car Picker Quiz** – a Streamlit application that quizzes players on car make/model/year using the picture dataset generated from the _The Car Connection_ scraper.

Both projects share the same root workspace but live in different folders. The sections below describe how to run each app and how the code is organised.

---

## World Pulse Timer

- Open `timer/index.html` in any modern browser (Chrome, Edge, Safari are recommended).
- The dashboard highlights the selected time zone (default: local) with AM/PM formatting, millisecond precision, and punchy number transitions when hours/minutes/seconds change.
- KO/EN buttons swap UI language instantly.
- The world map background and city cards show six major cities with day/night status relative to the current moment.

### File Map
- `timer/index.html` – structural HTML, controls, and semantic regions.
- `timer/style.css` – glassmorphism-inspired visuals, animated transitions, and world map overlay.
- `timer/script.js` – language & timezone switching logic, animation triggers, city card updates.
- `timer/assets/world-map.svg` – light-weight SVG backdrop.

---

## Car Picker Quiz (Streamlit)

### Prerequisites
- Python 3.9+
- Install dependencies from `car_picker/requirements.txt`
  ```bash
  pip install -r car_picker/requirements.txt
  ```
- Place the scraped images inside `car_picker/data/`. The quiz expects the original naming scheme from the [predicting-car-price-from-scraped-data](https://github.com/nicolas-gervais/predicting-car-price-from-scraped-data/tree/master/picture-scraper) project (`Make_Model_Year_..._XYZ.jpg`).

### Running the App
```bash
streamlit run car_picker/app.py
```

### Feature Highlights
- **64k image index**: the first launch scans `car_picker/data/`, builds a duplicate-free metadata cache (`car_picker/metadata/index.json`), and reuses it on subsequent runs. You can rebuild the cache from the sidebar if the dataset changes.
- **Difficulty modes**:
  - `하 (easy)`: guess the manufacturer only.
  - `중 (medium)`: guess manufacturer + model.
  - `상 (hard)` *(default)*: guess manufacturer + model + year with 10-way multiple choice.
- **Session design**: 20-question default (5–40 configurable), no partial credit, no repeated cars inside a session, and real-time score feedback.
- **Cards-first UI**: every question shows a thumbnail (512px max, generated on demand under `car_picker/static/thumbs/`) with two-column answer cards.
- **History log**: completed sessions are appended to `car_picker/_state/sessions.json` so multiple users on the same host can track recent scores.

### Code Structure
- `car_picker/app.py` – Streamlit UI, sidebar controls, session management, and feedback flow.
- `car_picker/config.py` – central paths and constants (data directory, cache paths, defaults).
- `car_picker/quiz/parser.py` – filename parser and index builder that deduplicates on metadata key.
- `car_picker/quiz/engine.py` – quiz logic, distractor generation, and 10-option assembly per difficulty.
- `car_picker/quiz/utils.py` – helpers for thumbnail creation, deterministic shuffling, and atomic writes.
- `car_picker/quiz/store.py` – lightweight session history persistence.
- `car_picker/requirements.txt` – minimal dependency list (`streamlit`, `Pillow`).

### Workflow Tips
- The first index build can take a few minutes for all 64,467 images. The spinner stays active during the process.
- Thumbnails are generated lazily when an image is first displayed; keep `car_picker/static/thumbs/` writable.
- Use the sidebar buttons to restart a session or rebuild metadata when you add/remove images.

---

## Repository Status

- Timer redesign committed as `[WIP] Redesign timer dashboard`.
- Car Picker Quiz implementation is currently in-progress and should be reviewed/tested before tagging a release.
