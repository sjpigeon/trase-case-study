# Trase × RxCo — Product Demo & Impact Analysis

A professional 18-slide PowerPoint presentation generated programmatically with
[python-pptx](https://python-pptx.readthedocs.io/).

The deck covers:
- RxCo workflow overview and business model
- Live product demo of the Eligibility Determination Agent (technical architecture, edge cases)
- Full quantitative impact analysis across all six enrollment-workflow steps
- 12-month ROI projection (3.1× — exceeds the 2× guarantee by 55 %)
- Assumptions, validation plan, and recommended next steps

---

## Files

| File | Description |
|------|-------------|
| `generate_deck.py` | Python script that builds the `.pptx` from scratch |
| `Trase_x_RxCo_Product_Demo_and_Impact_Analysis.pptx` | Pre-built deck (ready to download) |
| `requirements.txt` | Python dependencies |

---

## Regenerate the deck

```bash
pip install -r requirements.txt
python generate_deck.py
```

This overwrites `Trase_x_RxCo_Product_Demo_and_Impact_Analysis.pptx` in the current directory.

---

## Import into Google Slides

1. Open [Google Slides](https://slides.google.com) and create a new presentation.
2. Go to **File → Import Slides → Upload**.
3. Upload `Trase_x_RxCo_Product_Demo_and_Impact_Analysis.pptx`.
4. Select **All Slides** and click **Import Slides**.

All text remains fully editable after import.