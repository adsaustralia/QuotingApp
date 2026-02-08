# Quoting App — Adidas MVP (Streamlit)

This is a starter internal web app to quote from the Adidas pricing sheet format.

## What it does (MVP)
- Upload Adidas Excel
- Pick header row + sheet
- Pick columns (Size, Stock, Qty mode)
- Melt Adidas "variant qty columns" into line items
- Parse the variant header (newline-separated spec)
- Calculate SQM (rectangle/circle/odd)
- Map customer stock names to your internal standard stocks (sqm rate)
- Export Quote as **Excel** and **PDF**

## Run
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Files
- `app.py` Streamlit app
- `data/standard_stocks.csv` sample internal stock rates
- `data/mappings/*_stock_map.json` saved customer stock mappings
