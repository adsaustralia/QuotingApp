# Quoting App — Simple UI (Streamlit)

Minimal horizontal UI:
- Size
- Stock
- Qty
- Sides (SS/DS)

Supports:
- Width+Height rectangles
- Circles (Diameter or 'DIA 600' in text)
- Adidas-style multiple qty columns (melt by selected start/end columns)

## Run
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Standard stocks CSV
`data/standard_stocks.csv` must contain:
- `stock_name_std`
- `sqm_rate`
