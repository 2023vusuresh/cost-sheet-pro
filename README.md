# Cost Sheet Pro — Final Version

Professional Streamlit cost-sheet generator with:
- Individual cost-component selection
- Comprehensive standard cost library
- Custom cost component entry
- Add/remove/clear cost lines
- Live standard cost-sheet calculations
- Unit economics
- Category and component analysis
- Professional Excel workbook export

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Costing flow
Prime Cost → Works/Factory Cost → Cost of Production → Cost of Goods Sold → Cost of Sales → Profit → Sales Value → Sales Value.


## Tax treatment
GST/Tax is intentionally excluded from the standard cost-sheet output because it is a statutory billing/tax item rather than a product-cost component.

## Cost Component Input Method
For every cost component, enter **Total Qty / Units** and **Total Amount**.
The application automatically calculates **Rate / Unit = Total Amount ÷ Total Qty / Units**.
The calculated rate is read-only and is used consistently in the cost calculations and Excel export.
