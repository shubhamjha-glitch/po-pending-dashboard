from pathlib import Path

requirements = streamlit
pandas
openpyxl
numpy


readme = # PO Pending Dashboard

A Streamlit web application that automates the creation and analysis of PO Pending data.

## Features

- Upload the daily PO Pending Excel file.
- Upload the daily Schedule Date Excel file.
- Use a fixed Barcode Status master file.
- Automatically set `DELIVERY DATE = SCHEDULE DATE`.
- Calculate `PKT QTY = PENDING QTY  OUTER PACK`.
- Create a consolidated PO Pending dataset.
- Calculate Order, Received, Cancelled and Pending values.
- Calculate Expire Days and Delay Groups.
- Match barcode status from the master file.
- Display an interactive dashboard with filters.
- Download the final PO Pending Excel file.
- Download pivot-style summary tables.

## Project Structure

```text
PO_PENDING_DASHBOARD
├── app.py
├── BARCODE STATUS.xlsx
├── requirements.txt
└── README.md