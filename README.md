# BGN to EUR Excels Verifier

A Python tool designed to verify currency conversions in financial Excel reports. It compares a source Excel file (BGN) against a target Excel file (EUR) to ensure that the exchange rate calculation is accurate to the cent.

**Fixed Exchange Rate:** `1 EUR = 1.95583 BGN`

## ✨ Features

* **Dual Modes:** Run via Command Line (CLI) for automation or via Streamlit (Web UI) for visual interaction.
* **Precise Math:** Uses Python's `Decimal` class for high-precision financial rounding (Round Half Up).
* **Formula Handling:** Reads computed values (`data_only=True`) rather than Excel formulas.
* **Flexible verification:** Can verify specific columns or scan entire files.
* **Detailed Reporting:** Pinpoints the exact row and column where a mismatch occurs.

## 🛠️ Prerequisites

* Python 3.9+
* `pip` (Python package manager)

### Installation

1.  Clone or download this repository.
2.  Install the required dependencies:

```bash
pip install -r requirements.txt