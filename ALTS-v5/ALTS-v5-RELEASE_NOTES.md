# Release Notes - ALTS v5.0.0

### Highlights & New Features
* **5-Column Invoice Layout:** Updated HTML rendering engine to explicitly break out line-item discounts alongside Unit Price and Amount.
* **Inline Math Expressions:** Line amounts are calculated as `(QTY * Unit Price) - Discount` directly within the loop appender.
* **Null-Safe Discount Logic:** Implemented `if(equals(trim(string(item()?['Discount'])), ''), 0, float(item()?['Discount']))` to handle blank/null Excel cells without runtime flow crashes.

### File Structure
* `Send_Multi_Service_Invoices_V5.zip`: Complete Power Automate package export.
* `invoice_template.html`: 5-column HTML invoice template.
* `ALTS_VERSION3.xlsx`: Data source schema supporting the `Discount` column.