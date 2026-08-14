# Power Automate Invoice Generator - Version 2.0.0 (`Testt Copy1`)

## 📌 Overview
Version 2 enhances the automated invoicing workflow by generating and dispatching invoices populated with single-service line items per client based on renewal criteria.

---

## 🛠️ Architecture & Flow Structure

### 1. Trigger & Data Retrieval
* **Recurrence:** Automates schedule execution (daily/weekly).
* **List rows present in a table (Excel):** Reads client data, billing status, renewal dates, and service details.

### 2. Client Iteration & Filtering
* **For each (Outer Loop):** Processes rows returned from the client table.
* **Condition:** Checks the determinant flag (`EmailDeterminant` / `Alert` condition) to confirm if an invoice needs to be generated.

### 3. Invoice Metadata & Header Updates (True Branch)
* **Get a row:** Fetches current invoice sequence metadata from the control table.
* **Compose - Calculate Next Invoice:** Calculates the next incremental invoice number.
* **Compose - Format Invoice String:** Formats the invoice code (e.g., `ITS_2026_1049`).
* **Update a row (Header):** Writes updated invoice metadata (Invoice Number, Date, Address, Client Name) into the target template sheet.

### 4. Line Item Population (Single Service)
* **Update a row 1 (Line Items):** Populates row `1` of the line items table directly with:
  * `DESCRIPTION` (`Service`)
  * `QTY`
  * `UNIT PRICE` (`Cost`)

### 5. Conversion, Export & Delivery
* **Run script:** Executes Office Script in Excel to recalculate formulas and prep the document layout.
* **Delay:** Waits (30 seconds) for OneDrive file synchronization.
* **Convert file:** Converts the target Excel template into a PDF document.
* **Create file:** Saves the temporary PDF file into OneDrive storage.
* **Send an email (V2):** Dispatches the generated PDF invoice as an attachment to the client's email address.

---

## 📝 Key Improvements over Version 1
* Direct Excel table row updates for single-service line items (`Update a row 1`).
* Full end-to-end PDF conversion and Office 365 Outlook email integration.