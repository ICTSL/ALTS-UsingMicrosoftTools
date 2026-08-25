# ALTS-UsingMicrosoftTools

![Microsoft Power Automate](https://img.shields.io/badge/Power%20Automate-0066B8?style=for-the-badge&logo=microsoftpowerautomate&logoColor=white)
![Microsoft Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![TypeScript](https://img.shields.io/badge/Office%20Scripts-3178C6?style=for-the-badge&logo=typescript&logoColor=white)
![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=for-the-badge&logo=html5&logoColor=white)

An end-to-end cloud solution built with **Microsoft Power Automate**, **Excel Online**, **Office Scripts**, and **HTML dynamic templates**[cite: 1]. 

The **Automated Lifecycle Tracking System (ALTS)** monitors client service lifecycle dates, evaluates renewal thresholds, dynamically populates and increments custom invoice templates, converts generated bills to PDF, archives sent copies to OneDrive, and dispatches automated billing emails[cite: 1].

---

## 🌟 What's New in Version 4.0 (`ALTS-v4`)[cite: 1]

Version 4 introduces significant improvements for multi-service itemization, HTML template rendering, and robust sequential billing[cite: 1]:

- **Dynamic Sequential Bill Incrementing:** Resolved recurring `Bill_No` resets across execution dates by fetching the global maximum bill number from Excel using `max(body('Select'))` before loop execution[cite: 1].
- **HTML & PDF Dynamic Rendering:** Replaced Excel template population with a standalone UTF-8 HTML template (`invoice_template.html`) converted directly to PDF via OneDrive[cite: 1].
- **Naira Currency & UTF-8 Encoding:** Added explicit `<meta charset="UTF-8">` headers and HTML entities (`&#8358;`) to guarantee consistent rendering of the Nigerian Naira (`₦`) symbol[cite: 1].
- **Excel Serial Date Conversion:** Converted raw Excel serial integer dates (e.g., `46254`) into standard formatted string dates (`dd-MMM-yyyy`) using expression mapping[cite: 1].
- **File & Variable Synchronization:** Synchronized dynamic bill numbers across HTML invoice headers, PDF file generation, dynamic email attachments, and Excel tracking updates (`Update a row 1`)[cite: 1].

---

## 📺 Video Tutorial & Demonstrations[cite: 1]

### Video Guide[cite: 1]
Watch the complete step-by-step video guide on YouTube[cite: 1]:

[![Watch the Tutorial](https://img.youtube.com/vi/mvzyZDertSg/maxresdefault.jpg)](https://www.youtube.com/watch?v=mvzyZDertSg)[cite: 1]

### Workflow Demonstrations[cite: 1]
https://github.com/user-attachments/assets/e76e6fdd-48cf-4403-a769-9fecae7b571c[cite: 1]

https://github.com/user-attachments/assets/ac277f27-49bd-441a-a12b-30a5e0cab44f[cite: 1]

---

## 🏗️ Repository Directory Structure[cite: 1]

```text
ALTS-UsingMicrosoftTools/
├── ALTS-v4/
│   ├── Send_Multi-Service_Invoices_v4.zip   # Version 4 Exported Power Automate Flow
│   ├── invoice_template.html               # UTF-8 HTML Invoice Template
│   ├── ALTS_VERSION3.xlsx                  # Master Excel Database for v4
│   ├── sample_invoice.png                  # Sample PDF Invoice Output
│   └── README.md                           # v4 Directory Notes
├── ALTS_Excel_INVOICE1.xlsx                 # Legacy Excel Invoice Template (v1-v3)
├── ALTS_TUTORIAL1.xlsx                     # Legacy Master Database (v1-v3)
├── LICENSE                                 # GNU General Public License v3.0
└── README.md                               # Main Project Documentation
```[cite: 1]

---

## ⚙️ Prerequisites & Setup[cite: 1]

- **Microsoft 365 Account** with Power Automate, OneDrive for Business, and Excel Online access[cite: 1].
- **ALTS Master Database (`ALTS_VERSION3.xlsx` or `ALTS_TUTORIAL1.xlsx`)**: Tracks `Client`, `Service`, `Cost`, `NextRenewal`, `EmailDeterminant`, and client email parameters[cite: 1].
- **HTML Invoice Template (`ALTS-v4/invoice_template.html`)**: Dynamic UTF-8 HTML billing layout[cite: 1].
- **OneDrive Folder (`/Sent_Invoices`)**: Archive folder for sent PDF invoices[cite: 1].

---

## 🛠️ Version 4 Deployment & Setup Guide[cite: 1]

### Step 1: Excel Data Setup[cite: 1]
Ensure your master Excel file (`ALTS_VERSION3.xlsx`) contains two structured tables[cite: 1]:
1. **`ClientTable`**: Columns: `Client`, `AddressLine1`, `AddressLine2`, `AddressLine3`, `Bill_No`, `Invoice_No`, `Date`[cite: 1].
2. **`ServiceTable`**: Columns: `Client`, `Description`, `Qty`, `UnitPrice`, `Amount`[cite: 1].

### Step 2: Import Power Automate Flow Package[cite: 1]
1. Download `ALTS-v4/Send_Multi-Service_Invoices_v4.zip` from this repository[cite: 1].
2. Go to **Power Automate** > **My flows** > **Import** > **Import Package (.zip)**[cite: 1].
3. Re-bind connection references to your tenant's resources[cite: 1]:
   - Excel Online (Business)[cite: 1]
   - OneDrive for Business[cite: 1]
   - Office 365 Outlook[cite: 1]

### Step 3: Global Bill Incrementing Logic[cite: 1]
To prevent duplicate bill numbers upon each flow run[cite: 1]:
1. **List rows present in ClientTable**[cite: 1].
2. **Select** (Select integer values of `Bill_No`)[cite: 1]:
   - **From:** `outputs('List_rows_present_in_ClientTable')?['body/value']`[cite: 1]
   - **Map:** `int(item()?['Bill_No'])`[cite: 1]
3. **Initialize varBillNo**[cite: 1]:
   - **Value Expression:** `max(body('Select'))`[cite: 1]

---

## 📜 Legacy Architecture (v1–v3): Synchronous Excel & Office Script Setup[cite: 1]

*(Maintained for backwards compatibility with `ALTS_Excel_INVOICE1.xlsx` setup)*[cite: 1]

### 1. Lifecycle Tracking Logic (`EmailDeterminant`)[cite: 1]
The master tracking file (`ALTS_TUTORIAL1.xlsx`) calculates the remaining lifecycle days for each service (`EmailDeterminant`)[cite: 1].
- **Condition Criteria:** Triggered when `EmailDeterminant` is between `-1` and `2` (greater than or equal to -1 AND less than or equal to 2)[cite: 1].

### 2. Office Script (`ForceSave`)[cite: 1]
Add an Office Script to your Excel invoice template (`ALTS_Excel_INVOICE1.xlsx`) to force immediate calculation and sync[cite: 1]:

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Force full calculation and commit cell updates
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}
```[cite: 1]

### 3. Legacy Flow Blueprint (v1–v3)[cite: 1]

```text
[ Recurrence / Scheduled Trigger ][cite: 1]
               │
               ▼
   [ List rows present in a table ]  <-- (ALTS_TUTORIAL1.xlsx)[cite: 1]
               │
               ▼
         [ For each ]  (Concurrency = 1)[cite: 1]
               │
               ▼
         [ Condition ]  <-- Evaluates Lifecycle Threshold (-1 to 2 days)[cite: 1]
         │           │
      (False)     (True)[cite: 1]
         │           │
    [ Do Nothing ]   ├──► 1. Get a row (Read Invoice Counter)[cite: 1]
                     ├──► 2. Compose - Calculate Next Invoice[cite: 1]
                     ├──► 3. Compose - Format Invoice String[cite: 1]
                     ├──► 4. Update a row (Populate Template & Increment Counter)[cite: 1]
                     ├──► 5. Run script (ForceSave)[cite: 1]
                     ├──► 6. Delay (5 Seconds Buffer)[cite: 1]
                     ├──► 7. Convert file (OneDrive - Excel to PDF)[cite: 1]
                     ├──► 8. Create file (Save PDF to /Sent_Invoices)[cite: 1]
                     └──► 9. Send an email (V2) (Attach PDF & send to Client)[cite: 1]
```[cite: 1]

---

## 🔍 Troubleshooting & Key Learnings[cite: 1]

| Issue | Root Cause | Solution |
| :--- | :--- | :--- |
| **PDF contains un-incremented bill number (v1-v3)** | OneDrive file conversion reading stale binary cache before Excel writes updates to disk[cite: 1]. | Add `ForceSave` script + 5-second Delay[cite: 1]. |
| **Duplicate Bill Numbers across Runs (v4)** | Static incrementing resetting per run[cite: 1]. | Fetch global max via `max(body('Select'))` from Excel before loop execution[cite: 1]. |
| **Naira Symbol (`₦`) garbled in PDF** | Missing character encoding headers in HTML template[cite: 1]. | Add `<meta charset="UTF-8">` and use `&#8358;` HTML entity in `invoice_template.html`[cite: 1]. |
| **Race conditions in loop** | Concurrent executions editing the same template simultaneously[cite: 1]. | Enforce **Concurrency Control = 1** on the `For each` loop[cite: 1]. |

---

## 📄 License[cite: 1]

This project is open-source and available under the **GNU General Public License v3.0**[cite: 1].