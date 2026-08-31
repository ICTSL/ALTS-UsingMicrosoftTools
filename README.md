# ALTS-UsingMicrosoftTools

![Microsoft Power Automate](https://img.shields.io/badge/Power%20Automate-0066B8?style=for-the-badge&logo=microsoftpowerautomate&logoColor=white)
![Microsoft Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![TypeScript](https://img.shields.io/badge/Office%20Scripts-3178C6?style=for-the-badge&logo=typescript&logoColor=white)
![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=for-the-badge&logo=html5&logoColor=white)

An end-to-end cloud solution built with **Microsoft Power Automate**, **Excel Online**, **Office Scripts**, and **HTML dynamic templates**.

The **Automated Lifecycle Tracking System (ALTS)** monitors client service lifecycle dates, evaluates renewal thresholds, dynamically populates and increments custom invoice templates, converts generated bills to PDF, archives sent copies to OneDrive, and dispatches automated billing emails.

---

## 🌟 What's New in Version 4.0 (`ALTS-v4`)

Version 4 introduces significant architectural improvements for multi-service itemization, HTML template rendering, robust sequential billing, and type-safe data filtering:

- **Type-Safe Multi-Service Lifecycle Filtering:** Enhanced `Filter Services For Client` using robust WDL expressions (`float(coalesce(...))`) to evaluate `EmailDeterminant` ranges (`-1` to `2`) directly across multi-service rows while preventing `String vs Integer` type-mismatch crashes and null cell errors.
- **Dynamic Sequential Bill Incrementing:** Resolved recurring `Bill_No` resets across execution dates by fetching the global maximum bill number from Excel using `max(body('Select'))` before loop execution.
- **HTML & PDF Dynamic Rendering:** Replaced Excel template population with a standalone UTF-8 HTML template (`invoice_template.html`) converted directly to PDF via OneDrive.
- **Naira Currency & UTF-8 Encoding:** Added explicit `<meta charset="UTF-8">` headers and HTML entities (`&#8358;`) to guarantee consistent rendering of the Nigerian Naira (`₦`) symbol.
- **Excel Serial Date Conversion:** Converted raw Excel serial integer dates (e.g., `46254`) into standard formatted string dates (`dd-MMM-yyyy`) using expression mapping.
- **File & Variable Synchronization:** Synchronized dynamic bill numbers across HTML invoice headers, PDF file generation, dynamic email attachments, and Excel tracking updates (`Update a row 1`).

---

## 📺 Video Tutorial & Demonstrations

### Video Guide
Watch the complete step-by-step video guide on YouTube:

[![Watch the Tutorial](https://img.youtube.com/vi/mvzyZDertSg/maxresdefault.jpg)](https://www.youtube.com/watch?v=mvzyZDertSg)

### Workflow Demonstrations
https://github.com/user-attachments/assets/e76e6fdd-48cf-4403-a769-9fecae7b571c

https://github.com/user-attachments/assets/ac277f27-49bd-441a-a12b-30a5e0cab44f

---

## 🏗️ Repository Directory Structure

```text
ALTS-UsingMicrosoftTools/
├── ALTS-v4/
│   ├── Send_Multi-Service_Invoices_v4.zip   # Version 4 Exported Power Automate Flow
│   ├── invoice_template.html               # UTF-8 HTML Invoice Template
│   ├── ALTS_VERSION3.xlsx                  # Master Excel Database for v4
│   ├── sample_invoice.png                  # Sample PDF Invoice Output
│   └── README.md                            # v4 Directory Notes
├── ALTS_Excel_INVOICE1.xlsx                 # Legacy Excel Invoice Template (v1-v3)
├── ALTS_TUTORIAL1.xlsx                      # Legacy Master Database (v1-v3)
├── LICENSE                                  # GNU General Public License v3.0
└── README.md                                # Main Project Documentation
```
⚙️ Prerequisites & SetupMicrosoft 365 Account with Power Automate, OneDrive for Business, and Excel Online access.ALTS Master Database (ALTS_VERSION3.xlsx or ALTS_TUTORIAL1.xlsx): Tracks Client, Service, Cost, NextRenewal, EmailDeterminant, and client email parameters.HTML Invoice Template (ALTS-v4/invoice_template.html): Dynamic UTF-8 HTML billing layout.OneDrive Folder (/Sent_Invoices): Archive folder for sent PDF invoices.🛠️ Version 4 Deployment & Setup GuideStep 1: Excel Data SetupEnsure your master Excel file (ALTS_VERSION3.xlsx) contains two structured tables:ClientTable: Columns: Client, AddressLine1, AddressLine2, AddressLine3, Bill_No, Invoice_No, Date.ServiceTable: Columns: Client, Service, Cost, Billing, LastPaid, Active Status, NextRenewal, Alert, Remarks, EmailDeterminant, QTY, Discount.Step 2: Import Power Automate Flow PackageDownload ALTS-v4/Send_Multi-Service_Invoices_v4.zip from this repository.Go to Power Automate > My flows > Import > Import Package (.zip).Re-bind connection references to your tenant's resources:Excel Online (Business)OneDrive for BusinessOffice 365 OutlookStep 3: Global Bill Incrementing LogicTo prevent duplicate bill numbers upon each flow run:List rows present in ClientTable.Select (Select integer values of Bill_No):From: outputs('List_rows_present_in_ClientTable')?['body/value']Map: int(item()?['Bill_No'])Initialize varBillNo:Value Expression: max(body('Select'))Step 4: Power Automate Workflow Architecture & Logic1. Data Retrieval & InitializationRecurrence / Manual TriggerList rows present in ClientTable (ALTS_VERSION3.xlsx)List rows present in ServiceTable (ALTS_VERSION3.xlsx)Select (Select - Extract Bill Numbers): Extracts array of all Bill_No integers using @int(item()?['Bill_No']).Initialize variable (varBillNo): Sets global counter base using @max(body('Select')).2. Itemization & Dynamic FilteringApply to each (For Each Client): (Concurrency Control = 1)Filter array (Filter Services For Client): Filters ServiceTable matching current Client AND checks EmailDeterminant lifecycle threshold range (-1 to 2):Plaintext@and(
  equals(item()?['Client'], items('For_each_Client')?['Client']),
  greaterOrEquals(float(coalesce(item()?['EmailDeterminant'], '999')), -1),
  lessOrEquals(float(coalesce(item()?['EmailDeterminant'], '999')), 2)
)
Condition (Check If Client Has Eligible Services): Checks if filtered item count is greater than 0 (length(body('Filter_Services_For_Client')) > 0).3. HTML Generation, Conversion & Email DispatchIf True:Increment variable (varBillNo): Increments counter by 1.Get file content (Get HTML Template): Loads /ALTS-v4/invoice_template.html.Select (Build HTML Table Rows): Maps filtered service items to dynamic <tr><td>...</td></tr> HTML rows.Compose (Replace Placeholders in HTML): Replaces placeholders ({{ClientName}}, {{BillNo}}, {{ServiceRows}}). Formats serial dates via addDays('1899-12-30', int(...), 'dd-MMM-yyyy').Create file (Create Temp HTML): Writes temporary file to OneDrive /ALTS-v4/.Convert file (Convert HTML to PDF): Leverages OneDrive conversion engine.Create file (Save PDF Archive): Writes archived PDF to /ALTS-v4/Sent_Invoices/.Send an email (V2): Dispatches email with generated PDF attachment to the client.Update a row 1 (Excel Writeback): Persists updated Bill_No back to ClientTable.Delete file (Clean Up Temp HTML): Deletes temporary HTML file from OneDrive.📜 Legacy Architecture (v1–v3): Synchronous Excel & Office Script Setup(Maintained for backwards compatibility with ALTS_Excel_INVOICE1.xlsx setup)1. Lifecycle Tracking Logic (EmailDeterminant)The master tracking file (ALTS_TUTORIAL1.xlsx) calculates the remaining lifecycle days for each service (EmailDeterminant).Condition Criteria: Triggered when EmailDeterminant is between -1 and 2 (greater than or equal to -1 AND less than or equal to 2).2. Office Script (ForceSave)Add an Office Script to your Excel invoice template (ALTS_Excel_INVOICE1.xlsx) to force immediate calculation and sync:TypeScriptfunction main(workbook: ExcelScript.Workbook) {
    // Force full calculation and commit cell updates
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}
3. Legacy Flow Blueprint (v1–v3)Plaintext[ Recurrence / Scheduled Trigger ]
               │
               ▼
   [ List rows present in a table ]  <-- (ALTS_TUTORIAL1.xlsx)
               │
               ▼
         [ For each ]  (Concurrency = 1)
               │
               ▼
         [ Condition ]  <-- Evaluates Lifecycle Threshold (-1 to 2 days)
         │           │
      (False)     (True)
         │           │
    [ Do Nothing ]   ├──► 1. Get a row (Read Invoice Counter)
                     ├──► 2. Compose - Calculate Next Invoice
                     ├──► 3. Compose - Format Invoice String
                     ├──► 4. Update a row (Populate Template & Increment Counter)
                     ├──► 5. Run script (ForceSave)
                     ├──► 6. Delay (5 Seconds Buffer)
                     ├──► 7. Convert file (OneDrive - Excel to PDF)
                     ├──► 8. Create file (Save PDF to /Sent_Invoices)
                     └──► 9. Send an email (V2) (Attach PDF & send to Client)
🔍 Troubleshooting & Key LearningsIssueRoot CauseSolutionType Mismatch Error in FilterComparing Excel string columns directly against integers in WDL (greaterOrEquals).Wrap cell inputs with float(coalesce(item()?['EmailDeterminant'], '999')).Blank Cell Runtime FailureNull values passed into math or conversion functions (int/float).Use coalesce(..., '999') to supply a dummy value that safely fails range criteria.PDF contains un-incremented bill number (v1-v3)OneDrive file conversion reading stale binary cache before Excel writes updates to disk.Add ForceSave script + 5-second Delay.Duplicate Bill Numbers across Runs (v4)Static incrementing resetting per run.Fetch global max via max(body('Select')) from Excel before loop execution.Naira Symbol (₦) garbled in PDFMissing character encoding headers in HTML template.Add <meta charset="UTF-8"> and use &#8358; HTML entity in invoice_template.html.Race conditions in loopConcurrent executions editing the same template simultaneously.Enforce Concurrency Control = 1 on the For each loop.📄 LicenseThis project is open-source and available under the GNU General Public License v3.0.
