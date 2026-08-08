# ALTS-UsingMicrosoftTools

![Microsoft Power Automate](https://img.shields.io/badge/Power%20Automate-0066B8?style=for-the-badge&logo=microsoftpowerautomate&logoColor=white)
![Microsoft Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![TypeScript](https://img.shields.io/badge/Office%20Scripts-3178C6?style=for-the-badge&logo=typescript&logoColor=white)

An end-to-end cloud solution built with **Microsoft Power Automate**, **Excel Online**, and **Office Scripts**. 

The **Automated Lifecycle Tracking System (ALTS)** monitors client service lifecycle dates, evaluates renewal thresholds, dynamically populates and increments custom invoice templates, forces synchronous workbook calculations, converts generated bills to PDF, archives sent copies to OneDrive, and dispatches automated billing emails.

---

## 📺 Video Tutorial & Demonstration

Watch the complete step-by-step video guide on YouTube:

[![Watch the Tutorial](https://img.youtube.com/vi/mvzyZDertSg/maxresdefault.jpg)](https://www.youtube.com/watch?v=mvzyZDertSg)



**Workflow Demonstration:**


https://github.com/user-attachments/assets/e76e6fdd-48cf-4403-a769-9fecae7b571c



https://github.com/user-attachments/assets/ac277f27-49bd-441a-a12b-30a5e0cab44f



---

## 🏗️ System Architecture & Logic

ALTS operates using a two-stage process: **Lifecycle Evaluation** followed by **Automated Document & Email Generation**.

### 1. Lifecycle Tracking Logic (`EmailDeterminant`)
The master tracking file (`ALTS_TUTORIAL1.xlsx`) calculates the remaining lifecycle days for each subscription/service (`EmailDeterminant`). 
* **Condition Criteria:** Automated invoice generation and notification email triggers **only when the days remaining are less than 2 days and not more than 1 day overdue** (e.g., `EmailDeterminant` between `-1` and `2`).

### 2. Synchronous Invoice Generation Architecture
When automating Excel updates and PDF conversions in Power Automate, Excel Online often defers writing binary file changes immediately. To prevent stale data or duplicate invoice numbers:
1. Fetch the master invoice counter and calculate the incremented Bill No in memory.
2. Update the dynamic invoice template (`ALTS_Excel_INVOICE1.xlsx`).
3. Execute a TypeScript **Office Script** (`ForceSave`) to force full calculation and sync.
4. Apply a **5-second delay buffer** for cloud sync stabilization.
5. Convert the saved template directly to **PDF**.
6. Archive the PDF into the `/Sent_Invoices` directory on OneDrive.
7. Send the email with the attached PDF invoice.

---

## ⚙️ Prerequisites & Setup

- **Microsoft 365 Account** with Power Automate, OneDrive for Business, and Excel Online access.
- **ALTS Master Database (`ALTS_TUTORIAL1.xlsx`)**: Tracks `Client`, `Service`, `Cost`, `NextRenewal`, `EmailDeterminant`, and client email parameters.
- **ALTS Invoice Template (`ALTS_Excel_INVOICE1.xlsx`)**: Formatted Excel layout linked to invoice counters.
- **OneDrive Folder (`/Sent_Invoices`)**: Archive folder for sent PDF invoices. (Use your own folder path here)

---

## 📜 Step 1: Office Script Setup (`ForceSave`)

Before building the flow, add an Office Script to your Excel invoice template (`ALTS_Excel_INVOICE1.xlsx`) to force immediate calculation.

1. Open **`ALTS_Excel_INVOICE1.xlsx`** in Excel Online.
2. Go to **Automate** $\rightarrow$ **New Script**.
3. Clear the default script editor text and paste:

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Force full calculation and commit cell updates
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}
```
Save the script as ForceSave.

🔄 Step 2: Complete Power Automate Flow Blueprint
Flow Diagram
```
[ Recurrence / Scheduled Trigger ]
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
```
                     

📋 Step 3: Action-by-Action Flow Configuration
Stage A: Trigger & Loop Filter
Trigger: Recurrence (e.g., Daily schedule).

List rows present in a table: Points to ALTS_TUTORIAL1.xlsx.

For each:

Settings: Set Concurrency Control to On with a Degree of Parallelism of 1 to avoid file-locking during script calculation.

Condition:

Filter rows based on lifecycle days:
EmailDeterminant is greater than or equal to -1 AND EmailDeterminant is less than or equal to 2.

Stage B: True Branch Pipeline
1. Fetch Current Invoice Counter
Action: Get a row (Excel Online)

File: ALTS_Excel_INVOICE1.xlsx

2. Compute Next Invoice Number
Action: Compose (Rename: Compose-Calculate Next Invoice)

Expression: add(int(outputs('Get_a_row')?['body/Bill_No']), 1)

3. Format Invoice ID String
Action: Compose (Rename: Compose-Format Invoice String)

Expression: concat('ITS_2026_', outputs('Compose-Calculate_Next_Invoice'))

4. Update Invoice Template
Action: Update a row (Excel Online)

File: ALTS_Excel_INVOICE1.xlsx

Write the updated client billing data, service costs, and new Bill_No into the template fields.

5. Force Excel Sync
Action: Run script (Excel Online)

File: ALTS_Excel_INVOICE1.xlsx

Script: ForceSave

6. Synchronization Buffer
Action: Delay (Schedule)

Count: 5 | Unit: Seconds

7. Convert File to PDF
Action: Convert file (OneDrive for Business)

File: ALTS_Excel_INVOICE1.xlsx

Target Type: PDF

8. Archive PDF to Sent_Invoices Folder
Action: Create file (OneDrive for Business)

Folder Path: /Sent_Invoices

File Name: concat(outputs('Compose-Format_Invoice_String'), '.pdf')

File Content: Dynamic File content output from Convert file.

9. Dispatch Email Notification
Action: Send an email (V2) (Office 365 Outlook)

To: Dynamic Emails field from ALTS_TUTORIAL1.xlsx.

CC: Dynamic In_copy field.

Subject: ALTS Invoice Notice: @{outputs('Compose-Format_Invoice_String')}

Attachment Name: @{outputs('Compose-Format_Invoice_String')}.pdf

Attachment Content: Dynamic File content from Convert file.

🔍 Troubleshooting & Key Learnings
Issue	Root Cause	Solution
PDF contains un-incremented bill number	OneDrive file conversion reading stale binary cache before Excel writes updates to disk.	Add ForceSave script + 5-second Delay.
Race conditions in loop	Default concurrent executions editing the same Excel template simultaneously.	Enforce Concurrency Control = 1 on the For each loop.
📄 License
This project is open-source and available under the GNU General Public License v3.0.
