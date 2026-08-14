Power Automate Invoice Generator - Version 3.0.0 (Testt Copy2)
📌 Overview
Version 3 enhances the invoice flow by adding multi-service line item indexing per client account. It extracts unique clients, filters individual services per client loop, updates sequential line items in Excel using an integer variable counter, and sends out a consolidated PDF invoice.

🛠️ Complete Flow Architecture & Action Sequence
1. Root Setup & Client Selection
Recurrence: Scheduled flow execution.

List rows present in a table: Fetches raw service billing data from Excel.

Initialize variable: Declares the integer variable (RowIndex) at the flow root.

Select: Extracts client identification fields from the table output.

Compose Unique Clients: Filters unique client names using @union() expressions to ensure one invoice per client.

2. Client Loop (For each)
Set variable: Resets RowIndex to 1 for every new client iteration.

Filter array: Isolates all service items belonging specifically to the current client.

Condition: Checks if the filtered records meet billing generation rules (True path continues).

3. Header Updates & Nested Service Iteration (True Branch)
Get a row: Fetches current invoice sequence metadata from the control table.

Compose - Calculate Next Invoice: Increments the sequential invoice number.

Compose - Format Invoice String: Formats string output (e.g., ITS_2026_1049).

Update a row: Writes client details, invoice number, and date into Excel template header fields.

For each Service (Nested Loop):

Update a row 1: Writes DESCRIPTION, QTY, and UNIT PRICE into row RowIndex of the line items table.

Increment variable: Increments RowIndex by 1 to move to the next line item position.

4. Excel Processing & Email Dispatch
Run script: Triggers Office Script to execute formula recalculation across the active sheet.

Delay: Adds a buffer delay to allow OneDrive cloud file synchronization.

Convert file: Converts the updated Excel invoice file into PDF format.

Create file: Generates and stores the PDF invoice file in OneDrive storage.

Send an email (V2): Dispatches the generated PDF invoice attachment to the client's email address.

📝 Key Improvements over Version 2
Introduced top-level client array deduplication (Select + Compose Unique Clients).

Implemented dynamic row indexing (RowIndex variable incrementation inside nested For each Service loop).

Fully supports single-page dynamic multi-service consolidation.
