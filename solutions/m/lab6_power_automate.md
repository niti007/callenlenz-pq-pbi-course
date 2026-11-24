# Lab 6 - Email every cardholder their own card transactions (Power Automate)

Goal: after Power Query has cleaned the card statements, send each of the 10 cardholders an email that
lists **only their own** transactions (with their total), at the click of a button.

Three ways are shown. Build **Option A** (cloud flow) in the lab; **B** and **C** are quick fallbacks.

---

## Option A - Power Automate cloud flow

### 0. Prerequisites
- A work/school Microsoft 365 account with Power Automate (cloud flows) and OneDrive for Business or SharePoint.
- Excel Online connector available (standard connectors only - no premium licence needed).
- Your workbook must be **saved in OneDrive for Business / SharePoint** (not a local drive, not "Personal" OneDrive).
- Power Automate can only read **Excel Tables** - so every data range below must be formatted as a Table
  (Insert > Table). Power Query "Close & Load to Table" already does this.
- **Demo safety:** in the training data the emails are `first.last@callenlenz-demo.example` - these do not exist.
  Before running, **swap the Email column for your own address (e.g. yourname+oliver@yourcompany.com)** in
  `CardholderSummary`, or use the "Override" step in section 6. Never send the test to real colleagues.

### 1. Prepare the workbook (`Card_Flow.xlsx` in OneDrive)
1. Sheet **Transactions**: load the `Transactions` query (lab6_cards.m) to a table. Rename the table
   (Table Design > Table Name) to **`Transactions`**. Keep these columns: Date, Merchant, Category, City, Amount GBP, Last Four, Employee Name, Email.
   - Make **Last Four** text and Date a plain text `dd/MM/yyyy` column too - dates arrive in flows as serial numbers otherwise. Do this in a COPY, not in `Transactions` itself: right-click `Transactions` > Reference, add the step `Date.ToText([Date], "dd/MM/yyyy")` there, and load that copy to the table the flow reads. Keep the original `Transactions` query unchanged (the Lab 6 pivots and the one-tab-per-person step still need a real Date).
2. Sheet **CardholderSummary**: Reference `Transactions` > Group By **Employee Name, Email, Last Four** > add
   `Transactions` (Count Rows) and `Total GBP` (Sum of Amount GBP) > Load to table named **`CardholderSummary`**.
   It should have 10 rows (one per cardholder).
3. Close the workbook in Excel (a flow may fail to read a file that is still open or locked by an unsaved session).

**Step-1 checklist before you build the flow**
- [ ] The table the flow reads is named exactly `Transactions`, with Last Four as text and a text `Date` column. Build it as a Reference copy of the original `Transactions` query with the step `Date.ToText([Date], "dd/MM/yyyy")` (capital MM, because `mm` means minutes), and keep the original query unchanged.
- [ ] `CardholderSummary` table exists with 10 rows and columns Employee Name, Email, Last Four, Transactions, Total GBP.
- [ ] The Email column in `CardholderSummary` has been swapped for your own fake/test addresses (see Prerequisites).
- [ ] Workbook saved in OneDrive for Business / SharePoint and closed.

### 2. Create the flow
1. make.powerautomate.com > **Create** > **Instant cloud flow** > name `Email card transactions` >
   trigger **Manually trigger a flow** > Create.
   (For a monthly send choose **Scheduled cloud flow** instead - Repeat every 1 month, starting on the 5th at 09:00.)

### 3. Action 1 - List rows present in a table (CardholderSummary)
- **+ New step** > search **Excel Online (Business)** > action **List rows present in a table**.
- Rename the action to `List CardholderSummary` (click the title > Rename).
- Location = OneDrive for Business (or your SharePoint site); Document Library = OneDrive / Documents;
  File = `Card_Flow.xlsx`; Table = **CardholderSummary**.

### 4. Action 2 - List rows present in a table (Transactions) - BEFORE the loop
Read the transactions once, outside the loop, so the flow does not re-read the sheet 10 times.
- **+ New step** > Excel Online > **List rows present in a table**, rename `List Transactions`.
- Same file; Table = **Transactions**. (Settings > Pagination ON, threshold 5000 so nothing is cut at 256 rows.)

### 5. Action 3 - Apply to each
- **+ New step** > **Control** > **Apply to each**.
- Select an output: **value** from `List CardholderSummary` (Dynamic content).
- All following actions go **inside** this loop.

**5a. Filter array**
- **Data Operation > Filter array**, rename `Filter this cardholder`.
- From: **value** of `List Transactions` (the action above the loop).
- Condition (switch to **Edit in advanced mode** and paste):
  ```
  @equals(item()?['Last Four'], items('Apply_to_each')?['Last Four'])
  ```
  If Last Four arrives as a number, use `string(item()?['Last Four'])` and pad with
  `formatNumber(int(items('Apply_to_each')?['Last Four']),'0000')`. Easiest: keep it text in the table.
  (Leading-zero cards - 0937 and 0184 - are the usual trap.)

**5b. Create HTML table**
- **Data Operation > Create HTML table**, rename `Build transaction table`.
- From: **Body** of `Filter this cardholder`.
- Columns: **Custom**. Header / Value pairs:

  | Header | Value |
  |---|---|
  | Date | `item()?['Date']` |
  | Merchant | `item()?['Merchant']` |
  | Category | `item()?['Category']` |
  | City | `item()?['City']` |
  | Amount GBP | `formatNumber(item()?['Amount GBP'], 'N2', 'en-GB')` |

**5c. Send an email (V2)**
- **Office 365 Outlook > Send an email (V2)**, rename `Email cardholder`.
- To: `items('Apply_to_each')?['Email']`
- Subject: `concat('Your card transactions - ', items('Apply_to_each')?['Employee Name'])`
- Body (use the `</>` code view button):
  ```
  <p>Hi @{first(split(items('Apply_to_each')?['Employee Name'],' '))},</p>
  <p>You have <b>@{items('Apply_to_each')?['Transactions']}</b> card transactions this quarter totalling
     <b>&pound;@{formatNumber(items('Apply_to_each')?['Total GBP'],'N2','en-GB')}</b>.</p>
  @{body('Build_transaction_table')}
  <p>Please attach receipts using the Expense Template and return to Finance by Friday.</p>
  ```
  (If your action names differ, the internal names use underscores: `Build_transaction_table`, `Apply_to_each`.)

### 6. Test
1. Edit the `CardholderSummary` Email column so every address is yours (`you+oliver@yourcompany.com` style plus-addressing works on M365) - **or**
   add a `Compose` before 5d and change the To field to your own email.
2. Save > **Test** > Manually > Run flow. Open **Run history**: click each action to see Inputs/Outputs.
3. Common failures:
   - *"Table not found"* - the table name must match exactly and the file must be saved/closed.
   - *Only 256 rows* - turn on Pagination in List rows > Settings.
   - *Empty Filter array* - Last Four type mismatch (text vs number) - compare `string()` of both sides.
   - *Dates look like 46054* - convert to text in Power Query first (step 1).
4. Expected outcome: 10 emails; e.g. one cardholder's email lists their rows only, and the grand total over all
   10 emails equals the Q1 total of `Transactions`.

---

## Option B - One tab per person (PivotTable "Show Report Filter Pages")
1. Load `Transactions` to a table. Insert > PivotTable (new sheet).
2. Rows: Date, Merchant, Category, City. Values: Sum of Amount GBP. **Filters: Employee Name.**
3. Design > Report Layout > Show in Tabular Form; Grand totals on for rows only.
4. PivotTable Analyze > Options (arrow) > **Show Report Filter Pages...** > pick **Employee Name** > OK.
   Excel creates one sheet per cardholder (10 tabs), each filtered to that person.
5. Right-click a tab > Move or Copy > (new book) to send one person's sheet. Convert to values before sending.

## Option C - Word mail merge (one letter per cardholder)
1. Prepare a Word letter with merge fields: `«Employee Name»`, `«Transactions»`, `«Total GBP»`.
2. Mailings > Select Recipients > Use an Existing List > `Card_Flow.xlsx` > sheet **CardholderSummary**.
3. Insert merge fields; for the transaction list, either paste a per-person table (Option B) or attach the per-person tab.
4. Mailings > Finish & Merge > **Send Email Messages** > To: `Email` > Subject line > OK (needs Outlook).
   Use **Edit individual documents** first to preview.

---

## Which option when?
| Need | Option |
|---|---|
| Fully automatic, repeatable every month | A |
| One-off, everything stays in Excel | B |
| Formal letter / printed pack | C |
