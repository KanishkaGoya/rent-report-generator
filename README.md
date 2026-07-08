\# Rent Report Generator



A Streamlit web application that generates a formatted \*\*Rent Report\*\* by

merging a Vendor Master file with a Rent Report file, both supplied as

Excel (`.xlsx`) uploads. The app performs validation, cleans and aggregates

the data, matches suppliers to vendors, and produces a fully formatted

Excel workbook — all in memory, with no data written to disk.



\---



\## Purpose



Organizations that pay rent to multiple landlords/vendors typically export

two separate Excel files each period:



1\. A \*\*Vendor Master\*\* containing vendor codes, names, GST numbers, and

&#x20;  addresses.

2\. A \*\*Rent Report\*\* containing raw rent transactions per supplier

&#x20;  (which may include duplicate entries, and both positive and negative

&#x20;  amounts, e.g. reversals or credit notes).



This application merges those two files into a single, professionally

formatted "Detail of Rent Expenses" report — replicating the layout and

formatting previously produced via a VBA macro — while also flagging any

suppliers that could not be matched to a vendor.



\---



\## Input Files



\### 1. Vendor Master (`.xlsx`)



Must contain \*\*exactly\*\* the following headers (case-sensitive):



| Column                  |

|-------------------------|

| Vendor                  |

| City                    |

| Name 1                  |

| GST Registration No.    |

| Name 2                  |

| Street                  |

| Name 3                  |

| Name 4                  |

| PostalCode              |



\### 2. Rent Report (`.xlsx`)



May contain any number of columns, but only the following two are used;

all others are ignored:



| Column   |

|----------|

| Supplier |

| Amount   |



\---



\## Processing Rules



\- \*\*Duplicate suppliers\*\* in the Rent Report are grouped and their

&#x20; `Amount` values are summed. Positive and negative amounts naturally

&#x20; offset one another (e.g. `5000 + (-1000) + 300 = 4300`).

\- Suppliers whose net summed amount is \*\*zero are still included\*\* in the

&#x20; report — they are never dropped.

\- Merging matches \*\*Supplier\*\* (Rent Report) to \*\*Vendor\*\* (Vendor

&#x20; Master).

\- Addresses are built in this order, skipping any blank fields:

&#x20; `Name 2` → `Street` → `Name 3` → `Name 4` → `City - PostalCode`

&#x20; (or just `City` if `PostalCode` is blank).

\- Suppliers present in the Rent Report but \*\*not found\*\* in the Vendor

&#x20; Master are placed in a separate `Missing\_Vendors` sheet instead of the

&#x20; main report.



\---



\## Output



The app produces a single downloadable file: \*\*`Final Report.xlsx`\*\*,

containing two sheets:



1\. \*\*Rent Report\*\* — the formatted report with:

&#x20;  - A 3-line header: Company Name, `ASSESSMENT YEAR <year>`, and Report

&#x20;    Title (Arial 14 Bold, centered).

&#x20;  - Column headers: `Vendor Code`, `Particulars`, `GST No`,

&#x20;    `Amount(Rs)`.

&#x20;  - One block per supplier: Vendor Code / Vendor Name (uppercase, bold,

&#x20;    Arial 13) / GST No / Amount (right-aligned, `#,##0.00` format),

&#x20;    followed by each address line on its own row, followed by one blank

&#x20;    spacer row before the next supplier.

2\. \*\*Missing\_Vendors\*\* — a simple two-column sheet (`Supplier`, `Amount`)

&#x20;  listing any suppliers from the Rent Report that could not be matched

&#x20;  to a Vendor Master entry.



The file is generated entirely \*\*in-memory\*\* using `BytesIO` — nothing is

ever written to local disk, making the app safe to run in ephemeral cloud

environments.



\---



\## How to Run Locally



1\. \*\*Clone or download\*\* this project folder.



2\. \*\*Create and activate a virtual environment\*\* (recommended):



```bash

&#x20;  python3 -m venv venv

&#x20;  source venv/bin/activate      # On Windows: venv\\Scripts\\activate

```



3\. \*\*Install dependencies:\*\*



```bash

&#x20;  pip install -r requirements.txt

```



4\. \*\*Run the Streamlit app:\*\*



```bash

&#x20;  streamlit run app.py

```



5\. Streamlit will print a local URL (typically `http://localhost:8501`).

&#x20;  Open it in your browser.



6\. Upload your Vendor Master and Rent Report `.xlsx` files, fill in the

&#x20;  Company Name / Assessment Year / Report Title fields, and click

&#x20;  \*\*Generate Report\*\*. Once processing completes, click

&#x20;  \*\*Download Final Report.xlsx\*\* to save the result.



\---



\## How to Deploy to Streamlit Community Cloud



1\. \*\*Push this project to a GitHub repository.\*\* Ensure the repository

&#x20;  includes at minimum:

&#x20;  - `app.py`

&#x20;  - `report\_generator.py`

&#x20;  - `requirements.txt`



2\. Go to \*\*https://share.streamlit.io\*\* and sign in with your GitHub

&#x20;  account.



3\. Click \*\*"New app"\*\*.



4\. Select:

&#x20;  - \*\*Repository:\*\* the GitHub repo you just pushed.

&#x20;  - \*\*Branch:\*\* `main` (or your default branch).

&#x20;  - \*\*Main file path:\*\* `app.py`



5\. Click \*\*"Deploy"\*\*.



6\. Streamlit Community Cloud will automatically install the packages

&#x20;  listed in `requirements.txt` and launch the app. Once deployed, you'll

&#x20;  receive a public URL (e.g. `https://your-app-name.streamlit.app`) that

&#x20;  you can share with others.



7\. \*\*Note:\*\* Since the app performs all file processing in-memory (no

&#x20;  disk writes), it is fully compatible with Streamlit Community Cloud's

&#x20;  ephemeral, read-only-friendly filesystem — no additional configuration

&#x20;  is required.



\---



\## Project Structure

