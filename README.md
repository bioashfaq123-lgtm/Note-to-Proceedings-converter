# Note → Proceeding  (Agricultural Polytechnic, Jagtial)

Upload a **Note (.docx)** → review → download the **Proceeding (.docx)** in the office format.
Works fully offline in the browser. Nothing is uploaded to any server.

## How to run
Double-click `index.html` (Chrome / Edge recommended). That's it.
To put it on an office computer/intranet, copy the whole folder to any web server
(or host free on Netlify / GitHub Pages / Vercel — it is a static site).

## What it reads from the Note automatically
- Note date (`Dt: …`)
- Items table (item, qty, rate, amount) and the total (warns if the total does not add up)
- Signatory designation (used for "Note submitted by…", "To", and "sanction accorded to")
- Head of account (adds 121541 – Agricultural Polytechnic, Jagtial; "Revolving Fund" etc.)
- Purpose (e.g. "Purchase of plant protection chemicals for Lemon orchard") — a best guess, always editable

## What you type once per Proceeding
Proc. No. (the next number is suggested after each download), supplier/firm, and any wording changes.
Everything is shown in a live preview before you download.

## What it generates
Amount in Indian-system words, subject line, references (note added last), order paragraph,
items table with total, payment line, head of account, PRINCIPAL block, To / CC list — **Legal page,
Times New Roman 12 pt, university logo** as in your sample Proceedings.

## Customising
| What | Where |
|---|---|
| Office name, "Present:", Proc. No. suffix, standard head of account | **⚙ Office settings** (top right) |
| References, CC list, delegation clause | Section 4 of the form → **Save … as my defaults** |
| Acronym expansions (PAMP → Participatory Agricultural Management Programme) | `js/parser.js` → `NP.ACRONYMS` |
| Wording of subject / order paragraph | `js/generator.js` → `NP.derive()` |
| Layout (fonts, widths, margins) | `js/generator.js` → `NP.buildDocx()` |

## Folder
```
index.html          landing page + review screen
css/style.css       styling
js/parser.js        reads the Note .docx
js/generator.js     builds the Proceeding .docx
js/numwords.js      amount in words (lakh/crore system)
js/app.js           page logic, live preview, saved defaults
js/logo.js          university logo (embedded)
vendor/             JSZip + docx libraries (bundled, no internet needed)
samples/            your 2 sample notes and 2 sample proceedings for testing
```

## Notes / limits
- Supports `.docx` Notes laid out like your samples (a "Sir," paragraph, an items table, a head-of-account line, a signature block). A Note laid out very differently may need a few fields corrected in the review screen.
- The subject/purpose is detected with simple rules, not AI — always glance at it before downloading.
