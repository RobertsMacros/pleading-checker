# Pleadings Checker

A Word VBA macro-based document proofreading system designed for legal pleadings and formal documents. Enforces 34 proofreading rules covering British English spelling, formatting, Hart's Rules compliance, and legal style conventions.

## Features

- **34 configurable proofreading rules** covering spelling, formatting, numbering, quotation marks, footnotes, legal terms, and more
- **Interactive form UI** with checkboxes to enable/disable individual rules
- **Yellow highlighting** of flagged text with inline comments
- **Tracked changes** for auto-fixable suggestions
- **JSON report export** with full issue details and statistics
- **Optional page range filtering** to check specific sections only
- **Configurable brand name enforcement** with persistence

## Rules Overview

| # | Rule | Description |
|---|------|-------------|
| 1 | British Spelling | Flags ~95 US English spellings (color→colour, organize→organise, etc.) |
| 2 | Repeated Words | Detects consecutive duplicate words (the the, is is) |
| 3 | Sequential Numbering | Validates clause/section numbering sequence |
| 4 | Heading Capitalisation | Ensures headings start with capital letters |
| 5 | Custom Term Whitelist | Loads user-defined exception terms |
| 6 | Paragraph Break Consistency | Flags inconsistent paragraph spacing |
| 7 | Defined Terms | Tracks quoted/parenthetical definitions; flags unused or inconsistent terms |
| 8 | Clause Number Format | Enforces "N." format (not bracketed) |
| 9 | Date/Time Format | Flags ambiguous or US-style dates |
| 10 | Inline List Format | Enforces semicolons between inline list items |
| 11 | Font Consistency | Flags font changes from the dominant font |
| 12 | Licence/License | Enforces British noun/verb distinction |
| 13 | Colour Formatting | Flags US "color" spelling |
| 14 | Slash Style | Flags slash constructions (and/or, he/she) |
| 15 | List Punctuation | Validates terminal punctuation in list items |
| 16 | Bracket Integrity | Checks matching/balanced brackets (), [], {} |
| 17 | Quotation Mark Consistency | Flags mixed smart/straight quotation marks |
| 18 | Page Range | Configuration-only rule for restricting checks to a page range |
| 19 | Currency/Number Format | Flags dollar signs and formatting issues |
| 20 | Footnote Integrity | Validates footnote/endnote reference numbering |
| 21 | Title Formatting | Checks first paragraph uses Title/Heading style |
| 22 | Brand Name Enforcement | Enforces correct forms (PwC, HMRC, FCA, EY, KPMG, Deloitte) |
| 23 | Phrase Consistency | Flags inconsistent phrasing (notwithstanding vs despite, etc.) |
| 24 | Footnotes Not Endnotes | Requires footnotes, not endnotes (Hart's Rules) |
| 25 | Footnote Terminal Full Stop | Footnotes must end with a period |
| 26 | Footnote Initial Capital | Footnotes must start with a capital (exceptions: ibid, eg, ie, cf) |
| 27 | Footnote Abbreviation Dictionary | Enforces Hart-style abbreviations without dots (eg not e.g., pp not pgs) |
| 28 | Mandated Legal Term Forms | Requires hyphens in Solicitor-General, Attorney-General |
| 29 | Always Capitalise Terms | Enforces capitalisation of Prime Minister, Law Lords, etc. |
| 30 | Anglicised Terms Not Italic | Prima facie, per se, etc. must be roman, not italic |
| 31 | Foreign Names Not Italic | Foreign court/institution names must be roman |
| 32 | Single Quotes Default | Outer quotations must use single quotes, not double |
| 33 | Smart Quote Consistency | All quotes must be curly (smart), not straight |
| 34 | Spell Out Under Ten | Numbers 1–9 must be written out in prose |

## Installation (Word VBA)

### Prerequisites

- Microsoft Word 2010 or later
- Microsoft Scripting Runtime reference

### Steps

1. Open the VBA Editor: **Alt+F11**
2. Enable the Scripting Runtime: **Tools > References** > check **Microsoft Scripting Runtime**
3. Import all project files via **File > Import File**:
   - `PleadingsEngine.bas` — core engine
   - `PleadingsIssue.cls` — structured issue class
   - `frmPleadingsChecker.frm` — user interface form
   - All 34 rule files from the `Rules/` directory (`Rule01_british-spelling.bas` through `Rule34_spell-out-under-ten.bas`)
   - *(Optional)* `TestBucket1Rules.bas` — unit tests for Hart-style rules
4. Close the VBA Editor

### Running the Macro

**Option A — Via the form UI:**
1. Open the document you want to check
2. Run the macro `PleadingsChecker` (via **Developer > Macros**, or assign it to a ribbon button/keyboard shortcut)
3. The form opens with checkboxes for all 34 rules
4. Select the rules you want to apply (or use **Select All**)
5. Optionally set a page range to restrict checks
6. Click **Run Checks**
7. Issues are highlighted in yellow with comments in the document
8. Use **Apply Suggestions** to accept auto-fixable changes as tracked changes
9. Use **Export Report** to save a JSON report

**Option B — Via VBA code:**
```vba
Dim config As Scripting.Dictionary
Set config = PleadingsEngine.InitRuleConfig()
' Optionally disable specific rules:
' config("british_spelling") = False

Dim allIssues As Collection
Set allIssues = PleadingsEngine.RunAllPleadingsRules(ActiveDocument, config)

' Apply highlights
PleadingsEngine.ApplyHighlights ActiveDocument, allIssues

' Export JSON report
PleadingsEngine.GenerateReport allIssues, "C:\path\report.json"
```

## Running Tests

The test suite covers Hart-style rules (Rules 24–34) with 42 test cases.

In the VBA Immediate window:
```vba
TestBucket1Rules.RunAllBucket1Tests
```

Output:
```
========================================
  Bucket 1 Rule Tests
========================================
  PASS: FootnotesNotEndnotes: footnotes only -> pass (count=0)
  PASS: FootnotesNotEndnotes: endnotes only -> fail (count=1 >= 1)
  ...
========================================
  PASSED: 42
  FAILED: 0
  TOTAL:  42
========================================
```

## Python Test Harness

For environments without Microsoft Word (e.g. Linux CI), a Python-based rule checker is included:

```bash
pip install python-docx
python3 run_checks.py
```

This creates a test document (`test_pleading.docx`) with content triggering all 34 rules, runs Python equivalents of each rule, and saves an annotated output (`test_pleading_TEST1OUTPUT.docx`) with yellow highlighting and issue annotations.

## Brand Name Configuration

Brand rules are configurable via the form UI or by editing the persistent file at:
```
%APPDATA%\PleadingsChecker\brand_rules.txt
```

Format: `CorrectForm=variant1,variant2,variant3`

Default brands: PwC, Deloitte, HMRC, FCA, EY, KPMG

## Project Structure

```
pleading-checker/
├── PleadingsEngine.bas          Core rule coordinator and engine
├── PleadingsIssue.cls           Structured issue result class
├── frmPleadingsChecker.frm      User interface form
├── TestBucket1Rules.bas         Unit tests for Rules 24–34
├── run_checks.py                Python-based test harness
├── Rules/
│   ├── Rule01_british-spelling.bas
│   ├── Rule02_repeated-words.bas
│   ├── ...
│   └── Rule34_spell-out-under-ten.bas
├── test_pleading.docx           Test input document
└── test_pleading_TEST1OUTPUT.docx  Annotated output
```

## Severity Levels

- **error** — definite issue requiring correction
- **warning** — likely issue, review recommended
- **possible_error** — may be intentional (e.g. "had had"), requires manual review

## Output Formats

1. **In-document highlighting** — yellow background on flagged text with red annotation comments
2. **Tracked changes** — auto-fixable suggestions applied as revisions
3. **JSON report** — machine-readable export with rule name, location, severity, issue description, and suggestion for each finding
