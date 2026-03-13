# Pleadings Checker

A Word VBA macro-based document proofreading system designed for legal pleadings and formal documents. Enforces 34 proofreading rules covering British English spelling, formatting, Hart's Rules compliance, and legal style conventions.

## Features

- **34 configurable proofreading rules** covering spelling, formatting, numbering, quotation marks, footnotes, legal terms, and more
- **Modular architecture** — import only the rule modules you need
- **Interactive form UI** with checkboxes to enable/disable individual rules
- **Yellow highlighting** of flagged text with inline comments
- **Tracked changes** for auto-fixable suggestions
- **JSON report export** with full issue details and statistics
- **Optional page range filtering** to check specific sections only
- **Configurable brand name enforcement** with persistence
- **Smart quote/italic handling** — text in quotes or italics is flagged but not auto-fixed

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
| 13 | Colour Formatting | Flags non-standard font colours |
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

## Installation (Recommended: Modular)

Use the files in the `Code/` directory for the recommended modular installation.

### Prerequisites

- Microsoft Word 2010 or later
- No early-bound references are required (the project uses late binding exclusively)

### Core Files (Required)

These three files must be imported for the system to work:

| File | Purpose |
|------|---------|
| `PleadingsEngine.bas` | Rule dispatcher, highlighting engine, report generator |
| `PleadingsIssue.cls` | Structured issue result class |
| `PleadingsLauncher.bas` **or** `frmPleadingsChecker.frm` | UI (choose one) |

### Optional Rule Modules

Import **any combination** of these — the engine gracefully skips modules that are absent:

| Module | Rules | Description |
|--------|-------|-------------|
| `Rules_Spelling.bas` | 1, 12, 13 | UK/US spelling, licence/license, colour formatting |
| `Rules_TextScan.bas` | 2, 34 | Repeated words, spell out numbers under ten |
| `Rules_Numbering.bas` | 3, 8 | Sequential numbering, clause number format |
| `Rules_Headings.bas` | 4, 21 | Heading capitalisation, title formatting |
| `Rules_Terms.bas` | 5, 7, 23 | Custom terms, defined terms, phrase consistency |
| `Rules_Formatting.bas` | 6, 11 | Paragraph breaks, font consistency |
| `Rules_NumberFormats.bas` | 9, 18, 19 | Date format, page range config, currency format |
| `Rules_Lists.bas` | 10, 15 | Inline list format, list punctuation |
| `Rules_Punctuation.bas` | 14, 16 | Slash style, bracket integrity |
| `Rules_Quotes.bas` | 17, 32, 33 | Quotation marks, single quotes, smart quotes |
| `Rules_FootnoteIntegrity.bas` | 20 | Footnote reference numbering |
| `Rules_Brands.bas` | 22 | Brand name enforcement |
| `Rules_FootnoteHarts.bas` | 24–27 | Hart's Rules for footnotes |
| `Rules_LegalTerms.bas` | 28, 29 | Legal term forms and capitalisation |
| `Rules_Italics.bas` | 30, 31 | Anglicised terms and foreign names |

### Installation Steps

1. Open the VBA Editor: **Alt+F11**
2. Import the 3 core files via **File > Import File**
3. Import whichever optional rule modules you need
4. Close the VBA Editor

### Generating the Form UI

The form file (`frmPleadingsChecker.frm`) is generated by a Python script. All controls are created at runtime — no `.frx` binary file is needed.

```bash
cd Code
python3 generate_form.py
```

This produces a self-contained `.frm` file sized at 720×1080 px with generous spacing for all controls.

### Running the Macro

**Option A — Via the form UI** (requires `frmPleadingsChecker.frm`):
1. Open the document you want to check
2. Run the macro `PleadingsChecker` (via **Developer > Macros**)
3. Select rules, set page range, click **Run Checks**
4. Use **Apply Suggestions** to accept auto-fixable changes as tracked changes
5. Use **Export Report** to save a JSON report

**Option B — Via the dialog launcher** (requires `PleadingsLauncher.bas`):
1. Open the document you want to check
2. Run the macro `PleadingsChecker` — follows MsgBox/InputBox prompts

**Option C — Via VBA code:**
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

## Italic and Quoted Text Handling

The UK/US spelling rules (Rules 1, 12) intelligently handle italic and quoted text:

- **Italic text** — flagged as `possible_error` with a note to review manually; not auto-fixed
- **Quoted text** — flagged as `possible_error` with a note to review manually; not auto-fixed
- **Normal text** — flagged as `error` with an auto-fix suggestion

This prevents unwanted changes to quoted passages, case names, or foreign terms that may intentionally use different spelling conventions.

## Running Tests

The test suite covers Hart-style rules (Rules 24–34) with 42 test cases.

In the VBA Immediate window:
```vba
TestBucket1Rules.RunAllBucket1Tests
```

## Python Test Harness

For environments without Microsoft Word (e.g. Linux CI), a Python-based rule checker is included:

```bash
pip install python-docx
python3 run_checks.py
```

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
├── Code/                          Recommended install (modular)
│   ├── PleadingsEngine.bas            Core engine (REQUIRED)
│   ├── PleadingsIssue.cls             Issue class (REQUIRED)
│   ├── PleadingsLauncher.bas          Dialog-based UI (choose one UI)
│   ├── frmPleadingsChecker.frm        Full form UI (choose one UI)
│   ├── generate_form.py               FRM generator script
│   ├── Rules_Spelling.bas             Rules 1, 12, 13 (optional)
│   ├── Rules_TextScan.bas             Rules 2, 34 (optional)
│   ├── Rules_Numbering.bas            Rules 3, 8 (optional)
│   ├── Rules_Headings.bas             Rules 4, 21 (optional)
│   ├── Rules_Terms.bas                Rules 5, 7, 23 (optional)
│   ├── Rules_Formatting.bas           Rules 6, 11 (optional)
│   ├── Rules_NumberFormats.bas        Rules 9, 18, 19 (optional)
│   ├── Rules_Lists.bas                Rules 10, 15 (optional)
│   ├── Rules_Punctuation.bas          Rules 14, 16 (optional)
│   ├── Rules_Quotes.bas               Rules 17, 32, 33 (optional)
│   ├── Rules_FootnoteIntegrity.bas    Rule 20 (optional)
│   ├── Rules_Brands.bas               Rule 22 (optional)
│   ├── Rules_FootnoteHarts.bas        Rules 24–27 (optional)
│   ├── Rules_LegalTerms.bas           Rules 28, 29 (optional)
│   └── Rules_Italics.bas              Rules 30, 31 (optional)
├── Rules/                             Legacy (individual rule files)
│   ├── Rule01_british-spelling.bas
│   ├── ...
│   └── Rule34_spell-out-under-ten.bas
├── PleadingsEngine.bas                Legacy root-level engine
├── PleadingsIssue.cls                 Legacy root-level class
├── frmPleadingsChecker.frm            Legacy root-level form
├── TestBucket1Rules.bas               Unit tests for Rules 24–34
├── run_checks.py                      Python-based test harness
├── test_pleading.docx                 Test input document
└── test_pleading_TEST1OUTPUT.docx     Annotated output
```

## Severity Levels

- **error** — definite issue requiring correction (auto-fix suggested)
- **warning** — likely issue, review recommended
- **possible_error** — may be intentional (e.g. italic/quoted text, ambiguous context), requires manual review

## Output Formats

1. **In-document highlighting** — yellow background on flagged text with red annotation comments
2. **Tracked changes** — auto-fixable suggestions applied as revisions
3. **JSON report** — machine-readable export with rule name, location, severity, issue description, and suggestion for each finding
