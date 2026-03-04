# zotero-field-insert

Insert Zotero-compatible field codes into Word (.docx) documents programmatically, with optional LLM-powered citation formatting.

## Problem

When you generate Word documents from Markdown (via pandoc), LaTeX, or Python, there is no way to include live Zotero citations. Pandoc's `--citeproc` produces plain text references that the Zotero Word plugin cannot recognize or manage. This toolkit bridges that gap.

## Tools

### 1. `zotero_field_insert.py` — Insert Zotero field codes

Finds `[@citationkey]` markers in a `.docx` file and replaces each with proper `ADDIN ZOTERO_ITEM CSL_CITATION` Word field codes. Also inserts an `ADDIN ZOTERO_BIBL` field for the bibliography.

### 2. `llm_reference_formatter.py` — LLM-powered citation formatting

Uses Claude to identify informal citations in a markdown document (e.g., "Banwell et al., 2023"), looks up full metadata via CrossRef, optionally adds items to your Zotero library, and outputs pandoc-ready markdown with `[@citekey]` markers plus a CSL JSON bibliography.

## Installation

```bash
pip install python-docx anthropic requests
```

The `anthropic` and `requests` packages are only needed for `llm_reference_formatter.py`. Set the `ANTHROPIC_API_KEY` environment variable to use the LLM formatter.

## Usage

### Basic (standalone bibliography)

```bash
python zotero_field_insert.py document.docx --bib references.json
```

### With Zotero library linking

```bash
python zotero_field_insert.py document.docx \
    --bib references.json \
    --keymap keymap.json \
    --zotero-db ~/Zotero/zotero.sqlite \
    --output document_with_zotero.docx
```

### Arguments

| Argument | Required | Description |
|---|---|---|
| `input` | Yes | Input `.docx` file containing `[@key]` citation markers |
| `--bib`, `-b` | Yes | CSL JSON bibliography file |
| `--output`, `-o` | No | Output path (default: overwrites input) |
| `--keymap`, `-k` | No | JSON mapping citation keys to Zotero item keys |
| `--zotero-db` | No | Path to `zotero.sqlite` for library linking |
| `--font` | No | Font for citation superscripts (default: Calibri) |
| `--size` | No | Font size in pt (default: 11) |
| `--bib-heading` | No | Heading text for bibliography location (default: References) |

## Input files

### Bibliography (CSL JSON)

An array of CSL-JSON items, each with an `id` matching the citation keys in your document:

```json
[
  {
    "id": "banwell2023",
    "type": "article-journal",
    "title": "Diagnosis of myelin oligodendrocyte glycoprotein...",
    "author": [{"family": "Banwell", "given": "B"}],
    "container-title": "The Lancet Neurology",
    "volume": "22",
    "issue": "3",
    "page": "268-282",
    "issued": {"date-parts": [[2023]]},
    "DOI": "10.1016/S1474-4422(22)00431-8"
  }
]
```

You can export this from Zotero (File > Export Library > CSL JSON), or build it from the Zotero SQLite database, or write it by hand.

### Keymap (optional)

Maps citation keys to Zotero item keys so the plugin can link back to your library:

```json
{
  "banwell2023": "BBJKWN9G",
  "sattarnezhad2018": null
}
```

Set a key to `null` if the item is not in your Zotero library. The script will generate a placeholder URI.

## LLM Reference Formatter

### What it does

1. Reads a markdown document with informal citations (e.g., "Banwell et al., 2023", "(Smith 2020)")
2. Uses Claude to identify every citation and extract author/year/title metadata
3. Searches CrossRef for full bibliographic metadata (DOI, journal, volume, pages)
4. Optionally looks up items in your local Zotero database or adds new ones via the Zotero Web API
5. Rewrites the document with `[@citekey]` pandoc citation markers
6. Outputs `references.json` (CSL JSON) and `keymap.json` for `zotero_field_insert.py`

### Usage

```bash
# Basic: identify citations and format
python llm_reference_formatter.py paper.md -o paper_cited.md

# With Zotero library lookups (local DB)
python llm_reference_formatter.py paper.md -o paper_cited.md \
    --zotero-db ~/Zotero/zotero.sqlite

# With Zotero Web API (add missing items to your library)
python llm_reference_formatter.py paper.md -o paper_cited.md \
    --zotero-db ~/Zotero/zotero.sqlite \
    --zotero-api-key YOUR_API_KEY \
    --zotero-library-id 1793208

# Dry run (show what would be done)
python llm_reference_formatter.py paper.md --dry-run
```

### Arguments

| Argument | Required | Description |
|---|---|---|
| `input` | Yes | Input markdown file with informal citations |
| `--output`, `-o` | No | Output markdown path (default: `input_cited.md`) |
| `--bib-output` | No | CSL JSON output path (default: `references.json`) |
| `--keymap-output` | No | Keymap output path (default: `keymap.json`) |
| `--zotero-api-key` | No | Zotero Web API key (for adding items to library) |
| `--zotero-library-id` | No | Zotero user library ID |
| `--zotero-db` | No | Path to local `zotero.sqlite` (for key lookups) |
| `--model` | No | Claude model (default: claude-sonnet-4-20250514) |
| `--no-crossref` | No | Skip CrossRef lookups |
| `--dry-run` | No | Preview without writing files |

### Getting a Zotero API key

1. Go to https://www.zotero.org/settings/keys
2. Create a new key with read/write access to your library
3. Your library ID is visible in the URL when viewing your library on zotero.org

## Full Pipeline

The two tools work together for a complete workflow:

```bash
# 1. Write your document with informal citations
#    "MOGAD is diagnosed using criteria (Banwell et al., 2023)"

# 2. Format citations with the LLM tool
python llm_reference_formatter.py paper.md -o paper_cited.md \
    --zotero-db ~/Zotero/zotero.sqlite

# 3. Convert to docx with pandoc (no --citeproc)
pandoc paper_cited.md -o paper.docx --reference-doc=template.docx

# 4. Insert Zotero field codes
python zotero_field_insert.py paper.docx \
    --bib references.json \
    --keymap keymap.json \
    --zotero-db ~/Zotero/zotero.sqlite

# 5. Open paper.docx in Word, click Zotero > Refresh
```

## Workflow with pandoc (manual)

A typical workflow for generating a Word document from Markdown with live Zotero citations:

```bash
# 1. Write your markdown with [@key] citation markers
#    (same syntax as pandoc-citeproc, but don't use --citeproc)

# 2. Convert to docx with pandoc (no --citeproc flag)
pandoc paper.md -o paper.docx --reference-doc=template.docx

# 3. Insert Zotero field codes
python zotero_field_insert.py paper.docx \
    --bib references.json \
    --keymap keymap.json \
    --zotero-db ~/Zotero/zotero.sqlite

# 4. Open paper.docx in Word, click Zotero > Refresh
```

## How it works

A Zotero citation in a Word document is a Word "complex field" consisting of 5 XML runs:

1. `<w:fldChar fldCharType="begin"/>` - field start
2. `<w:instrText>ADDIN ZOTERO_ITEM CSL_CITATION {...json...}</w:instrText>` - the Zotero payload
3. `<w:fldChar fldCharType="separate"/>` - separator
4. `<w:t>1</w:t>` - visible citation text (placeholder until Zotero refreshes)
5. `<w:fldChar fldCharType="end"/>` - field end

The instrText contains a JSON object with the citation ID, CSL-JSON item metadata, and a Zotero URI linking back to the user's library. The bibliography uses the same structure with `ADDIN ZOTERO_BIBL` instead.

This script constructs these XML elements using python-docx's low-level `OxmlElement` API and inserts them at each `[@key]` marker location.

## Limitations

- Citation display text is a placeholder until you click Refresh in the Zotero Word plugin
- The document must be opened in Word (not Google Docs or LibreOffice) with the Zotero plugin installed to refresh citations
- Multi-key citations like `[@key1; @key2]` are not yet supported (use separate markers)

## License

MIT
