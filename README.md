# zotero-field-insert

Insert Zotero-compatible field codes into Word (.docx) documents programmatically.

## Problem

When you generate Word documents from Markdown (via pandoc), LaTeX, or Python, there is no way to include live Zotero citations. Pandoc's `--citeproc` produces plain text references that the Zotero Word plugin cannot recognize or manage. This tool bridges that gap.

## What it does

1. Finds `[@citationkey]` markers in a `.docx` file
2. Replaces each with a proper `ADDIN ZOTERO_ITEM CSL_CITATION` Word field code containing full CSL-JSON metadata
3. Inserts an `ADDIN ZOTERO_BIBL` field where the bibliography should appear
4. Optionally links citations back to your local Zotero library using item keys and your Zotero SQLite database

After running the script, open the document in Word and click **Refresh** in the Zotero tab. Zotero will renumber citations, apply your chosen citation style, and generate the bibliography.

## Installation

```bash
pip install python-docx
```

No other dependencies. The script reads the Zotero SQLite database directly (read-only) if you want to link citations to your library.

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

## Workflow with pandoc

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
