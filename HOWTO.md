# How to insert Zotero citations into pandoc-generated Word documents

## The problem

Pandoc's `--citeproc` generates plain text citations. The Zotero Word plugin cannot recognize or manage these. If you want live Zotero citations in a pandoc-generated document, you need to insert Zotero field codes after pandoc generates the .docx.

## Tool

**zotero-field-insert**: https://github.com/pbilodeau94/zotero-field-insert

## Step-by-step workflow

### 1. Prepare your bibliography

Export references from Zotero as CSL JSON, or build a `references.json` from the Zotero SQLite database:

```python
import sqlite3, json

db = sqlite3.connect("/Users/philbilodeau/Zotero/zotero.sqlite")

# Find an item by title
rows = db.execute("""
    SELECT i.key, idv.value
    FROM items i
    JOIN itemData id ON i.itemID = id.itemID
    JOIN itemDataValues idv ON id.valueID = idv.valueID
    JOIN fields f ON id.fieldID = f.fieldID
    WHERE f.fieldName = 'title' AND idv.value LIKE '%keyword%'
""").fetchall()

# Get full metadata for an item
def get_item(item_key):
    item_id = db.execute("SELECT itemID FROM items WHERE key = ?", (item_key,)).fetchone()[0]
    fields = {}
    for r in db.execute("""
        SELECT f.fieldName, idv.value FROM itemData id
        JOIN fields f ON id.fieldID = f.fieldID
        JOIN itemDataValues idv ON id.valueID = idv.valueID
        WHERE id.itemID = ?
    """, (item_id,)):
        fields[r[0]] = r[1]
    creators = []
    for r in db.execute("""
        SELECT cr.firstName, cr.lastName FROM itemCreators ic
        JOIN creators cr ON ic.creatorID = cr.creatorID
        WHERE ic.itemID = ? ORDER BY ic.orderIndex
    """, (item_id,)):
        creators.append({"family": r[1], "given": r[0]})
    return fields, creators
```

Build the CSL JSON array and save as `references.json`.

### 2. Create a keymap

Map your citation keys to Zotero item keys (the 8-character alphanumeric key visible in the Zotero URL bar):

```json
{
    "banwell2023": "BBJKWN9G",
    "sattarnezhad2018": null
}
```

Save as `keymap.json`. Set to `null` for items not in your library.

### 3. Write markdown with citation markers

Use `[@citationkey]` syntax in your markdown. Do NOT use pandoc's `--citeproc` flag.

```markdown
MOGAD is diagnosed using the 2023 international criteria.[@banwell2023]
```

### 4. Convert with pandoc (no citeproc)

```bash
pandoc paper.md -o paper.docx --reference-doc=template.docx
```

The `[@key]` markers will appear as literal text in the .docx.

### 5. Insert Zotero field codes

```bash
python zotero_field_insert.py paper.docx \
    --bib references.json \
    --keymap keymap.json \
    --zotero-db ~/Zotero/zotero.sqlite
```

### 6. Open in Word and refresh

Open the document in Word. Go to the Zotero tab and click **Refresh**. Zotero will:
- Renumber all citations
- Apply your citation style (e.g., AMA, Vancouver)
- Generate the bibliography

You can now add, edit, or remove citations using the Zotero plugin as usual.

## Zotero field code structure (for reference)

A Zotero citation in Word XML is 5 consecutive `<w:r>` elements:

```
Run 1: <w:fldChar fldCharType="begin"/>
Run 2: <w:instrText> ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"abc123","properties":{...},"citationItems":[{"id":996,"uris":["http://zotero.org/users/USERID/items/ITEMKEY"],"itemData":{...CSL-JSON...}}],"schema":"..."}</w:instrText>
Run 3: <w:fldChar fldCharType="separate"/>
Run 4: <w:t>1</w:t>  (display text, placeholder until refresh)
Run 5: <w:fldChar fldCharType="end"/>
```

The bibliography uses the same structure with:
```
ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY
```

## Key details

- The `instrText` must have `xml:space="preserve"` and start with a space
- Each citation item contains full CSL-JSON metadata embedded in the document
- The `citationID` is a random 8-character alphanumeric string
- Zotero URIs follow the format `http://zotero.org/users/{userID}/items/{itemKey}`
- For superscript numbered styles, apply `<w:vertAlign w:val="superscript"/>` to all runs
- The Zotero user ID can be found in the SQLite database: `SELECT value FROM settings WHERE setting='account' AND key='userID'`
