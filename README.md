# zotellm

Automatic citation formatting for academic manuscripts. Takes any document (.md or .docx) with informal citations (e.g., "Smith et al., 2020"), resolves them via PubMed and CrossRef, and produces a Word document with live Zotero field codes.

## Download

**macOS (Apple Silicon):** Download [zotellm-1.0.0-arm64.dmg](https://github.com/pbilodeau94/zotellm/releases/latest) from the latest release.

After downloading, open Terminal and run:
```bash
xattr -cr /Applications/zotellm.app
```
Then open the app normally. (This removes the macOS quarantine flag since the app is not code-signed.)

**All platforms (from source):** See [Running from Source](#running-from-source) below.

## Desktop App

The desktop app provides a GUI for the full workflow:

1. Select your input file (.md or .docx)
2. Choose an LLM provider (Claude CLI works with login, no API key needed)
3. Optionally connect to your Zotero library
4. Click "Format Citations" and resolve any ambiguous matches

The app uses Claude CLI by default. If you have [Claude Code](https://docs.anthropic.com/en/docs/claude-code) installed, it works automatically with your existing login.

## Citation Format

Write citations however is natural. Include **author, journal, and year** for best results:

| Format | Match rate | Example |
|---|---|---|
| PMID | ~99% | `(PMID: 30578015)` |
| DOI | 100% | `(doi: 10.1016/S1474-4422(22)00431-8)` |
| Author, Journal, Year | ~90% | `(Cryan et al., Nature Reviews Neuroscience, 2019)` |
| Author, Year | ~75-89% | `(Cryan et al., 2019)` |
| Author only | ~50% | `(Cryan et al.)` |

Any style works. The LLM identifies citations and uses surrounding context to disambiguate. When a match is uncertain, the app shows a disambiguation dialog with journal names and scores so you can pick the correct paper.

## How It Works

1. Extracts text from your `.md` or `.docx`
2. LLM identifies all citations and extracts author/year/journal/context
3. Searches PubMed (structured fields) then CrossRef (free text) for each citation
4. Scores candidates using author, year, journal, title, and context overlap
5. Auto-selects high-confidence matches; shows disambiguation dialog for uncertain ones
6. Optionally looks up items in your Zotero library or adds new ones via the Zotero Web API
7. Replaces informal citations with Zotero field codes
8. Output: `.docx` ready to open in Word and click Zotero > Refresh

## Running from Source

### Prerequisites

- Python 3.8+
- [pandoc](https://pandoc.org/installing.html)
- One of: [Claude Code](https://docs.anthropic.com/en/docs/claude-code) (recommended), [ollama](https://ollama.com), or an API key for Anthropic/OpenAI

### Command Line

```bash
pip install python-docx requests

# Using Claude CLI (works with login, no API key)
python zotellm.py paper.md --provider cli

# Using Claude CLI with Zotero integration
python zotellm.py paper.docx --provider cli --zotero-db ~/Zotero/zotero.sqlite

# Using Anthropic API
export ANTHROPIC_API_KEY=sk-ant-...
python zotellm.py paper.md --provider anthropic

# Using OpenAI API (or any OpenAI-compatible endpoint)
export OPENAI_API_KEY=sk-...
python zotellm.py paper.md --provider openai

# Local model via OpenAI-compatible API (ollama, vLLM, LM Studio)
python zotellm.py paper.md --provider openai --api-base http://localhost:11434/v1 --model llama3
```

### Desktop App (from source)

```bash
# 1. Build the Python backend
pip install pyinstaller python-docx requests
chmod +x build_backend.sh
./build_backend.sh

# 2. Run the desktop app
cd desktop
npm install
npm start
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| `input` | Yes | Input file (`.md` or `.docx`) with informal citations |
| `--output`, `-o` | No | Output `.docx` path (default: `input_zotero.docx`) |
| `--provider`, `-p` | No | `cli` (default), `openai`, or `anthropic` |
| `--model`, `-m` | No | Model name (default depends on provider) |
| `--api-base` | No | API base URL for custom endpoints |
| `--api-key` | No | API key (overrides env var) |
| `--cli-command` | No | Custom CLI command for `--provider cli` |
| `--zotero-db` | No | Path to local `zotero.sqlite` (for key lookups) |
| `--zotero-api-key` | No | Zotero Web API key (for adding items to library) |
| `--zotero-library-id` | No | Zotero user library ID |
| `--reference-doc` | No | Pandoc reference `.docx` template (for `.md` input) |
| `--font` | No | Font for citation text (default: Calibri) |
| `--size` | No | Font size in pt (default: 11) |
| `--bib-heading` | No | Heading for bibliography location (default: References) |
| `--no-crossref` | No | Skip CrossRef/PubMed lookups |
| `--dry-run` | No | Preview without writing files |

## Getting a Zotero API Key

Only needed if you want to add new items to your Zotero library automatically:

1. Go to https://www.zotero.org/settings/keys
2. Create a new key with read/write access to your library
3. Your library ID is visible in the URL when viewing your library on zotero.org

## License

MIT
