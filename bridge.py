"""
bridge.py

NDJSON bridge between Electron frontend and zotellm backend.
Reads commands from stdin, writes events to stdout as JSON lines.

Protocol:
  Electron -> Python (stdin):
    {"type": "start", "args": {"input": "/path/file.docx", ...}}
    {"type": "resolve_response", "id": "req_1", "choice": "skip" | {"DOI":"..."} | 0}

  Python -> Electron (stdout):
    {"type": "log", "text": "..."}
    {"type": "resolve", "id": "req_1", "citation_text": "...", "candidates": [...]}
    {"type": "done", "success": true, "message": "..."}
"""

import argparse
import io
import json
import sys
import threading

from zotellm import run_zotellm


# Use a lock so JSON lines don't interleave
_write_lock = threading.Lock()
_original_stdout = sys.stdout


def _send(obj):
    """Write a JSON line to the original stdout."""
    with _write_lock:
        _original_stdout.write(json.dumps(obj) + "\n")
        _original_stdout.flush()


class _StdoutCapture(io.TextIOBase):
    """Replaces sys.stdout to capture print() calls as log events."""

    def write(self, text):
        if text and text.strip():
            _send({"type": "log", "text": text})
        return len(text) if text else 0

    def flush(self):
        pass


def _read_line():
    """Read a line from stdin (blocking)."""
    line = sys.stdin.readline()
    if not line:
        raise EOFError("stdin closed")
    return json.loads(line.strip())


def _resolve_callback(citation_text, candidates):
    """Called by run_zotellm when a citation match is uncertain.

    Sends a resolve request to Electron, blocks until the user responds.
    """
    req_id = f"req_{id(candidates)}"

    # Serialize candidates for JSON transport
    serialized = []
    for item, score in candidates:
        serialized.append({"item": item, "score": score})

    _send({
        "type": "resolve",
        "id": req_id,
        "citation_text": citation_text,
        "candidates": serialized,
    })

    # Block until Electron sends back a resolve_response
    while True:
        msg = _read_line()
        if msg.get("type") == "resolve_response" and msg.get("id") == req_id:
            choice = msg.get("choice")
            if choice == "skip":
                return None
            elif isinstance(choice, int):
                # User picked a candidate by index
                return candidates[choice][0]
            elif isinstance(choice, str):
                # DOI or PMID string
                return choice
            elif isinstance(choice, dict):
                # Direct crossref item
                return choice
            return None


def main():
    # Wait for the start command
    msg = _read_line()
    if msg.get("type") != "start":
        _send({"type": "done", "success": False, "message": f"Expected 'start', got '{msg.get('type')}'"})
        return

    gui_args = msg.get("args", {})

    # Build argparse.Namespace from the provided args
    args = argparse.Namespace(
        input=gui_args.get("input", ""),
        output=gui_args.get("output") or None,
        provider=gui_args.get("provider", "cli"),
        model=gui_args.get("model") or None,
        api_base=gui_args.get("api_base") or None,
        api_key=gui_args.get("api_key") or None,
        cli_command=gui_args.get("cli_command") or None,
        zotero_db=gui_args.get("zotero_db") or None,
        zotero_api_key=gui_args.get("zotero_api_key") or None,
        zotero_library_id=gui_args.get("zotero_library_id") or None,
        reference_doc=gui_args.get("reference_doc") or None,
        font=gui_args.get("font", "Calibri"),
        size=gui_args.get("size", 11),
        bib_heading=gui_args.get("bib_heading", "References"),
        no_crossref=gui_args.get("no_crossref", False),
        dry_run=gui_args.get("dry_run", False),
    )

    # Redirect stdout so print() calls become log events
    sys.stdout = _StdoutCapture()

    try:
        run_zotellm(args, resolve_callback=_resolve_callback)
        _send({"type": "done", "success": True, "message": "Formatting complete."})
    except Exception as e:
        _send({"type": "done", "success": False, "message": str(e)})
    finally:
        sys.stdout = _original_stdout


if __name__ == "__main__":
    main()
