#!/usr/bin/env python3
"""READ-ONLY first-pass sidebar dependency inventory for an Apps Script project.

Prints findings to terminal. Does not change any source, spreadsheet, live
deployment, or repository. Reports candidates for manual review; does not
claim that static text matching proves a UI works or is broken.
"""
import argparse
import re
from pathlib import Path

FUNCTION_DECL = re.compile(r"\bfunction\s+([A-Za-z_$][\w$]*)\s*\(")
ASSIGNMENT_DECL = re.compile(r"\b(?:const|let|var)\s+([A-Za-z_$][\w$]*)\s*=\s*(?:async\s*)?(?:function\s*\(|(?:\([^)]*\)|[A-Za-z_$][\w$]*)\s*=>)")
HANDLER = re.compile(r"\bon(?:click|change|submit|input)\s*=\s*([\"'])(.*?)\1", re.I | re.S)
HANDLER_FUNC = re.compile(r"^\s*(?:return\s+)?([A-Za-z_$][\w$]*)\s*\(")
SERVER_CALL_LINE = re.compile(r"^\s*\.\s*([A-Za-z_$][\w$]*)\s*\(", re.M)
SERVER_DIRECT = re.compile(r"\bgoogle\s*\.\s*script\s*\.\s*run\s*\.\s*([A-Za-z_$][\w$]*)\s*\(")
IGNORED = {"withSuccessHandler", "withFailureHandler", "withUserObject"}
TARGETS = (
    "ce_Sidebar_W6.html", "ce_Sidebar_W2C.html",
    "ce_Sidebar_FI.html", "glb_BatchSidebar.html",
    "glc_ComparisonSidebar.html", "glc_ArticleCompareSidebar.html",
    "glc_Sol20_Progress.html", "glse_Progress.html"
)

def all_defs(content):
    return set(FUNCTION_DECL.findall(content)) | set(ASSIGNMENT_DECL.findall(content))

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True, type=Path)
    args = ap.parse_args()
    root = args.root.resolve()
    if not root.is_dir():
        raise SystemExit("Missing source folder: " + str(root))
    scripts = sorted([*root.glob("*.js"), *root.glob("*.gs")])
    server_defs = set()
    for p in scripts:
        server_defs.update(all_defs(p.read_text(encoding="utf-8-sig", errors="replace")))
    # Includes must be in the entrypoint template, not merely somewhere on disk.
    entry = root / "ce_Sidebar_Main.html"
    entry_text = entry.read_text(encoding="utf-8-sig", errors="replace") if entry.is_file() else ""
    def included(name):
        # Google Apps Script template include expression; tolerate quote styles.
        return ("include('" + name + "')" in entry_text or
                'include("' + name + '")' in entry_text)
    html_files = sorted(root.glob("*.html"))
    shared_html_defs = set()
    for shared in html_files:
        if included(shared.stem):
            shared_html_defs.update(all_defs(shared.read_text(encoding="utf-8-sig", errors="replace")))
    summaries = []
    for filename in TARGETS:
        p = root / filename
        if not p.exists():
            print("NOT PRESENT:", filename, "(review whether this sidebar is expected)")
            continue
        if filename.startswith("ce_Sidebar_") and not included(p.stem):
            print("  REVIEW: Sidebar not included in main entrypoint:", filename)
        content = p.read_text(encoding="utf-8-sig", errors="replace")
        local_defs = all_defs(content)
        inline_handlers = []
        unresolved_handlers = []
        shared_handlers = []
        for match in HANDLER.finditer(content):
            expression = match.group(2)
            called = HANDLER_FUNC.search(expression)
            if called:
                name = called.group(1)
                inline_handlers.append(name)
                if name not in local_defs:
                    if name in shared_html_defs:
                        shared_handlers.append(name)
                    else:
                        unresolved_handlers.append(name)
        # Server functions typically appear on a continuation line after
        # google.script.run and optional withSuccessHandler()/FailureHandler().
        # Identify calls with a known server-style prefix, rather than treating
        # every JS method chain as an Apps Script RPC. This is incomplete by
        # design: ambiguous calls require human review.
        server_candidates = set(SERVER_DIRECT.findall(content)) - IGNORED
        for m in SERVER_CALL_LINE.finditer(content):
            name = m.group(1)
            if name in IGNORED:
                continue
            if re.match(r"^(?:ce_|glb_|glc_|glse_|w2c|bc_|[A-Z]{2,}_)", name):
                server_candidates.add(name)
        server_candidates = sorted(server_candidates)
        missing_servers = [x for x in server_candidates if x not in server_defs]
        summaries.append((filename, len(inline_handlers), len(server_candidates), unresolved_handlers, missing_servers, shared_handlers))
        print("\n" + filename)
        print("  Inline HTML actions:", len(inline_handlers), "| server-looking calls:", len(server_candidates))
        for name in sorted(set(shared_handlers)):
            print("  SHARED: HTML action is defined in another HTML file (inclusion not verified):", name)
        for name in sorted(set(unresolved_handlers)):
            print("  REVIEW: HTML action has no matching function in this file:", name)
        for name in missing_servers:
            print("  REVIEW: called server-looking function not found in JS/GS:", name)
        if not unresolved_handlers and not missing_servers:
            print("  No missing matches found in this preliminary static scan.")
    print("\nTOTAL: sidebars inspected:", len(summaries))
    print("Inline actions reviewed:", sum(x[1] for x in summaries))
    print("Server-looking calls reviewed:", sum(x[2] for x in summaries))
    print("Unmatched inline action references:", sum(len(x[3]) for x in summaries))
    print("Unmatched server candidates:", sum(len(x[4]) for x in summaries))
    print("Shared HTML action references:", sum(len(x[5]) for x in summaries))
    print("\nNOTE: No changes made. These are preliminary static checks.")
    print("Dynamic calls, shared HTML includes, aliases, and event delegation may")
    print("need manual inspection; passing does not prove buttons work.")

if __name__ == "__main__":
    main()
