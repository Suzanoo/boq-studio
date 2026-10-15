# S0 persistence safety

S0 retains the existing local version-1 session JSON. It does not introduce a portable Project format. Keep the original Excel workbook unchanged and at its original path to resume normally. Exported Excel remains an independent deliverable, not a project backup.

## Everyday workflow

1. Open the source workbook. If a matching saved session exists, choose Yes to resume it. Cancel leaves the currently open work intact. No starts fresh only after preserving the previous session in a recovery file.
2. Configure mappings and edit the Draft Tree. Saving an unchanged mapping preserves the reviewed tree. A changed mapping requires confirmation before removing that sheet's tree and approval; the old mapping/tree are archived first.
3. Save Draft saves an unapproved tree. Approve saves the exact approved candidate. Closing the editor and choosing Discard abandons its changes, including changes following Undo/Redo.
4. Save Project writes the local session. Selection/mapping/tree commits also save it. A failed save is reported; failed tree commits keep the editor open. Close/Open cannot abandon dirty work after a failed save unless a separate recovery JSON has been successfully written.
5. Reopen the same original workbook and choose Resume. Export requires an actual approved tree and a successful session save. While export runs, changing work or closing is blocked.

If Excel changes the original workbook, session saves and new exports are blocked rather than associating old row decisions with different data. Restore the exact original file to resume it. Materially changed BOQs require separate work; merging is not part of S0.

## Recovery files

Session locations:

- Windows: `%LOCALAPPDATA%/BOQ Studio/sessions`
- macOS: `~/Library/Application Support/BOQ Studio/sessions`
- Linux: `$XDG_DATA_HOME/boq-studio/sessions`, default `~/.local/share/boq-studio/sessions`

The primary filename is the SHA-256 of the resolved original workbook path, followed by `.json`. A `.json.bak` holds the previous valid changed save. Identical saves do not rotate it. Explicit start-fresh/mapping-replacement decisions create unique `.recovery-<UTC>-<UUID>.json` archives which later saves do not rotate. A `.json.lock` is an empty/one-byte coordination file, not project data; leave it alone while any app instance is active.

When a primary session is invalid or missing, Open offers a matching valid backup when available. It never silently regenerates an invalid saved tree. If saving is blocked, the recovery-file dialog lets you select a NEW JSON filename elsewhere; it refuses to overwrite existing recovery files. Recovery JSON contains mappings, ordered trees, overrides and workflow flags, but NOT the original workbook.

Manual recovery (Coordinator-assisted; no new Project Open interface in S0):

1. Close all BOQ Studio instances. Preserve the original workbook and all current session/backup files separately. Restore the original workbook byte-for-byte if it was changed.
2. Using the application's Python environment, obtain the target primary path (replace the sample path):

   `python -c "from pathlib import Path; from boq_studio.session import session_path_for; print(session_path_for(Path('C:/BOQ/source.xlsx')))"`

3. Verify the chosen recovery JSON with `load_session(Path(...))` and verify its `source.sha256` equals `workbook_fingerprint(Path(...))` for the original workbook. Do not edit its fingerprint to bypass this check. The functions are in `boq_studio.session`.
4. Preserve any existing primary under a new filename, then COPY the verified recovery JSON to the exact primary path from step 2. Keep the recovery original. Do not use the Excel Open dialog to open JSON.
5. Open the original workbook, choose Resume, check sheet mappings, tree ordering and approval, then Save Project.

Archives are local recovery aids, not protection against disk loss. Retain independent copies of the source and recovery data. Do not delete archives until their work is no longer needed.

## Desktop acceptance gate (Windows and macOS)

Use disposable copies of real representative BOQs. Run both supported application entry points. Record OS, Python/Tk versions, file sizes, and results.

1. Configure at least two sheets, classification/level overrides, and a tree with renamed, moved, inserted and reordered nodes. Save Draft, close app, reopen original and Resume: all decisions/order persist; draft remains unapproved.
2. Approve; restart and Resume. Approval matches the exact saved tree. Edit and Save Draft: approval/exported status clears. Reapprove and verify persistence.
3. Edit, Undo/Redo, then Discard: parent tree, disk session and previous approval remain unchanged. Save unchanged mapping: no tree reset. Change mapping: No preserves everything; Yes archives old mapping/tree before clearing that sheet's approval/tree. Restore archive using the guide and verify old decisions.
4. Cancel Resume; try unreadable/non-XLSX input; fail archive creation using a disposable unwritable session directory. Current project remains usable. Choose No at Resume when writable: archive remains readable after several later saves.
5. Corrupt a copied primary JSON with a valid backup present. Open offers backup; accept, close/reopen and verify recovered tree. Try unknown-version JSON or no usable backup: no silent restoration/regeneration or overwrite.
6. Make session writes fail (permissions/controlled disk-full fixture). Save reports failure; tree editor stays open with candidate intact. Close/switch stays blocked if recovery is declined, canceled or fails. Save recovery to a writable NEW filename and verify it; only then may close/switch proceed. Restore it and compare the candidate tree. Existing recovery filenames must be refused.
7. Open same original in two app instances. Save from first; second's stale save must report conflict without overwriting first. Preserve second's work with separate recovery. Close processes abruptly and verify locks release.
8. Modify source in Excel after opening, both before and during tree editing. Save/export must not bind old decisions to the new bytes. Candidate recovery remains possible. Restore exact original and retry. Change source during export: either reject before output or export the verified snapshot, retaining original input path in LOG.
9. During export try Open, Save, selection change, mapping/tree editing and window close: all blocked. Successful results alone become exported; warnings remain ready. Inject session-save failure after successful output: Excel remains available, status explicitly says session not saved.
10. Export representative pre-S0 and S0 approved trees; compare sheets, 11-column contract, formulas, Project/SUMMARY/LOG content, totals and node order. Open/save/close/reopen outputs in Excel and validate recalculation. No new project metadata should appear.
11. Test Unicode/Thai/spaced/long supported paths; local fixed disk and realistic large BOQ. Test read-only preferences file: warning must not prevent tree Save/Discard. Record save responsiveness and snapshot disk use.

Headless regression tests cover real files and application handlers, but do not substitute for these visual/OS checks. Unsaved edits are not periodically journaled; an unexpected process/OS failure can still lose edits made since the last successful save.
