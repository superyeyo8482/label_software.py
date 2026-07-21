# Label Software (refactored)

Tkinter GUI for printing Zebra ZPL jewelry price labels: manual product
entry, paste-from-Excel import, batch summary, per-machine license
activation. This is a modular rewrite of `label_software_2.0.py`
(the original file, kept untouched one directory up as reference).

**Read `REFACTOR_NOTES.md` first.** `label_software_2.0.py` (the file this
refactor started from) did not run at all — it failed to even parse as
Python, and 11 methods the UI wires up to buttons didn't exist anywhere in
it. A second file, `label_software_v2_backup.py`, later turned up and
turned out to be the complete, working version those methods were missing
from — see REFACTOR_NOTES.md "Stage 3" and "Stage 4" for exactly what was
recovered vs. what came from that file vs. the one place (Excel-paste
parsing) where the two disagreed and had to be reconciled.

## Structure

| File | Responsibility |
|---|---|
| `config.py` | Printer name, print offsets, file paths, ZPL template, logging setup. No secrets. |
| `i18n.py` | `TEXTOS` translation dict (es/en) + `t(idioma, key)` lookup. |
| `licensing.py` | Machine ID, activation code validation, license file persistence. **Contains an explicit, unfixed security warning — read its docstring.** |
| `printing.py` | ZPL rendering (`generar_zpl`), printer output behind a `PrinterBackend` interface, and `imprimir_lote` (threaded-batch-print orchestration, STOP-aware) — all testable without Windows/pywin32/a real printer. |
| `models.py` | Product dicts, ID generation, batch summary math. |
| `excel_paste.py` | Pure parsing of pasted Excel/Sheets clipboard text into product dicts — no clipboard/tkinter access, so it's independently testable. |
| `csv_io.py` | Pure CSV row ↔ product dict conversion for Cargar/Guardar CSV — no file I/O, so it's independently testable. |
| `gui.py` | `EtiquetadoraApp` (tkinter) and the activation dialog. Delegates everything above; this file should only contain widget wiring. |
| `app.py` | Entry point: logging setup, license gate, `mainloop()`. |
| `tests/` | Unit tests for everything above that doesn't need a display (see below). |
| `stage1_recovered.py` | Intermediate single-file recovery step from before `label_software_v2_backup.py` was found — historical record only, not used by `app.py`. See REFACTOR_NOTES.md "Stage 3". |

## Running

```
cd label_software
pip install -r requirements.txt
python app.py
```

On first run on a new machine it will show the activation dialog (same
flow as the original — copy the device ID, get a code from whoever issues
them, paste it in).

## Testing

```
cd label_software
pip install -r requirements.txt
python -m pytest tests/ -v
```

37 tests, all passing, covering ZPL generation, batch-print orchestration
(STOP mid-batch, first-failure aborts the batch), license code validation
(including a test that demonstrates the license-key vulnerability itself,
so it can't silently regress into something worse), batch summary math,
CSV row conversion, and Excel-paste parsing edge cases (missing columns,
price/gram cleanup, missing-provider rows, missing-id handling). None of
this requires a display or Windows — `printing.py`'s Windows-only code
path is behind `PrinterBackend` and is never exercised by the test suite.

Also manually smoke-tested end to end on macOS with `win32print` stubbed
out: launched the real GUI, added a product through the form (auto-ID
generation), duplicated a row, toggled the language button (verified the
window title and labels actually change), and generated ZPL from live
app state.

## Known limitations (carried over from the source, not introduced here)

- **License key is insecure by construction.** See `licensing.py`'s module
  docstring and `REFACTOR_NOTES.md` §5. Not fixed — flagged for you to
  decide on (asymmetric signing vs. server-side check).
- **Excel-paste requirements and field-cleaning were merged from two
  disagreeing source files** (only `proveedor` required, from the
  confirmed-working backup; comma-separator/unit-word/non-breaking-space
  handling, from the other file's more defensive fragment). See
  `REFACTOR_NOTES.md` "Stage 4" if your real supplier data suggests a
  different call should've been made here — the two halves are
  independent and easy to change separately.
- **The window title is only ever set inside `cambiar_idioma`.** Preserved
  as-is from the source rather than "fixed" by also setting it at
  startup — it's a real, if odd, existing behavior (the window keeps its
  default Tk title until the language button is clicked once).
