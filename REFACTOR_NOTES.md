# Stage 1: Recovery & Deduplication — Findings

This covers only the mechanical recovery step (`stage1_recovered.py`), not the
modular refactor. Full modularization is paused pending review of this stage,
per request.

## 1. The original file does not run

`label_software_2.0.py` (13,997 lines) does not parse as valid Python.
Verified with `ast.parse`:

```
SyntaxError at line 638: unmatched ')'
    root.mainloop(), '').replace(' ', '').replace(',', '').replace('\xa0', '').strip()
```

This isn't a cosmetic duplication problem — the file cannot be imported or
run in its current on-disk form. It fails before `crear_widgets()` or
`mainloop()` is ever reached.

## 2. What the "×243 duplication" actually is

`cargar_ejemplos` and `cambiar_idioma` are indeed repeated 243 times each,
confirmed byte-identical for `cargar_ejemplos`. But `cambiar_idioma`'s 243
copies are **not** clean repeats of the method — each one is immediately
followed by a spliced-in, broken fragment:

```python
    def cambiar_idioma(self):
        # (Mantén la función que ya tienes para cambiar idioma)
        pass

# ========== MAIN ==========
if __name__ == "__main__":
    if not validar_licencia():
        sys.exit()

    root = tk.Tk()
    root.geometry("1300x800")
    app = EtiquetadoraApp(root)
    root.mainloop(), '').replace(' ', '').replace(',', '').replace('\xa0', '').strip()
                if precio_clean == '':
                ...                                    # <- dangling continuation
                                                         #    of pegar_desde_excel,
                                                         #    then straight into the
                                                         #    next `def cargar_ejemplos`
```

So this isn't 243 clean copies of two methods — it's a corrupted template
(a stray `if __name__ == "__main__":` block glued mid-class to a fractured
tail of `pegar_desde_excel`) reinserted 242 times, with one clean copy of the
real `if __name__` block surviving at true EOF. Whatever process generated
this (bad find/replace, bad merge, or a broken automated edit) did not
produce valid duplicates — it produced noise that happens to *look* like
duplication at a glance.

**Net effect: none of the 13,369 lines after the first clean copy (line
628) contain any additional real logic.** They were deleted, not "merged."

## 3. Two real (pre-existing) bugs were hiding in the "unique" 628 lines

These are not related to the duplication — they were already broken in the
one legitimate copy of the code, and are part of why the file doesn't parse:

- **`actualizar_resumen`** was defined nested *inside* `crear_widgets`
  (indented 8 spaces instead of 4), with its body indented no deeper than
  its own `def` line — a plain `IndentationError`. Because it was a local
  closure, not a method, `self.actualizar_resumen()` (called from
  `__init__` and `actualizar_tabla`) would never have resolved even if the
  indentation were merely cosmetic.
- **`pegar_desde_excel`** had its `def` line indented 20 spaces instead of
  4 — also invalid.

Both are fixed in `stage1_recovered.py` by dedenting to proper method level.
No logic inside either method was changed.

## 4. Eleven methods referenced by the UI do not exist anywhere in the file

`crear_widgets` wires up buttons with `command=self.agregar`,
`self.probar_offset`, `self.eliminar`, `self.duplicar`,
`self.imprimir_chequeados`, `self.imprimir_todos`, `self.detener_impresion`,
`self.limpiar`, `self.cargar`, `self.guardar`, `self.on_click_check`,
`self.editar_celda`. None of these are defined anywhere in the source —
confirmed by searching every indentation level and every occurrence of each
name in the file (they only appear in `TEXTOS` strings and in the
`command=` wiring itself). There is no inheritance, no dynamic
`setattr`/`exec`, and no second source file that could be supplying them.

Checked for a better source before assuming this: `LabelSoftware.zip` /
`LabelSoftware.exe` is a compiled PyInstaller build (2026-06-06), and its
README points at `github.com/superyeyo8482/label-software-zd220` — but that
repo contains only the compiled `.exe`, README, and manual PDF, not source.

**Per your instruction, these are left as explicit stubs, not invented.**
Each shows a "not implemented" message box instead of guessing real
business logic (CSV format, print-selected semantics, etc.). They're marked
`RECOVERED-STUB` in `stage1_recovered.py`. The app now launches and every
piece of logic that *does* exist (ZPL generation, licensing, Excel paste
parsing, batch summary, table refresh) is reachable and testable — but
Agregar/Eliminar/Duplicar/Imprimir/Limpiar/Cargar CSV/Guardar CSV/edit-cell/
click-to-check all need real implementations written from scratch before
this is production-usable again.

## 5. Confirmed live: the license key vulnerability (issue #2)

Loaded `CLAVE_MAESTRA` with `cryptography.fernet.Fernet` directly: it's a
valid, currently-active symmetric key. With only the source file, I
generated a valid activation token for an arbitrary fake machine ID and
`validar_licencia_ingresada` accepted it. This is a live bypass, not a
theoretical one — anyone with a copy of the `.py` file (or `strings` on the
compiled `.exe`) can activate the software on any machine without ever
contacting whoever issues licenses. Flagged, not fixed — see the comment
in `stage1_recovered.py` above `CLAVE_MAESTRA`.

## 6. Size impact

| | Lines |
|---|---|
| Original (`label_software_2.0.py`) | 13,997 |
| Recovered, deduplicated, valid (`stage1_recovered.py`) | 740 (incl. new comments + 12 stub methods) |
| Estimated real original logic (excl. added comments/stubs) | ~600 |

This matches your own estimate of "600-700 lines of real logic" — the
mechanism was just corruption, not clean duplication.

## Stage 2: modular split (done)

`stage1_recovered.py` has been split into `config.py`, `i18n.py`,
`licensing.py`, `printing.py`, `models.py`, `excel_paste.py`, `gui.py`,
`app.py`, plus `tests/` (26 tests, all passing — run with
`python -m pytest tests/ -v` from inside `label_software/`). See
`README.md` for what each module owns.

Beyond splitting files, three substantive changes were made in this stage,
all preserving external behavior:

1. **Error handling (issue #3).** Bare `except:` / silent `except
   Exception: pass` blocks were replaced with narrowed exception types
   (`OSError`, `json.JSONDecodeError`, `ValueError`, `InvalidToken`, etc.)
   plus `logging` calls, so failures are recorded instead of vanishing.
   `printing.py` deliberately keeps one broad `except Exception` — at the
   win32print call boundary, where the underlying library raises
   `pywintypes.error`, which isn't a stable cross-platform importable
   type — but it's now logged and re-raised as a typed `PrintError`
   rather than swallowed. Messageboxes were moved out of `printing.py`
   entirely; that module has no tkinter dependency now, and `gui.py`
   decides when a failure is user-facing.
2. **`models.calcular_resumen`** consolidates a calculation that existed
   twice in the original with two different (disagreeing) error-handling
   behaviors — `actualizar_resumen` skipped products with a bad `precio`,
   `pegar_desde_excel`'s inline copy had no such guard and could raise
   `ValueError` on the same data. Both call sites now share the safer
   (skip-on-error) version. This removes a latent crash bug; it does not
   change output for any valid data.
3. **`printing.py`'s `WindowsZplBackend`/`PosixFileBackend` split**
   isolates the only Windows-only dependency (`win32print`) behind an
   interface, per your requirement — every other module, including
   `gui.py`, imports and runs fine on macOS/Linux with no pywin32
   installed. Confirmed: the full test suite and a live GUI smoke-launch
   (constructing `EtiquetadoraApp`, verifying the product table populates,
   closing cleanly) both ran on this machine (macOS) with `win32print`
   stubbed out.

Everything else — ZPL template text, TEXTOS strings, Excel-paste field
cleaning rules, license file naming, widget layout/labels — is unchanged
from the recovered stage 1 source.

## Stage 3: a second, complete source turned up (`label_software_v2_backup.py`)

After stage 2 shipped, `label_software_v2_backup.py` (987 lines) appeared
in the same directory. Checked before touching anything:

```
ast.parse(...)  -> PARSES OK
```

It contains every method that stage 1/2 had to stub as "not implemented":
`agregar`, `eliminar`, `duplicar`, `limpiar`, `cargar`, `guardar`,
`imprimir_chequeados`, `imprimir_todos`, `probar_offset`, `editar_celda`,
`on_click_check`, `detener_impresion`, plus threaded printing
(`_imprimir`/`_impresion_hilo`) - and `cambiar_idioma` has a real,
complete implementation instead of the `pass` stub found in
`label_software_2.0.py`. Its header comment identifies it as the
corrected version: *"Versión 2.0 - Corregida y funcional. Solo Proveedor
es obligatorio. ID se genera automáticamente si está vacío."*

**Conclusion: `label_software_2.0.py` was an incomplete, corrupted attempt
at a further edit on top of what this backup already had working** - not
an earlier draft that this backup extended. The stub methods filled in
during stage 1/2 were reconstructed guesses at behavior that, it turns
out, already existed. All of it has now been replaced with the real
implementations from this file, ported module-by-module (`gui.py`
primarily; `csv_io.py` is new, holding CSV row conversion that didn't
exist before).

Two concrete behavioral differences between the two source files, adopted
from this backup as the authoritative one:

- **`LS_VAL`/`LT_VAL` defaults are `50`/`53`**, not `-50`/`-43`. Updated
  in `config.py`.
- **TEXTOS wording differs** in a few keys (`gramos` label now says
  "(3 dígitos)"/"(3 digits)"; `campo_obligatorio` now says only Proveedor
  is required, not Proveedor+ID; batch-summary label wording). Updated in
  `i18n.py`.

`stage1_recovered.py` is left as-is (not rewritten) - it's a historical
record of what recovery from the corrupted file looked like, not a
current source of truth. Don't build on it; build on this file's modules.

## Stage 4: the two versions of `pegar_desde_excel` disagreed - reconciled, not picked

This was the one function that meaningfully differs between the two
source files, not just "one has it and one doesn't." They disagree on two
independent things:

1. **What's required.** The backup only requires `proveedor` (missing
   `id` triggers a "generate IDs automatically?" prompt, matching its own
   documented design and how `agregar()` already treats optional fields).
   The corrupted file's fragment requires both `proveedor` *and* `id`,
   silently dropping rows missing either - which contradicts the backup's
   own stated design.
2. **How defensively each field is cleaned.** The corrupted file's
   fragment detects tab *or* comma separators and pads short rows instead
   of rejecting them; strips unit words from gramos ("15gr", "15 g");
   and cleans prices with non-breaking-space handling plus a
   comma-as-decimal fallback pass. The backup's version is simpler on all
   three counts (tab-only, hard-rejects rows under 7 columns, no unit-word
   or non-breaking-space handling).

Adopted the backup's *requirements* (it's the deliberate, documented,
confirmed-working design) combined with the corrupted fragment's *field
cleaning* (pure parsing robustness, no conflict with the requirements
question). See `excel_paste.py`'s module docstring for the same
breakdown in-code. `tests/test_excel_paste.py` covers both halves
explicitly (e.g. `test_missing_id_is_kept_not_skipped...` for the
requirements side, `test_price_cleaning_strips_non_breaking_space` /
`test_gramos_cleaning_strips_unit_words` for the cleaning side) so this
decision doesn't silently drift later.

If this call is wrong for your real supplier spreadsheets, the two halves
are independent - e.g. reverting to "both proveedor and id required" is a
one-line change to the check in `parsear_filas_pegadas`, without touching
the cleaning functions.
