# Usage Examples

Replace `<skill-root>` with the path to this skill folder, and adjust the input/output paths to match your project.

## Example 1: Convert a LaTeX manuscript to Word

```powershell
python <skill-root>/scripts/tex2docx.py `
  "C:\path\to\main.tex" `
  "C:\path\to\main_editable.docx" `
  --pandoc "C:\path\to\pandoc.exe"
```

If pandoc is on PATH, omit `--pandoc`.

## Example 2: Compare original and edited Word files

```powershell
python <skill-root>/scripts/compare_docx.py `
  "C:\path\to\main_editable.docx" `
  "C:\path\to\main_edited.docx" `
  "C:\path\to\changes.json"
```

## Example 3: Apply safe text changes automatically

```powershell
python <skill-root>/scripts/apply_text_changes.py `
  "C:\path\to\main.tex" `
  "C:\path\to\main_editable.docx" `
  "C:\path\to\main_edited.docx" `
  "C:\path\to\main_updated.tex"
```

A backup `main.tex.bak` is created next to the original `.tex`.

## Example 4: Manual patch for math-heavy paragraphs

If the automatic script skips the abstract because it contains `$T_{\\mathrm{e}}$`, read the diff entry and apply a patch like:

```diff
- while plasma temperature ($T_{\\mathrm{e}}$) and electron density ($N_{\\mathrm{e}}$) are derived from baseline-corrected spectra.
+ while plasma temperature ($T_{\\mathrm{e}}$) and electron density ($N_{\\mathrm{e}}$) are extracted from baseline-corrected spectra.
```

Always recompile after applying changes.
