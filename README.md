# tex-docx-roundtrip

A [Codex](https://github.com/openai/codex) skill for LaTeX ↔ Word round-trip editing.

Convert a LaTeX manuscript into an editable Microsoft Word document, make changes in Word, then synchronize those changes back into the original `.tex` source.

## What it does

1. **tex → docx**: Converts a `.tex` file (with figures, tables, headings) into a `.docx` using pandoc. PDF figures are rasterized to PNG so they appear in Word.
2. **docx → tex**: After you edit the Word file, identifies the changes and writes them back into the original `.tex` source.

## Requirements

- [pandoc](https://pandoc.org)
- `pdftoppm` from [Poppler](https://poppler.freedesktop.org) (for figure conversion)
- Python package: `python-docx`

## Installation as a Codex skill

Clone or download this repository, then add the folder to your Codex skills path, or symlink it into your Codex skills directory:

```bash
# macOS / Linux
ln -s "$(pwd)/tex-docx-roundtrip" ~/.codex/skills/tex-docx-roundtrip

# Windows (PowerShell, as Administrator)
New-Item -ItemType SymbolicLink -Path "$env:USERPROFILE\.codex\skills\tex-docx-roundtrip" `
  -Target "C:\path\to\tex-docx-roundtrip"
```

Restart Codex after installing.

## Usage

See [`references/usage.md`](references/usage.md) for detailed examples.

Quick start:

```bash
# 1. Generate editable Word document
python scripts/tex2docx.py manuscript.tex manuscript.docx

# 2. Edit manuscript.docx in Microsoft Word

# 3. Compare original and edited Word files
python scripts/compare_docx.py manuscript.docx manuscript_edited.docx changes.json

# 4. Apply safe text changes back to .tex
python scripts/apply_text_changes.py manuscript.tex manuscript.docx manuscript_edited.docx manuscript_updated.tex
```

## Important limitations

- Math and macros are fragile. Pandoc renders LaTeX math as Word text; restore `$...$` and macros manually when converting back.
- Figures are embedded as images in Word. Do not resize or move them if you want the original `.tex` figure code to remain unchanged.
- Tables may lose formatting. Simple text edits inside cells can be synced, but layout changes should be done in LaTeX.
- References and citations are usually turned into plain text by pandoc. Keep `\cite{key}` in the `.tex` instead of copying the rendered text back literally.

## License

[MIT](LICENSE)
