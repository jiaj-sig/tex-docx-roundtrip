---
name: tex-docx-roundtrip
description: Convert a LaTeX manuscript to an editable Word document and later apply the Word edits back to the LaTeX source. Use when the user wants to revise a .tex file through Microsoft Word and keep the two in sync.
---

# LaTeX ↔ Word Round-Trip Editing

This skill helps a user edit a LaTeX manuscript in Microsoft Word and then synchronize those edits back into the `.tex` source. It is useful when the user finds Word more convenient for wording, section reordering, or collaborative comments, but ultimately needs a compiled PDF from LaTeX.

## What it does

1. **tex → docx**: Converts a `.tex` file (with figures, tables, headings) into a `.docx` using pandoc. PDF figures are rasterized to PNG so they appear in Word.
2. **docx → tex**: After the user edits the Word file, identifies the changes and writes them back into the original `.tex` source.

## Requirements

- **pandoc** must be available. The helper script searches PATH and common install locations. If pandoc is missing, install it from <https://pandoc.org> or download a portable build and pass `--pandoc /path/to/pandoc.exe`.
- **pdftoppm** (from Poppler) is used to convert PDF figures to PNG. It is usually available in the Codex workspace; if not, figures will be omitted from the Word output.
- Python packages: `python-docx` (already present in the Codex workspace).

## Workflow

### Step 1: Generate the editable Word document

Run the helper script from the skill:

```bash
python <skill-root>/scripts/tex2docx.py < manuscript.tex > [manuscript.docx] [--pandoc /path/to/pandoc.exe]
```

If no output name is given, the `.docx` is placed next to the `.tex` with the same stem.

Then show/open the generated `.docx` for the user and tell them to edit it and return the modified file path.

### Step 2: Wait for the user to provide the edited Word file

Pause after generating the docx. The user must manually edit the Word document and tell you where the edited file is.

### Step 3: Identify changes

Run the comparison helper to produce a structured diff:

```bash
python <skill-root>/scripts/compare_docx.py <original.docx> <edited.docx> [changes.json]
```

Read `changes.json` (or the script output). Each entry contains:

- `index`: paragraph position
- `style`: Word paragraph style (Heading 1, Body Text, Abstract, …)
- `old`: original text
- `new`: edited text
- `kind`: changed / added / removed

### Step 4: Apply changes back to the .tex file

Use the safest method that works for the change:

1. **Automatic exact replacement** (safe for plain-text paragraphs without math):  
   ```bash
   python <skill-root>/scripts/apply_text_changes.py <original.tex> <original.docx> <edited.docx> [output.tex]
   ```
   The script only replaces text that appears exactly once in the `.tex` file and skips headings/captions. Review its report.

2. **Manual patch** (required for paragraphs containing math, citations, or complex LaTeX):  
   Use the diff report to locate the corresponding location in the `.tex` file and apply a precise patch with `apply_patch`. Preserve LaTeX commands, math mode `$...$`, citations `\cite{}`, and references `\ref{}`.

3. **Structural changes** (added/removed/moved sections):  
   Edit the `.tex` directly to insert or remove `\section{}`, `\subsection{}`, figure/table environments, etc.

Always keep a backup of the original `.tex`. `apply_text_changes.py` creates a `.bak` file automatically.

### Step 5: Verify

Recompile the `.tex` (e.g., `latexmk -pdf main.tex` or `pdflatex main.tex`) and check for errors. If compilation fails, fix the introduced edits and recompile.

## Important limitations

- **Math and macros are fragile.** Pandoc renders LaTeX math as Word text. When converting back, you must restore `$...$` and macros manually; the automatic script will skip paragraphs containing math it cannot match.
- **Figures are embedded as images in Word.** Do not resize or move them if you want the original `.tex` figure code to remain unchanged. Caption text edits are safe if applied manually.
- **Tables may lose formatting.** Simple text edits inside table cells can be synced, but column/row layout changes should be done in LaTeX.
- **References and citations.** Pandoc usually turns `\cite{key}` into plain text like "(Author, 2020)". Do not copy those back literally; keep `\cite{key}` in the `.tex`.

## Recommended prompt style for the user

When the user invokes this skill, typical requests look like:

- "帮我把 main.tex 转成 Word，方便我修改。"
- "我改好了 Word，请把改动同步回 main.tex。"
- "生成 Word 版本并打开。"

Proceed through the five steps above and report which changes were applied automatically and which need manual review.
