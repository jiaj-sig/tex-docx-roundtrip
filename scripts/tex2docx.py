"""
Convert a LaTeX manuscript to a Word document for easy editing.
"""
import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile


def find_pandoc(extra_paths=None):
    pandoc = shutil.which("pandoc")
    if pandoc:
        return pandoc
    candidates = [
        os.path.expandvars(r"%LOCALAPPDATA%\Pandoc\pandoc.exe"),
        os.path.expandvars(r"%ProgramFiles%\Pandoc\pandoc.exe"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Pandoc\pandoc.exe"),
        os.path.expandvars(r"%USERPROFILE%\pandoc\pandoc.exe"),
    ]
    if extra_paths:
        candidates.extend(extra_paths)
    for c in candidates:
        if c and os.path.isfile(c):
            return c
    return None


def pdf_to_png(pdf_path, out_dir, dpi=150):
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    cmd = ["pdftoppm", "-png", "-r", str(dpi), pdf_path, os.path.join(out_dir, base)]
    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    src = os.path.join(out_dir, f"{base}-1.png")
    dst = os.path.join(out_dir, f"{base}.png")
    if os.path.exists(src):
        os.replace(src, dst)
    return dst if os.path.exists(dst) else None


def convert_figures(tex_dir, work_dir, dpi=150):
    fig_dir = os.path.join(tex_dir, "figure")
    converted = {}
    if not os.path.isdir(fig_dir):
        return converted
    for name in os.listdir(fig_dir):
        src = os.path.join(fig_dir, name)
        base, ext = os.path.splitext(name)
        if ext.lower() in (".png", ".jpg", ".jpeg"):
            dst = os.path.join(work_dir, name)
            shutil.copy2(src, dst)
            converted[name] = name
        elif ext.lower() == ".pdf":
            if shutil.which("pdftoppm"):
                pdf_to_png(src, work_dir, dpi=dpi)
                converted[name] = base + ".png"
            else:
                converted[name] = name
    return converted


def rewrite_tex_for_png(tex_path, work_dir, converted):
    with open(tex_path, "r", encoding="utf-8") as f:
        tex = f.read()
    def repl(m):
        prefix = m.group(1)
        filename = m.group(2)
        base, ext = os.path.splitext(filename)
        basename = os.path.basename(filename)
        if basename.lower() in converted:
            new_name = converted[basename.lower()]
        elif ext.lower() == ".pdf":
            new_name = os.path.basename(base) + ".png"
        else:
            new_name = filename
        return prefix + new_name + "}"
    tex = re.sub(r"(\\includegraphics(?:\[[^\]]*\])?\{)([^}]+)\}", repl, tex)
    out_tex = os.path.join(work_dir, "main.tex")
    with open(out_tex, "w", encoding="utf-8") as f:
        f.write(tex)
    return out_tex


def main():
    parser = argparse.ArgumentParser(description="Convert LaTeX to Word for editing")
    parser.add_argument("tex", help="Input .tex file")
    parser.add_argument("docx", nargs="?", help="Output .docx file")
    parser.add_argument("--pandoc", help="Path to pandoc executable")
    parser.add_argument("--dpi", type=int, default=150, help="DPI for PDF-to-PNG conversion")
    args = parser.parse_args()
    tex_path = os.path.abspath(args.tex)
    if not os.path.isfile(tex_path):
        print(f"Error: file not found: {tex_path}", file=sys.stderr)
        sys.exit(1)
    tex_dir = os.path.dirname(tex_path)
    if args.docx:
        docx_path = os.path.abspath(args.docx)
    else:
        docx_path = os.path.join(tex_dir, os.path.splitext(os.path.basename(tex_path))[0] + ".docx")

    extra_paths = [args.pandoc] if args.pandoc else []
    pandoc = find_pandoc(extra_paths=extra_paths)
    if not pandoc:
        print("Error: pandoc not found. Please install pandoc from https://pandoc.org "
              "or pass --pandoc /path/to/pandoc.exe", file=sys.stderr)
        sys.exit(1)
    work_dir = tempfile.mkdtemp(prefix="tex2docx_")
    try:
        converted = convert_figures(tex_dir, work_dir, dpi=args.dpi)
        print(f"Prepared {len(converted)} figure(s)")
        tmp_tex = rewrite_tex_for_png(tex_path, work_dir, converted)
        cmd = [pandoc, tmp_tex, "-o", docx_path, "--standalone"]
        result = subprocess.run(cmd, cwd=work_dir, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if result.returncode != 0:
            print(f"Error: pandoc failed\n{result.stderr}", file=sys.stderr)
            sys.exit(1)
        if result.stderr:
            print(result.stderr, file=sys.stderr)
        print(f"Success: {docx_path}")
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
