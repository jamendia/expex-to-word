# latex-to-word
A single-file Python script that converts a LaTeX (`.tex`) file into a Word (`.docx`) document. It's meant for papers and handouts that mix ordinary prose with math and and numbered linguistic examples with glosses, typeset with the ExPex package.

## Features

- Sections/subsections, itemize/enumerate lists, tables, and text formatting (bold, italics, small caps, etc.)
- Inline and display equations, converted to **native, editable** Word equations (OMML) rather than images
- Numbered linguistic examples with interlinear glosses, written with the [`expex`](https://www.ctan.org/pkg/expex) LaTeX package — rendered as aligned Word tables with automatic numbering, lettered sub-examples, and working cross-references (`\getref`, `\getfullref`, `\lastx`, etc.)

## Requirements

- Python 3.9+
- [`python-docx`](https://python-docx.readthedocs.io/):
  
  ```bash
  pip install python-docx
  ```
- [`texmath`](https://github.com/jgm/texmath) — only needed if your document contains equations *or*, more generally, any dollar sign. (E.g. if you protected indices in movement structures with dollar signs as in $[\ldots]_{i}$ or [\ldots]$_{i}$, the script will try to find `texmath`. it will as for it). This is a small command-line tool (not a Python package) that does the LaTeX-to-OMML conversion. Install it with [Haskell's Cabal](https://www.haskell.org/cabal/) or [Stack](https://docs.haskellstack.org/):
  
  ```bash
  cabal install -fexecutable texmath      # installs to ~/.cabal/bin/texmath
  # or
  stack install --flag texmath:executable # installs to ~/.local/bin/texmath
  ```
  The script looks for the binary at `~/.local/bin/texmath`. If yours ends up somewhere else (e.g. `~/.cabal/bin`), either move / symlink it, or edit the `texmath_path` line in `latex_to_omml()`. If you don't need equations, you can skip installing `texmath` entirely — the script will just leave any `$...$` unconverted and print a warning.

## Usage

```bash
python3 latex_to_word.py paper.tex              # writes paper.docx
python3 latex_to_word.py -o out.docx paper.tex  # choose the output path
python3 latex_to_word.py -v paper.tex           # verbose progress log
python3 latex_to_word.py --help                 # all options
```

## Troubleshooting

**`ModuleNotFoundError: No module named 'exceptions'`** — you have the wrong package installed. PyPI has two similarly-named packages: `docx` (an old, abandoned, Python-2-only package) and `python-docx` the one you need). Fix it with:

```bash
pip uninstall docx
pip install python-docx
```

## Known limitations

General:
- This is a pragmatic converter for common LaTeX usage, not a full LaTeX engine. Unsupported or unusual macros are generally left as-is or dropped silently — check the output against the source for unfamiliar documents.
- Figures/`\includegraphics` and bibliography/citation commands are not handled.

`expex` support specifically does **not** cover:
- `\pextable` (`\labels` / `\tl`)
- `labelgen=list` and `\labellist`
- `\deftag`, the external tag file, and `\gathertags`
- `glftpos=right` and `\beginglpanel` (the free translation is always placed below the gloss)
- IJAL style (`\lingset{style=IJAL}`)
- `chapter.arabic`-style numbering (falls back to plain arabic)

## Development

Built with the assistance of Claude.
