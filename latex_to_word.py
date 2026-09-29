"""
latex_to_word.py -- LaTeX to Word (.docx) converter with equation and expex support

Converts LaTeX .tex files to Word documents, handling:
  * LaTeX structure: sections/subsections, itemize/enumerate, tables, comments
  * Inline and display equations, converted to native OMML (editable Word math)
  * Numbered linguistic examples with interlinear glosses (the expex package),
    rendered as aligned Word tables with automatic numbering and cross-references

Usage:
    python3 latex_to_word.py paper.tex              # writes paper.docx
    python3 latex_to_word.py -o out.docx paper.tex   # custom output path
    python3 latex_to_word.py -v paper.tex            # verbose progress log

Requirements:
    - Python 3.9+
    - python-docx   (pip install python-docx)
    - texmath       (only needed for equation conversion -- see README.md)

See README.md for installation instructions, supported LaTeX/expex syntax,
and known limitations.
"""

import subprocess
import re
import copy
import unicodedata
import shutil
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml import parse_xml
from docx.oxml import OxmlElement
from docx.oxml.ns import nsdecls, qn
import os
from pathlib import Path


# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

def clean_latex_delimiters(latex_formula):
    r"""
    Clean up LaTeX delimiter commands for better OMML conversion.
    
    Handles:
    - \Bigl., \Bigr., \big., etc. (null delimiters) -> remove
    - Converts \Bigl(...) to just (...)
    - Converts \Bigr(...) to just (...)
    - \resizebox{...}{...}{$content$} -> content (strips both wrapper and $ signs)
    
    Args:
        latex_formula (str): LaTeX formula string
    
    Returns:
        str: Cleaned LaTeX formula
    """
    # Remove \resizebox{...}{...}{$...$} and keep only content without $ signs
    # This handles: \resizebox{.97\hsize}{!}{$equation$} -> equation
    def strip_resizebox_and_dollars(match):
        content = match.group(1)
        # Strip leading/trailing $ if present
        content = content.strip()
        if content.startswith('$') and content.endswith('$'):
            content = content[1:-1]
        return content
    
    formula = re.sub(r'\\resizebox\{[^}]*\}\{[^}]*\}\{(.*?)\}', strip_resizebox_and_dollars, latex_formula, flags=re.DOTALL)
    
    # Remove null delimiters: \Bigl., \Bigr., \big., \Big., \bigl., \bigr., etc.
    formula = re.sub(r'\\[Bb]ig[lr]?\.', '', formula)
    
    # Remove \Bigl and \Bigr commands (but keep the delimiter)
    # e.g., \Bigl( becomes (, \Bigr) becomes )
    formula = re.sub(r'\\[Bb]igl\s*', '', formula)
    formula = re.sub(r'\\[Bb]igr\s*', '', formula)
    formula = re.sub(r'\\big\s*', '', formula)
    formula = re.sub(r'\\Big\s*', '', formula)
    formula = re.sub(r'\\bigl\s*', '', formula)
    formula = re.sub(r'\\bigr\s*', '', formula)
    
    return formula


_TEXMATH_PATH = None       # resolved lazily by _find_texmath(), cached after the first call
_TEXMATH_WARNED = False    # so a missing binary is reported once per run, not once per equation


def _find_texmath():
    """Locate the texmath executable, checking common install locations and PATH.
    Returns the path, or None if it can't be found (result is cached)."""
    global _TEXMATH_PATH
    if _TEXMATH_PATH is not None:
        return _TEXMATH_PATH
    candidates = [
        os.path.expanduser("~/.local/bin/texmath"),   # stack install --flag texmath:executable
        os.path.expanduser("~/.cabal/bin/texmath"),   # cabal install -fexecutable
        shutil.which("texmath"),                      # anywhere else on PATH
    ]
    for c in candidates:
        if c and os.path.isfile(c):
            _TEXMATH_PATH = c
            return c
    _TEXMATH_PATH = False   # cache the "not found" result too, so we don't keep re-checking
    return None


def latex_to_omml(latex_formula):
    """
    Convert LaTeX formula to OMML (Office Math Markup Language) using texmath.
    
    Args:
        latex_formula (str): LaTeX string (without $ or $$ delimiters)
    
    Returns:
        str or None: OMML string or None if conversion fails
    """
    global _TEXMATH_WARNED
    texmath_path = _find_texmath()
    if not texmath_path:
        if not _TEXMATH_WARNED:
            print("⚠ texmath executable not found (checked ~/.local/bin, ~/.cabal/bin, "
                  "and PATH). Equations will be left as plain LaTeX text. "
                  "See README.md for installation instructions.")
            _TEXMATH_WARNED = True
        return None

    try:
        # Clean up delimiter commands first
        latex_formula = clean_latex_delimiters(latex_formula)
        
        # Call texmath to convert LaTeX to OMML
        result = subprocess.run(
            [texmath_path, '--from', 'tex', '--to', 'omml'],
            input=latex_formula.encode('utf-8'),
            capture_output=True,
            timeout=5
        )
        
        if result.returncode == 0:
            return result.stdout.decode('utf-8').strip()
        else:
            print(f"Error converting formula: {result.stderr.decode('utf-8')}")
            return None
    except Exception as e:
        print(f"Exception during conversion: {e}")
        return None


def remove_latex_comments(content):
    """
    Remove LaTeX comments from content.
    
    Removes everything from % to the end of line, but preserves % inside commands.
    
    Args:
        content (str): LaTeX content as string
    
    Returns:
        str: Content with comments removed
    """
    lines = content.split('\n')
    result = []
    for line in lines:
        if line.lstrip().startswith('%'):
            continue  # a comment-only line disappears completely in TeX
        # Remove comments (% to end of line, but be careful about \%
        # Split by % and process
        parts = []
        in_verb = False
        i = 0
        while i < len(line):
            if line[i] == '%' and (i == 0 or line[i-1] != '\\'):
                # Found a comment
                break
            parts.append(line[i])
            i += 1
        result.append(''.join(parts).rstrip())
    
    # Remove empty lines at the end of content
    while result and not result[-1]:
        result.pop()
    
    return '\n'.join(result)


def skip_latex_preamble(content):
    r"""
    Skip LaTeX document preamble (everything before \begin{document} or \section).
    
    Args:
        content (str): LaTeX content as string
    
    Returns:
        str: Content starting from main content
    """
    # Look for \begin{document}
    doc_start = content.find(r'\begin{document}')
    if doc_start != -1:
        return content[doc_start + len(r'\begin{document}'):]
    
    # If no \begin{document}, look for first \section, \chapter, etc.
    section_patterns = [r'\section{', r'\chapter{', r'\part{', r'\subsection{']
    earliest_pos = len(content)
    
    for pattern in section_patterns:
        pos = content.find(pattern)
        if pos != -1 and pos < earliest_pos:
            earliest_pos = pos
    
    if earliest_pos < len(content):
        return content[earliest_pos:]
    
    return content


def _looks_like_math(s):
    """Heuristic: is this plausibly real LaTeX math, or plain prose caught between
    two unrelated dollar signs (e.g. an unescaped currency sign)?"""
    s = s.strip()
    if not s:
        return False
    if re.search(r'[\\^_{}]', s):        # commands, sub/superscripts, groups -> real math
        return True
    words = re.findall(r'[A-Za-z]+', s)
    has_operator = re.search(r'[+\-*/=<>]', s) is not None
    if len(words) >= 3 and not has_operator:
        return False                      # several ordinary words, no math operator -> prose
    return True


def extract_latex_equations(content):
    r"""
    Extract LaTeX equations from LaTeX content.
    
    Extracts both display equations (\begin{equation}...\end{equation}, $$...$$)
    and inline equations ($...$).
    
    Args:
        content (str): LaTeX content as string
    
    Returns:
        list: List of tuples (equation, is_display_mode, label) where is_display_mode is bool and label is str or None
    """
    equations = []
    
    # Helper function to extract label(s) from equation
    def extract_label(eq_text):
        # Find all labels (for multi-line equations like align with multiple labels)
        label_matches = re.findall(r'\\label\{([^}]*)\}', eq_text)
        if label_matches:
            # Join multiple labels with comma
            label = ', '.join(label_matches)
            # Remove all labels from equation text
            eq_text = re.sub(r'\\label\{[^}]*\}', '', eq_text).strip()
            return eq_text, label
        return eq_text, None
    
    # First extract equation environments (display)
    # Patterns: \begin{equation*?}...\end{equation*?}, \begin{align*?}...\end{align*?}
    # For align/gather/etc, we need to keep the environment wrapper for texmath
    env_patterns_no_wrapper = [
        (r'\\begin\{equation\*?\}(.*?)\\end\{equation\*?\}', 'equation'),
    ]
    
    env_patterns_with_wrapper = [
        (r'\\begin\{(align\*?)\}(.*?)\\end\{\1\}', 'align'),
        (r'\\begin\{(gather\*?)\}(.*?)\\end\{\1\}', 'gather'),
        (r'\\begin\{(multline\*?)\}(.*?)\\end\{\1\}', 'multline'),
        (r'\\begin\{(split)\}(.*?)\\end\{\1\}', 'split'),
    ]
    
    # Extract equations that don't need wrapper (like equation environment)
    for pattern, env_name in env_patterns_no_wrapper:
        for match in re.finditer(pattern, content, re.DOTALL):
            eq = match.group(1).strip()
            if eq:
                eq_clean, label = extract_label(eq)
                equations.append((eq_clean, True, label))
    
    # Extract equations that need wrapper (like align, gather, etc.)
    for pattern, env_name in env_patterns_with_wrapper:
        for match in re.finditer(pattern, content, re.DOTALL):
            env_type = match.group(1)
            eq_content = match.group(2).strip()
            if eq_content:
                # Extract label before adding wrapper back
                eq_clean, label = extract_label(eq_content)
                # Keep the environment wrapper for texmath to parse alignment correctly
                eq_with_wrapper = f'\\begin{{{env_type}}}\n{eq_clean}\n\\end{{{env_type}}}'
                equations.append((eq_with_wrapper, True, label))
    
    # Remove equation environments from content for next extraction
    temp_content = content
    for pattern, _ in env_patterns_no_wrapper:
        temp_content = re.sub(pattern, '', temp_content, flags=re.DOTALL)
    for pattern, _ in env_patterns_with_wrapper:
        temp_content = re.sub(pattern, '', temp_content, flags=re.DOTALL)
    
    # Extract display equations ($$...$$)
    display_pattern = r'\$\$(.*?)\$\$'
    for match in re.finditer(display_pattern, temp_content, re.DOTALL):
        eq = match.group(1).strip()
        if eq and _looks_like_math(eq):
            eq_clean, label = extract_label(eq)
            equations.append((eq_clean, True, label))  # True = display mode
    
    # Remove display equations from content for inline extraction
    temp_content = re.sub(display_pattern, '', temp_content, flags=re.DOTALL)
    
    # Extract inline equations ($...$)
    inline_pattern = r'\$([^\$]+?)\$'
    for match in re.finditer(inline_pattern, temp_content):
        eq = match.group(1).strip()
        if eq and '\n' not in eq and _looks_like_math(eq):  # skip multiline and non-math
            eq_clean, label = extract_label(eq)
            equations.append((eq_clean, False, label))  # False = inline mode
    
    return equations


def convert_equations_to_omml(equations_list, verbose=False):
    """
    Convert a list of LaTeX equations to OMML format.
    
    Args:
        equations_list (list): List of tuples (equation, is_display, label)
        verbose (bool): If True, print progress information
    
    Returns:
        list: List of tuples (omml, is_display, label) for successfully converted equations
    """
    omml_equations = []
    
    if verbose:
        print("Converting extracted equations to OMML:\n")
    
    n_ok = 0
    for i, (latex_eq, is_display, label) in enumerate(equations_list, 1):
        mode = "DISPLAY" if is_display else "INLINE"
        label_info = f" [Label: {label}]" if label else ""
        if verbose:
            print(f"Converting equation {i} [{mode}]{label_info}: {latex_eq[:40]}...")
        
        omml = latex_to_omml(latex_eq)
        # Always keep one entry per input equation (omml=None on failure) so that
        # later equations don't shift into the wrong slot; the caller falls back
        # to the original LaTeX text wherever omml is None.
        omml_equations.append((omml, is_display, label, latex_eq))
        if omml:
            n_ok += 1
            if verbose:
                print(f"  ✓ Success")
        else:
            if verbose:
                print(f"  ✗ Failed (will keep original LaTeX text instead)")
    
    if verbose:
        print(f"\nSuccessfully converted {n_ok} out of {len(equations_list)} equations")
    
    return omml_equations


# ============================================================================
# LaTeX PROCESSING
# ============================================================================

def process_latex_structure(content):
    r"""
    Convert basic LaTeX structure to plain text with markers for processing.
    
    Handles:
    - Section headers (\section, \subsection, etc.)
    - Basic text processing
    - Preserves equations (marked with placeholders)
    - Extracts figure captions and labels while omitting figure contents
    - Preserves inline equations in figure captions for later conversion
    
    Args:
        content (str): LaTeX content
    
    Returns:
        str: Processed content with LaTeX commands removed/converted
    """
    # Handle \texorpdfstring{arg1}{arg2} - keep only arg2 (do this EARLY before other processing)
    # Need to handle nested braces properly
    def replace_texorpdfstring(content):
        while r'\texorpdfstring{' in content:
            match = re.search(r'\\texorpdfstring\{', content)
            if not match:
                break
            
            start_pos = match.end()
            # Find the end of first argument
            brace_count = 1
            pos = start_pos
            while pos < len(content) and brace_count > 0:
                if content[pos] == '{' and (pos == 0 or content[pos-1] != '\\'):
                    brace_count += 1
                elif content[pos] == '}' and (pos == 0 or content[pos-1] != '\\'):
                    brace_count -= 1
                pos += 1
            
            if brace_count != 0:
                break  # Malformed
            
            # Now pos is at the closing brace of arg1, next should be {arg2}
            if pos >= len(content) or content[pos] != '{':
                break
            
            # Find arg2
            start_arg2 = pos + 1
            brace_count = 1
            pos = start_arg2
            while pos < len(content) and brace_count > 0:
                if content[pos] == '{' and (pos == 0 or content[pos-1] != '\\'):
                    brace_count += 1
                elif content[pos] == '}' and (pos == 0 or content[pos-1] != '\\'):
                    brace_count -= 1
                pos += 1
            
            if brace_count != 0:
                break
            
            # Extract arg2 and replace the whole \texorpdfstring{arg1}{arg2}
            arg2 = content[start_arg2:pos-1]
            content = content[:match.start()] + arg2 + content[pos:]
        
        return content
    
    content = replace_texorpdfstring(content)
    
    # Strip \resizebox{...}{...}{content} and keep only content (without $ if present)
    # Use brace counting to handle nested braces properly
    # Note: This function is now redundant since we strip resizebox earlier in latex_to_word()
    # but keeping it here for any edge cases in process_latex_structure
    def strip_all_resizebox(text):
        iteration = 0
        while r'\resizebox{' in text:
            iteration += 1
            if iteration > 100:  # Safety check to prevent infinite loops
                break
                
            match = re.search(r'\\resizebox\{', text)
            if not match:
                break
            
            # The match.end() position is right after the opening brace of first arg
            # Pattern: \resizebox{width}{height}{content}
            pos = match.end() - 1  # Back to the '{' of first argument
            
            # Skip first two arguments {width}{height}
            for arg_num in range(2):
                if pos >= len(text) or text[pos] != '{':
                    break
                brace_count = 1
                pos += 1  # Move past the opening '{'
                while pos < len(text) and brace_count > 0:
                    if text[pos] == '{' and (pos == 0 or text[pos-1] != '\\'):
                        brace_count += 1
                    elif text[pos] == '}' and (pos == 0 or text[pos-1] != '\\'):
                        brace_count -= 1
                    pos += 1
            
            # Now pos should be at the opening '{' of the third argument (the content)
            if pos >= len(text) or text[pos] != '{':
                break
            
            start_content = pos + 1
            brace_count = 1
            pos = start_content
            while pos < len(text) and brace_count > 0:
                if text[pos] == '{' and (pos == 0 or text[pos-1] != '\\'):
                    brace_count += 1
                elif text[pos] == '}' and (pos == 0 or text[pos-1] != '\\'):
                    brace_count -= 1
                pos += 1
            
            if brace_count != 0:
                break
            
            # Extract content and remove surrounding $ if present
            content_inner = text[start_content:pos-1].strip()
            if content_inner.startswith('$') and content_inner.endswith('$'):
                content_inner = content_inner[1:-1]
            
            # Replace the whole \resizebox{...}{...}{...} with just the content
            text = text[:match.start()] + content_inner + text[pos:]
        
        return text
    
    content = strip_all_resizebox(content)
    
    # Extract figure captions and labels from figure environments
    def extract_figure_info(match):
        fig_content = match.group(1)
        
        # Extract caption - need to handle nested braces properly
        caption_match = re.search(r'\\caption\{', fig_content)
        caption_text = ""
        if caption_match:
            # Find the matching closing brace for \caption{
            start_pos = caption_match.end()
            brace_count = 1
            pos = start_pos
            while pos < len(fig_content) and brace_count > 0:
                if fig_content[pos] == '{' and (pos == 0 or fig_content[pos-1] != '\\'):
                    brace_count += 1
                elif fig_content[pos] == '}' and (pos == 0 or fig_content[pos-1] != '\\'):
                    brace_count -= 1
                pos += 1
            if brace_count == 0:
                caption_text = fig_content[start_pos:pos-1]
        
        label = re.search(r'\\label\{([^}]*)\}', fig_content)
        
        output = []
        if label:
            output.append(f"[Figure: {label.group(1)}]")
        if caption_text:
            # Encapsulate caption in brackets and keep inline equations intact (they'll be processed later)
            output.append(f"[Caption: {caption_text}]")
        
        if output:
            return '\n' + '\n'.join(output) + '\n'
        else:
            return '[Figure omitted - not supported in conversion]\n'
    
    # Handle figure environments - extract captions and labels
    content = re.sub(
        r'\\begin\{figure\*?\}(.*?)\\end\{figure\*?\}',
        extract_figure_info,
        content,
        flags=re.DOTALL
    )
    
    # Handle table environments - extract table content and captions
    def extract_table_info(match):
        table_content = match.group(1)
        
        # Extract caption
        caption_match = re.search(r'\\caption\{', table_content)
        caption_text = ""
        if caption_match:
            start_pos = caption_match.end()
            brace_count = 1
            pos = start_pos
            while pos < len(table_content) and brace_count > 0:
                if table_content[pos] == '{' and (pos == 0 or table_content[pos-1] != '\\'):
                    brace_count += 1
                elif table_content[pos] == '}' and (pos == 0 or table_content[pos-1] != '\\'):
                    brace_count -= 1
                pos += 1
            if brace_count == 0:
                caption_text = table_content[start_pos:pos-1]
        
        # Extract label
        label = re.search(r'\\label\{([^}]*)\}', table_content)
        
        # Extract tabular/tabularx content with proper brace counting
        tabular_start = re.search(r'\\begin\{(tabular[x]?)\}', table_content)
        table_data = None
        
        if tabular_start:
            env_name = tabular_start.group(1)
            pos = tabular_start.end()
            
            # For tabularx, skip two arguments: {width}{column_spec}
            # For tabular, skip one argument: {column_spec}
            num_args_to_skip = 2 if env_name == 'tabularx' else 1
            
            for _ in range(num_args_to_skip):
                if pos < len(table_content) and table_content[pos] == '{':
                    brace_count = 1
                    pos += 1
                    while pos < len(table_content) and brace_count > 0:
                        if table_content[pos] == '{' and (pos == 0 or table_content[pos-1] != '\\'):
                            brace_count += 1
                        elif table_content[pos] == '}' and (pos == 0 or table_content[pos-1] != '\\'):
                            brace_count -= 1
                        pos += 1
            
            # Now extract content until \end{tabular}
            start_pos = pos
            end_pattern = f'\\end{{{env_name}}}'
            end_pos = table_content.find(end_pattern, start_pos)
            
            if end_pos != -1:
                table_data = table_content[start_pos:end_pos].strip()
        
        output = []
        if label:
            output.append(f"[Table: {label.group(1)}]")
        if caption_text:
            output.append(f"[Caption: {caption_text}]")
        
        if table_data:
            # Mark the table content for processing
            output.append(f"__TABLE_START__\n{table_data}\n__TABLE_END__")
        else:
            output.append('[Table content could not be extracted]')
        
        if output:
            return '\n' + '\n'.join(output) + '\n'
        else:
            return '[Table omitted - not supported in conversion]\n'
    
    # Handle table environments
    content = re.sub(
        r'\\begin\{table\*?\}(.*?)\\end\{table\*?\}',
        extract_table_info,
        content,
        flags=re.DOTALL
    )
    
    # Handle itemize environments - convert to marked list items
    def process_itemize(match):
        items_content = match.group(1)
        # Extract individual \item entries
        items = re.findall(r'\\item\s+(.*?)(?=\\item|$)', items_content, re.DOTALL)
        result = []
        for item in items:
            item = item.strip()
            if item:
                result.append(f'__BULLET_ITEM__{item}')
        return '\n' + '\n'.join(result) + '\n'
    
    content = re.sub(
        r'\\begin\{itemize\}(.*?)\\end\{itemize\}',
        process_itemize,
        content,
        flags=re.DOTALL
    )
    
    # Handle enumerate environments - convert to marked numbered list items
    def process_enumerate(match):
        items_content = match.group(1)
        # Extract individual \item entries
        items = re.findall(r'\\item\s+(.*?)(?=\\item|$)', items_content, re.DOTALL)
        result = []
        for item in items:
            item = item.strip()
            if item:
                result.append(f'__NUMBERED_ITEM__{item}')
        return '\n' + '\n'.join(result) + '\n'
    
    content = re.sub(
        r'\\begin\{enumerate\}(.*?)\\end\{enumerate\}',
        process_enumerate,
        content,
        flags=re.DOTALL
    )
    
    # Remove \label{...} and convert reference commands
    content = re.sub(r'\\label\{[^}]*\}', '', content)
    # Convert \ref{label} to [label]
    def replace_ref(match):
        label = match.group(1)
        return f'[{label}]'
    content = re.sub(r'\\ref\{([^}]*)\}', replace_ref, content)
    # Keep citation labels: \cite{label1,label2} -> [label1,label2]
    def replace_cite(match):
        labels = match.group(1)
        return f'[{labels}]'
    content = re.sub(r'\\cite\{([^}]*)\}', replace_cite, content)
    
    # Remove \reffig{...} and \refeqn{...}
    content = re.sub(r'\\reffig\{[^}]*\}', 'Fig.', content)
    content = re.sub(r'\\refeqn\{[^}]*\}', 'Eq.', content)
    
    # Handle text formatting commands
    content = re.sub(r'\\textbf\{([^}]*)\}', r'\1', content)
    content = re.sub(r'\\textit\{([^}]*)\}', r'\1', content)
    content = re.sub(r'\\texttt\{([^}]*)\}', r'\1', content)
    content = re.sub(r'\\emph\{([^}]*)\}', r'\1', content)
    content = re.sub(r'\\text\{([^}]*)\}', r'\1', content)
    
    # Handle subscript and superscript in text (not math mode)
    content = re.sub(r'\\textsubscript\{([^}]*)\}', r'__SUB__\1__/SUB__', content)
    content = re.sub(r'\\textsuperscript\{([^}]*)\}', r'__SUP__\1__/SUP__', content)
    
    # Handle other common commands
    content = re.sub(r'\\mathrm\{([^}]*)\}', r'\1', content)
    content = re.sub(r'\\mathbf\{([^}]*)\}', r'\1', content)
    content = re.sub(r'\\mathcal\{([^}]*)\}', r'\1', content)
    content = re.sub(r'\\mathbb\{([^}]*)\}', r'\1', content)
    
    # Replace LaTeX non-breaking space ~ with regular space
    content = re.sub(r'~', ' ', content)
    
    # Handle line breaks (but protect table content)
    # Split by table markers and only process non-table parts
    parts = re.split(r'(__TABLE_START__.*?__TABLE_END__)', content, flags=re.DOTALL)
    for i in range(len(parts)):
        if not parts[i].startswith('__TABLE_START__'):
            parts[i] = re.sub(r'\\\\', '\n', parts[i])
    content = ''.join(parts)
    
    # Replace escaped percentage signs \% with %
    content = re.sub(r'\\%', '%', content)
    
    # Clean up extra whitespace
    lines = content.split('\n')
    lines = [line.strip() for line in lines]
    content = '\n'.join(lines)
    
    return content


# ============================================================================
# DOCUMENT GENERATION
# ============================================================================

def add_formatted_text(paragraph, text):
    """
    Add text to a paragraph with support for subscript and superscript formatting.
    
    Handles __SUB__text__/SUB__ and __SUP__text__/SUP__ markers.
    
    Args:
        paragraph: The paragraph object to add text to
        text (str): Text with formatting markers
    """
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    
    # Pattern to match subscript and superscript markers
    pattern = r'__SUB__([^_]+?)__/SUB__|__SUP__([^_]+?)__/SUP__'
    
    parts = re.split(pattern, text)
    
    for i, part in enumerate(parts):
        if part is None:
            continue
        if i % 3 == 0:  # Regular text
            if part:
                paragraph.add_run(part)
        elif i % 3 == 1:  # Subscript text (from __SUB__)
            if part:
                run = paragraph.add_run(part)
                run.font.subscript = True
        elif i % 3 == 2:  # Superscript text (from __SUP__)
            if part:
                run = paragraph.add_run(part)
                run.font.superscript = True


def parse_latex_table(table_text):
    """
    Parse LaTeX table content and return a list of rows with proper multicolumn handling.
    
    Args:
        table_text (str): Content between \\begin{tabular} and \\end{tabular}
    
    Returns:
        list: List of rows, where each row is a list of (content, colspan) tuples
    """
    # Remove \hline, \cline, and other formatting commands
    table_text = re.sub(r'\\hline', '', table_text)
    table_text = re.sub(r'\\cline\{[^}]*\}', '', table_text)
    table_text = re.sub(r'\\centering', '', table_text)
    
    # Split by \\ to get rows (but not \\\\)
    rows = re.split(r'\\\\(?!\\)', table_text)
    
    table_data = []
    for row in rows:
        row = row.strip()
        if not row:
            continue
        
        # Helper function to extract content from braces
        def extract_braced_content(text, start_pos):
            """Extract content within braces starting at start_pos (which should be at '{')"""
            if start_pos >= len(text) or text[start_pos] != '{':
                return "", start_pos
            
            content = ""
            pos = start_pos + 1
            brace_count = 1
            
            while pos < len(text) and brace_count > 0:
                if text[pos] == '\\' and pos + 1 < len(text):
                    content += text[pos:pos+2]
                    pos += 2
                elif text[pos] == '{':
                    brace_count += 1
                    content += text[pos]
                    pos += 1
                elif text[pos] == '}':
                    brace_count -= 1
                    if brace_count > 0:
                        content += text[pos]
                    pos += 1
                else:
                    content += text[pos]
                    pos += 1
            
            return content.strip(), pos
        
        # Parse cells with multicolumn support
        cells = []
        i = 0
        current_cell = ""
        brace_count = 0
        
        while i < len(row):
            # Check for \multicolumn{n}{format}{content}
            if row[i:].startswith('\\multicolumn'):
                i += 12
                # Skip whitespace
                while i < len(row) and row[i] in ' \t':
                    i += 1
                
                # Extract colspan (first argument)
                colspan_str, i = extract_braced_content(row, i)
                colspan = int(colspan_str) if colspan_str.isdigit() else 1
                
                # Skip format (second argument)
                _, i = extract_braced_content(row, i)
                
                # Extract content (third argument)
                content, i = extract_braced_content(row, i)
                
                cells.append((content, colspan))
                continue
            
            # Check for \multirow{n}{width}{content}
            elif row[i:].startswith('\\multirow'):
                i += 9
                # Skip whitespace
                while i < len(row) and row[i] in ' \t':
                    i += 1
                
                # Skip first two arguments
                _, i = extract_braced_content(row, i)
                _, i = extract_braced_content(row, i)
                
                # Extract content (third argument)
                content, i = extract_braced_content(row, i)
                
                cells.append((content, 1))
                continue
            
            # Check for & separator
            elif row[i] == '&' and brace_count == 0:
                if current_cell.strip():
                    cells.append((current_cell.strip(), 1))
                current_cell = ""
                i += 1
                continue
            
            # Track braces for non-command content
            elif row[i] == '\\' and i + 1 < len(row):
                current_cell += row[i:i+2]
                i += 2
            elif row[i] == '{':
                brace_count += 1
                current_cell += row[i]
                i += 1
            elif row[i] == '}':
                brace_count -= 1
                current_cell += row[i]
                i += 1
            else:
                current_cell += row[i]
                i += 1
        
        # Add the last cell
        if current_cell.strip():
            cells.append((current_cell.strip(), 1))
        
        if cells:
            table_data.append(cells)
    
    return table_data


def add_paragraph_with_equations(doc, text, inline_eqs, inline_eq_index, verbose=False):
    """
    Add a paragraph to the document, replacing inline equation placeholders with OMML.
    Also handles subscript/superscript formatting.
    
    Args:
        doc: The Document object
        text (str): Text containing $..$ inline equations and formatting markers
        inline_eqs (list): List of OMML inline equations
        inline_eq_index (int): Current index in inline_eqs list
        verbose (bool): Print debug info
    
    Returns:
        int: Updated inline_eq_index
    """
    inline_pattern = r'\$([^\$]+?)\$'
    parts = re.split(inline_pattern, text)
    
    if len(parts) > 1 and any(parts):
        p = doc.add_paragraph()
        for i, part in enumerate(parts):
            if i % 2 == 0:  # Text part
                if part:
                    # Handle subscripts and superscripts in text
                    add_formatted_text(p, part)
            elif not _looks_like_math(part):
                # extract_latex_equations() didn't treat this $..$ as math either
                # (e.g. a stray currency sign) -- keep it as literal text and don't
                # advance inline_eq_index, or every later equation would shift by one.
                add_formatted_text(p, f"${part}$")
            else:  # Equation part (inline)
                if inline_eq_index < len(inline_eqs):
                    omml, eq_latex = inline_eqs[inline_eq_index]
                    if omml is None:
                        add_formatted_text(p, f"${eq_latex}$")
                        if verbose:
                            print(f"  Kept inline equation {inline_eq_index + 1} as plain LaTeX text (conversion unavailable)")
                        inline_eq_index += 1
                        continue
                    try:
                        # For inline equations, extract just the <m:oMath> part
                        if '<m:oMathPara>' in omml:
                            # Extract the inner <m:oMath> from oMathPara
                            start = omml.find('<m:oMath>')
                            end = omml.find('</m:oMath>') + len('</m:oMath>')
                            if start != -1 and end > start:
                                omml_inner = omml[start:end]
                            else:
                                omml_inner = omml
                        else:
                            omml_inner = omml
                        
                        # Add namespace if needed
                        if 'xmlns:m' not in omml_inner:
                            omml_inner = omml_inner.replace(
                                '<m:oMath>',
                                '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
                            )
                        
                        # Create a run and insert OMML
                        run = p.add_run()
                        omml_element = parse_xml(omml_inner)
                        run._element.append(omml_element)
                        if verbose:
                            print(f"  ✓ Added inline equation {inline_eq_index + 1}")
                    except Exception as e:
                        p.add_run(f"[eq: {part}]")
                        if verbose:
                            print(f"  ✗ Failed to add inline equation: {e}")
                    inline_eq_index += 1
                else:
                    # No more OMML equations available, add as text placeholder
                    p.add_run(f"[eq: {part}]")
                    if verbose:
                        print(f"  ⚠ Warning: Not enough OMML equations for inline equation")
    else:
        p = doc.add_paragraph()
        add_formatted_text(p, text)
    
    return inline_eq_index


def create_word_doc_from_latex(latex_content, omml_equations_data, output_path="output.docx", verbose=False, expex_doc=None):
    """
    Create a Word document from LaTeX content with OMML equations.
    
    Args:
        latex_content (str): Full LaTeX text
        omml_equations_data (list): List of (omml, is_display) tuples
        expex_doc (ExpexDocument): Parsed expex examples (placeholders __EXAMPLE_n__ in the text)
        output_path (str): Path where the document will be saved
        verbose (bool): If True, print progress information
    
    Returns:
        str: Path to the created document
    """
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    
    # Process LaTeX structure to plain text
    processed_content = process_latex_structure(latex_content)
    
    # Create a new Document
    doc = Document()
    
    # Create a mapping of equation positions
    display_eqs = [(omml, label, latex_eq) for omml, is_display, label, latex_eq in omml_equations_data if is_display]
    inline_eqs = [(omml, latex_eq) for omml, is_display, label, latex_eq in omml_equations_data if not is_display]
    
    # Index to track which OMML equation we're on
    display_eq_index = 0
    inline_eq_index = 0
    
    # First, replace all display equations with placeholders
    env_patterns = [
        r'\\begin\{equation\*?\}(.*?)\\end\{equation\*?\}',
        r'\\begin\{align\*?\}(.*?)\\end\{align\*?\}',
        r'\\begin\{gather\*?\}(.*?)\\end\{gather\*?\}',
        r'\\begin\{multline\*?\}(.*?)\\end\{multline\*?\}',
        r'\\begin\{split\}(.*?)\\end\{split\}',
    ]
    
    display_placeholders = {}
    placeholder_counter = 0
    for pattern in env_patterns:
        for match in re.finditer(pattern, processed_content, re.DOTALL):
            placeholder = f"__DISPLAY_EQUATION_{placeholder_counter}__"
            display_placeholders[placeholder] = placeholder_counter
            processed_content = processed_content.replace(match.group(0), placeholder, 1)
            placeholder_counter += 1
    
    # Replace $$...$$ equations
    display_pattern = r'\$\$(.*?)\$\$'
    for i, match in enumerate(re.finditer(display_pattern, processed_content, re.DOTALL)):
        placeholder = f"__DISPLAY_EQUATION_{placeholder_counter}__"
        display_placeholders[placeholder] = placeholder_counter
        processed_content = processed_content.replace(match.group(0), placeholder, 1)
        placeholder_counter += 1
    
    # Split into lines and process
    lines = processed_content.split('\n')
    
    current_paragraph = None
    in_table = False
    table_content = []
    
    for line in lines:
        line = line.rstrip()
        
        # expex examples were cut out earlier and left as __EXAMPLE_n__ placeholders
        if not in_table and expex_doc is not None:
            ex_match = re.fullmatch(r'__EXAMPLE_(\d+)__', line.strip())
            if ex_match:
                current_paragraph = None
                render_expex_example(doc, expex_doc.examples[int(ex_match.group(1))],
                                     expex_doc, verbose=verbose)
                continue
        
        # Handle table markers
        if line == '__TABLE_START__':
            in_table = True
            table_content = []
            current_paragraph = None
            continue
        elif line == '__TABLE_END__':
            in_table = False
            if table_content:
                # Parse and create table
                joined_content = '\n'.join(table_content)
                table_rows = parse_latex_table(joined_content)
                if table_rows:
                    # Calculate actual number of columns (accounting for colspan)
                    max_cols = max(sum(colspan for _, colspan in row) for row in table_rows)
                    
                    # Create Word table
                    word_table = doc.add_table(rows=len(table_rows), cols=max_cols)
                    word_table.style = 'Table Grid'
                    
                    # Fill in the table
                    for i, row_data in enumerate(table_rows):
                        col_index = 0
                        for cell_text, colspan in row_data:
                            if col_index < max_cols:
                                cell = word_table.rows[i].cells[col_index]
                                
                                # Merge cells if colspan > 1
                                if colspan > 1 and col_index + colspan <= max_cols:
                                    merge_cell = word_table.rows[i].cells[col_index + colspan - 1]
                                    cell.merge(merge_cell)
                                
                                # Clear the cell first
                                cell.text = ''
                                paragraph = cell.paragraphs[0]
                                
                                # Process cell text with inline equations
                                # Find all inline math expressions
                                inline_math_pattern = r'\$([^\$]+)\$'
                                last_end = 0
                                
                                for match in re.finditer(inline_math_pattern, cell_text):
                                    # Add text before the equation
                                    if match.start() > last_end:
                                        text_before = cell_text[last_end:match.start()]
                                        # Clean LaTeX commands
                                        text_before = re.sub(r'\\textbf\{([^}]*)\}', r'\1', text_before)
                                        text_before = re.sub(r'\\pm', '±', text_before)
                                        paragraph.add_run(text_before)
                                    
                                    # Convert and add the equation
                                    latex_expr = match.group(1)
                                    omml = latex_to_omml(latex_expr)
                                    if omml:
                                        try:
                                            # For inline equations, extract just the <m:oMath> part
                                            if '<m:oMathPara>' in omml:
                                                # Extract the inner <m:oMath> from oMathPara
                                                start = omml.find('<m:oMath>')
                                                end = omml.find('</m:oMath>') + len('</m:oMath>')
                                                if start != -1 and end > start:
                                                    omml_inner = omml[start:end]
                                                else:
                                                    omml_inner = omml
                                            else:
                                                omml_inner = omml
                                            
                                            # Add namespace if needed
                                            if 'xmlns:m' not in omml_inner:
                                                omml_inner = omml_inner.replace(
                                                    '<m:oMath>',
                                                    '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
                                                )
                                            
                                            # Create a run and insert OMML
                                            run = paragraph.add_run()
                                            omml_element = parse_xml(omml_inner)
                                            run._element.append(omml_element)
                                        except Exception as e:
                                            # Fallback to plain text
                                            paragraph.add_run(latex_expr)
                                            if verbose:
                                                print(f"  ✗ Failed to add table cell equation: {e}")
                                    else:
                                        # Fallback to plain text
                                        paragraph.add_run(latex_expr)
                                    
                                    last_end = match.end()
                                
                                # Add remaining text after last equation
                                if last_end < len(cell_text):
                                    text_after = cell_text[last_end:]
                                    # Clean LaTeX commands
                                    text_after = re.sub(r'\\textbf\{([^}]*)\}', r'\1', text_after)
                                    text_after = re.sub(r'\\pm', '±', text_after)
                                    paragraph.add_run(text_after)
                                
                                col_index += colspan
                    
                    if verbose:
                        print(f"✓ Added table with {len(table_rows)} rows and {max_cols} columns")
            table_content = []
            continue
        elif in_table:
            table_content.append(line)
            continue
        
        if not line:
            if current_paragraph is None or not current_paragraph.text.strip():
                current_paragraph = doc.add_paragraph()
            else:
                current_paragraph = None
            continue
        
        # Process LaTeX section commands
        if line.startswith(r'\chapter{'):
            current_paragraph = None
            chapter_text = re.search(r'\\chapter\{([^}]*)\}', line)
            if chapter_text:
                heading = chapter_text.group(1)
                # Use level 0 for chapter (larger than section)
                doc.add_heading(heading, level=0)
        elif line.startswith(r'\section{'):
            current_paragraph = None
            section_text = re.search(r'\\section\{([^}]*)\}', line)
            if section_text:
                heading = section_text.group(1)
                doc.add_heading(heading, level=1)
        elif line.startswith(r'\subsection{'):
            current_paragraph = None
            subsec_text = re.search(r'\\subsection\{([^}]*)\}', line)
            if subsec_text:
                heading = subsec_text.group(1)
                doc.add_heading(heading, level=2)
        elif line.startswith(r'\subsubsection{'):
            current_paragraph = None
            subsubsec_text = re.search(r'\\subsubsection\{([^}]*)\}', line)
            if subsubsec_text:
                heading = subsubsec_text.group(1)
                doc.add_heading(heading, level=3)
        elif line.startswith('__BULLET_ITEM__'):
            # Handle bullet list items
            current_paragraph = None
            item_text = line[len('__BULLET_ITEM__'):]
            inline_eq_index = add_paragraph_with_equations(
                doc, item_text, inline_eqs, inline_eq_index, verbose=verbose
            )
            # Add bullet style to the last paragraph
            doc.paragraphs[-1].style = 'List Bullet'
        elif line.startswith('__NUMBERED_ITEM__'):
            # Handle numbered list items
            current_paragraph = None
            item_text = line[len('__NUMBERED_ITEM__'):]
            inline_eq_index = add_paragraph_with_equations(
                doc, item_text, inline_eqs, inline_eq_index, verbose=verbose
            )
            # Add numbered list style to the last paragraph
            doc.paragraphs[-1].style = 'List Number'
        else:
            # Regular text - replace equations with OMML
            # First check for display equation placeholders
            placeholder_match = re.search(r'__DISPLAY_EQUATION_(\d+)__', line)
            if placeholder_match:
                # This line contains a display equation placeholder
                current_paragraph = None
                eq_index = int(placeholder_match.group(1))
                if eq_index < len(display_eqs):
                    eq_para = doc.add_paragraph()
                    omml, label, eq_latex = display_eqs[eq_index]
                    if omml is None:
                        eq_para.add_run(f"$${eq_latex}$$")
                        if verbose:
                            print(f"  Kept display equation {eq_index + 1} as plain LaTeX text (conversion unavailable)")
                        if label:
                            doc.add_paragraph(f"[{label}]")
                        continue
                    try:
                        if 'xmlns:m' not in omml:
                            omml_with_ns = omml.replace(
                                '<m:oMathPara>',
                                '<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
                            )
                        else:
                            omml_with_ns = omml
                        omml_element = parse_xml(omml_with_ns)
                        eq_para._element.append(omml_element)
                        if verbose:
                            print(f"✓ Added display equation {eq_index + 1}")
                        
                        # Add equation label as a separate paragraph below the equation if it exists
                        if label:
                            label_para = doc.add_paragraph(f"[{label}]")
                            if verbose:
                                print(f"  Added label: [{label}]")
                    except Exception as e:
                        eq_para.add_run(f"[Equation - parsing error]")
                        if verbose:
                            print(f"✗ Error adding display equation {eq_index + 1}: {e}")
            else:
                # Handle inline equations using the helper function
                current_paragraph = None
                inline_eq_index = add_paragraph_with_equations(
                    doc, line, inline_eqs, inline_eq_index, verbose=verbose
                )
    
    # Save the document
    doc.save(output_path)
    if verbose:
        print(f"Document saved to: {output_path}")
    return output_path


# ============================================================================
# EXPEX SUPPORT: numbered linguistic examples with interlinear glosses
# ============================================================================
#
# Supported expex syntax (see expex.tex for the reference implementation):
#
#   \ex[opts]<tag> ... \xe              single numbered example
#   \pex[opts]<tag> preamble \a[..]<sub> ... \a ... \xe   lettered sub-examples
#                                       (nested \pex inside \a is supported)
#   \exdisplay ... \xe                  display material (no number printed)
#   \begingl[opts] \gla ..// \glb ..// \glc ..// \glft ..// \endgl   (wrap style)
#   \begingl[glstyle=nlevel] word[gloss/gloss] ... \glft ... \endgl (nlevel style)
#   \glpreamble ..//, \nogloss{..}, + (line break), [ ] brackets in \gla
#   \ljudge{*} / \judge{*}              grammaticality judgments
#   \lingset{labeltype=.., labelformat=.., exnoformat=.., exnotype=..,
#            everygla=.., glneveryline=..}
#   \ex[exno=..]                        special example numbers
#   \getref{tag}, \getref{tag.sub}, \getfullref{tag.sub},
#   \nextx \anextx \lastx \blastx \bblastx, \label / \ref on examples
#
# Pipeline: examples are cut out of the LaTeX source *before* equation
# extraction and replaced by "__EXAMPLE_n__" placeholder lines.  When the Word
# document is generated, each placeholder is rendered as a borderless Word
# table (number | label | content) with nested tables for the gloss columns.

_EX_OPEN_RE = re.compile(r'\\(pex|exdisplay|ex)(?![A-Za-z@])')
_EX_TOKEN_RE = re.compile(r'\\(pex|exdisplay|ex|xe)(?![A-Za-z@])')
_EXPEX_LINGSET_RE = re.compile(r'\\lingset\s*\{')
_EXPEX_LABEL_RE = re.compile(r'\\label\s*\{([^}]*)\}')
_EXPEX_USED_RE = re.compile(
    r'\\usepackage\s*(\[[^\]]*\])?\s*\{[^}]*\bexpex\b[^}]*\}|\\input\s+expex\b'
    r'|\\begingl(?![A-Za-z@])|\\pex(?![A-Za-z@])|\\xe(?![A-Za-z@])')

# Rough text-width model used to size columns and to wrap long glosses.
EXPEX_CHAR_WIDTH_IN = 0.095    # average character width (inches)
EXPEX_WORD_GAP_IN = 0.14       # gap between gloss columns (inches)
EXPEX_MATH_CHAR_FACTOR = 0.8   # math source chars are wider than rendered ones
EXPEX_SPACER_PT = 6            # blank space after each example (points)

# labeltype -> (label generator, first label, labelformat, fullrefformat)
_EXPEX_LABELTYPES = {
    'alpha':   ('char', 'a', 'A.', 'XA'),
    'caps':    ('char', 'A', 'A.', 'XA'),
    'numeric': ('number', 1, 'A.', 'X.A'),
    'roman':   ('roman', 1, '(A)', 'XA'),
}


# ----------------------------------------------------------------------------
# small parsing helpers
# ----------------------------------------------------------------------------

def _roman(n):
    """Lower-case roman numeral (like TeX's \\romannumeral)."""
    out = ''
    for val, sym in [(1000, 'm'), (900, 'cm'), (500, 'd'), (400, 'cd'), (100, 'c'),
                     (90, 'xc'), (50, 'l'), (40, 'xl'), (10, 'x'), (9, 'ix'),
                     (5, 'v'), (4, 'iv'), (1, 'i')]:
        while n >= val:
            out += sym
            n -= val
    return out


def _skip_ws(s, i):
    while i < len(s) and s[i].isspace():
        i += 1
    return i


def _match_brace(s, i):
    """s[i] == '{'. Return the index of the matching '}' or -1."""
    depth = 0
    j = i
    while j < len(s):
        c = s[j]
        if c == '\\':
            j += 2
            continue
        if c == '{':
            depth += 1
        elif c == '}':
            depth -= 1
            if depth == 0:
                return j
        j += 1
    return -1


def _match_bracket(s, i):
    """s[i] == '['. Return the index of the matching ']' (ignoring brackets
    inside braces) or -1."""
    depth = 0
    brace = 0
    j = i
    while j < len(s):
        c = s[j]
        if c == '\\':
            j += 2
            continue
        if c == '{':
            brace += 1
        elif c == '}':
            brace -= 1
        elif brace == 0 and c == '[':
            depth += 1
        elif brace == 0 and c == ']':
            depth -= 1
            if depth == 0:
                return j
        j += 1
    return -1


def _split_top(s, sep=','):
    """Split on `sep` at brace depth 0."""
    parts, depth, cur, j = [], 0, [], 0
    while j < len(s):
        c = s[j]
        if c == '\\' and j + 1 < len(s):
            cur.append(s[j:j + 2])
            j += 2
            continue
        if c == '{':
            depth += 1
        elif c == '}':
            depth -= 1
        if c == sep and depth == 0:
            parts.append(''.join(cur))
            cur = []
        else:
            cur.append(c)
        j += 1
    parts.append(''.join(cur))
    return parts


def _strip_group(v):
    """Remove one pair of outer braces if they enclose the whole string."""
    v = v.strip()
    if v.startswith('{') and _match_brace(v, 0) == len(v) - 1:
        return v[1:-1]
    return v


def _parse_keyvals(s):
    """Parse 'k1=v1, k2={v,2}, flag' into a dict."""
    out = {}
    for item in _split_top(s, ','):
        item = item.strip()
        if not item:
            continue
        if '=' in item:
            k, v = item.split('=', 1)
            out[k.strip()] = _strip_group(v)
        else:
            out[item] = ''
    return out


def _unescaped_matches(regex, text, start=0):
    """Yield regex matches for control words that are not preceded by a backslash."""
    for m in regex.finditer(text, start):
        if m.start() > 0 and text[m.start() - 1] == '\\':
            continue
        yield m


def _find_xe(text, start):
    """Index of the \\xe that closes an example whose opener ends at `start`."""
    depth = 1
    for m in _unescaped_matches(_EX_TOKEN_RE, text, start):
        if m.group(1) == 'xe':
            depth -= 1
            if depth == 0:
                return m.start()
        else:
            depth += 1
    return len(text)


def _parse_header(s, i):
    """Parse optional [opts] and <tag> following \\ex / \\pex / \\a.
    Returns (opts_dict, tag_or_None, index_after_header)."""
    opts, tag = {}, None
    j = _skip_ws(s, i)
    if j < len(s) and s[j] == '[':
        k = _match_bracket(s, j)
        if k != -1:
            opts = _parse_keyvals(s[j + 1:k])
            i = k + 1
            j = _skip_ws(s, i)
    if j < len(s) and s[j] == '<':
        k = s.find('>', j)
        if k != -1:
            tag = s[j + 1:k].strip()
            i = k + 1
    return opts, tag, i


def _flags_from_decl(s):
    """Translate declarations like '\\it', '\\bf', '\\sc' into a set of flags."""
    flags = set()
    for cmd in re.findall(r'\\([A-Za-z]+)', s or ''):
        if cmd in ('it', 'itshape', 'sl', 'slshape', 'em', 'textit', 'emph'):
            flags.add('italic')
        elif cmd in ('bf', 'bfseries', 'textbf'):
            flags.add('bold')
        elif cmd in ('sc', 'scshape', 'textsc'):
            flags.add('smallcaps')
        elif cmd in ('tt', 'ttfamily', 'texttt'):
            flags.add('mono')
        elif cmd in ('rm', 'rmfamily', 'normalfont', 'upshape', 'mdseries',
                     'textrm', 'textnormal'):
            flags.clear()
    return flags


# ----------------------------------------------------------------------------
# settings (\lingset) and numbering formats
# ----------------------------------------------------------------------------

class ExpexSettings:
    """The subset of expex's \\lingset parameters that affect the output text."""

    def __init__(self):
        self.exnoformat = '(X)'
        self.exnotype = 'arabic'
        self.glstyle = 'wrap'
        self.everygl = {'a': {'italic'}}        # expex default: everygla=\it
        self.everyline = [{'italic'}]           # expex default: glneveryline={\it}
        self.set_labeltype('alpha')

    def copy(self):
        return copy.deepcopy(self)

    def set_labeltype(self, name):
        gen, start, fmt, full = _EXPEX_LABELTYPES[name]
        self.labeltype = name
        self.labelgen, self.label_start = gen, start
        self.labelformat, self.fullrefformat = fmt, full

    def apply(self, opts):
        for key, val in opts.items():
            v = val.strip() if isinstance(val, str) else ''
            if key == 'labeltype' and v in _EXPEX_LABELTYPES:
                self.set_labeltype(v)
            elif key == 'labelformat':
                self.labelformat = v
            elif key == 'fullrefformat':
                self.fullrefformat = v
            elif key == 'exnoformat':
                self.exnoformat = v
            elif key == 'exnotype':
                self.exnotype = v
            elif key == 'glstyle':
                self.glstyle = v
            elif key == 'glneveryline':
                self.everyline = [_flags_from_decl(x) for x in _split_top(v, ',')]
            elif re.fullmatch(r'everygl[a-z]', key):
                self.everygl[key[-1]] = _flags_from_decl(v)

    # -- example numbers ---------------------------------------------------
    def raw_exno(self, n):
        if self.exnotype == 'roman':
            return _roman(n)
        return str(n)

    def format_exno(self, raw):
        pre, sep, post = self.exnoformat.partition('X')
        return pre + raw + post if sep else self.exnoformat + raw

    # -- part labels -------------------------------------------------------
    def raw_label(self, idx):
        if self.labelgen == 'char':
            return chr(ord(self.label_start) + idx)
        if self.labelgen == 'number':
            return str(int(self.label_start) + idx)
        return _roman(int(self.label_start) + idx)

    def format_label(self, raw):
        pre, sep, post = self.labelformat.partition('A')
        return pre + raw + post if sep else self.labelformat + raw

    def fullref(self, exno_raw, label_raw):
        m = re.match(r'^(.*?)X(.*?)A(.*)$', self.fullrefformat, re.DOTALL)
        if not m:
            return exno_raw + label_raw
        return m.group(1) + exno_raw + m.group(2) + label_raw + m.group(3)


# ----------------------------------------------------------------------------
# parsed example structures
# ----------------------------------------------------------------------------

class ExNode:
    """One \\ex / \\pex / \\exdisplay (or a nested \\pex)."""

    def __init__(self, kind):
        self.kind = kind            # 'ex', 'pex', 'exdisplay'
        self.blocks = []            # for ex/exdisplay: list of blocks
        self.preamble = []          # for pex: blocks before the first \a
        self.parts = []             # for pex: list of PartNode
        self.number_text = ''       # e.g. "(12)"; '' for exdisplay / nested
        self.raw = ''
        self.top_raw = ''
        self.tag = None
        self.excnt = 1
        self.step = 1


class PartNode:
    """One \\a part of a \\pex."""

    def __init__(self, label_text, blocks):
        self.label_text = label_text
        self.blocks = blocks


class GlossNode:
    """An interlinear gloss: columns of words, one cell per gloss line."""

    def __init__(self):
        self.preamble = None
        self.columns = []           # [{'cells': {'a': 'word', 'b': 'gloss'}}, ...]
        self.breaks = set()         # column indices that start a new row of words
        self.letters = []           # gloss line letters in order, e.g. ['a','b']
        self.line_flags = {}        # letter -> base format flags
        self.ft = []                # free translation(s)
        self.judge = None           # grammaticality judgment from \ljudge


class ExpexDocument:
    """All examples of one .tex file plus the tag/label tables for references."""

    def __init__(self, settings=None):
        self.settings = settings.copy() if settings else ExpexSettings()
        self.examples = []          # top-level ExNode objects, in document order
        self.tags = {}              # tag -> raw number | (label_raw, fullref)
        self.labels = {}            # \label name -> text printed by \ref

    def resolve_refs(self, text, excnt):
        """Replace \\getref, \\getfullref, \\lastx & co. and \\ref to example labels."""
        st = self.settings

        for name, delta in (('nextx', 0), ('anextx', 1), ('lastx', -1),
                            ('blastx', -2), ('bblastx', -3)):
            text = re.sub(r'\\' + name + r'(?![A-Za-z@])(\s*\{\})?',
                          lambda m, d=delta: st.raw_exno(max(excnt + d, 0)), text)

        def getref(m):
            full = m.group(1) == 'getfullref'
            tag = m.group(2).strip()
            val = self.tags.get(tag)
            if val is None:
                return '[' + tag + ']'
            if isinstance(val, tuple):
                return val[1] if full else val[0]
            return val

        text = re.sub(r'\\(getref|getfullref)\s*\{([^}]*)\}', getref, text)

        def ref(m):
            return self.labels.get(m.group(1).strip(), m.group(0))

        text = re.sub(r'\\ref\s*\{([^}]*)\}', ref, text)
        return text


# ----------------------------------------------------------------------------
# parsing: examples, parts, blocks, glosses
# ----------------------------------------------------------------------------

def _split_parts(rest):
    """Split the body of a \\pex at its top-level \\a commands."""
    depth = 0
    starts = []
    for m in _unescaped_matches(re.compile(r'\\(pex|exdisplay|ex|xe|a)(?![A-Za-z@])'), rest):
        name = m.group(1)
        if name in ('pex', 'ex', 'exdisplay'):
            depth += 1
        elif name == 'xe':
            depth -= 1
        elif depth == 0:
            starts.append(m)
    if not starts:
        return rest, []
    pre = rest[:starts[0].start()]
    parts = []
    for k, m in enumerate(starts):
        end = starts[k + 1].start() if k + 1 < len(starts) else len(rest)
        parts.append(rest[m.end():end])
    return pre, parts


def _tokenize_words(s):
    """Whitespace-split honouring braces, like TeX's space-delimited arguments.
    A token that is one complete {group} loses its outer braces."""
    tokens, i, n = [], 0, len(s)
    while i < n:
        i = _skip_ws(s, i)
        if i >= n:
            break
        j, depth = i, 0
        while j < n:
            c = s[j]
            if c == '\\':
                j += 2
                continue
            if c == '{':
                depth += 1
            elif c == '}':
                depth -= 1
            elif c.isspace() and depth <= 0:
                break
            j += 1
        tokens.append(_strip_group(s[i:j]) if s[i] == '{' else s[i:j])
        i = j
    return tokens


def _parse_gloss(gbody, st):
    """Parse the text between \\begingl and \\endgl."""
    j = _skip_ws(gbody, 0)
    gopts = {}
    if gbody[j:j + 1] == '[':
        k = _match_bracket(gbody, j)
        if k != -1:
            gopts = _parse_keyvals(gbody[j + 1:k])
            gbody = gbody[k + 1:]
    gst = st.copy()
    gst.apply(gopts)
    if gst.glstyle == 'nlevel':
        gloss = _parse_gloss_nlevel(gbody)
        gloss.line_flags = {l: (gst.everyline[i] if i < len(gst.everyline) else set())
                            for i, l in enumerate(gloss.letters)}
    else:
        gloss = _parse_gloss_wrap(gbody)
        gloss.line_flags = {l: gst.everygl.get(l, set()) for l in gloss.letters}
    return gloss


def _parse_gloss_wrap(gbody):
    """wrap style:  \\gla w1 w2 // \\glb g1 g2 // \\glft 'free translation' //"""
    gloss = GlossNode()
    lines = {}
    item_re = re.compile(r'\\gl(ft|preamble|[a-z])(?![A-Za-z@])')
    pos = 0
    while True:
        m = item_re.search(gbody, pos)
        if not m:
            break
        name, j = m.group(1), m.end()
        j2 = _skip_ws(gbody, j)
        if gbody[j2:j2 + 1] == '[':          # optional argument, ignored
            k = _match_bracket(gbody, j2)
            if k != -1:
                j = k + 1
        end = gbody.find('//', j)
        if end == -1:
            end = len(gbody)
        content = gbody[j:end]
        pos = min(end + 2, len(gbody))
        if name == 'ft':
            gloss.ft.append(content.strip())
        elif name == 'preamble':
            gloss.preamble = content.strip()
        else:
            lines[name] = _tokenize_words(content)

    if 'a' not in lines:
        return gloss

    cols, real, pending = [], [], ''
    for tok in lines['a']:
        if tok == '+':                       # forced line break
            gloss.breaks.add(len(cols))
        elif tok == '@':                     # (spacing control) - ignored
            continue
        elif tok == '[':
            pending += '['
        elif tok == ']':
            if cols:
                cols[-1]['cells']['a'] += ']'
        elif tok.startswith('\\nogloss'):    # word without gloss counterpart
            body = _strip_group(re.sub(r'^\\nogloss\s*', '', tok))
            cols.append({'cells': {'a': pending + body}})
            pending = ''
        else:
            col = {'cells': {'a': pending + tok}}
            cols.append(col)
            real.append(col)
            pending = ''
    for letter in sorted(l for l in lines if l != 'a'):
        for idx, tok in enumerate(lines[letter]):
            if idx >= len(real):             # more gloss words than source words
                col = {'cells': {'a': ''}}
                cols.append(col)
                real.append(col)
            real[idx]['cells'][letter] = tok
    gloss.columns = cols
    gloss.letters = ['a'] + sorted(l for l in lines if l != 'a')
    return gloss


def _parse_gloss_nlevel(gbody):
    """nlevel style:  word[gloss1/gloss2] word[...]  \\glft translation"""
    gloss = GlossNode()
    m = re.search(r'\\glpreamble(?![A-Za-z@])(.*?)\\endpreamble', gbody, re.DOTALL)
    if m:
        gloss.preamble = m.group(1).strip()
        gbody = gbody[:m.start()] + gbody[m.end():]
    fm = re.search(r'\\glft(?![A-Za-z@])', gbody)
    if fm:
        ft = gbody[fm.end():].strip()
        if ft:
            gloss.ft.append(ft)
        gbody = gbody[:fm.start()]

    n, i = len(gbody), 0
    max_letters = 1
    while True:
        i = _skip_ws(gbody, i)
        if i >= n:
            break
        if gbody.startswith('\\nogloss', i):
            i2 = _skip_ws(gbody, i + len('\\nogloss'))
            if gbody[i2:i2 + 1] == '{':
                k = _match_brace(gbody, i2)
                if k != -1:
                    gloss.columns.append({'cells': {'a': gbody[i2 + 1:k]}})
                    i = k + 1
                    continue
        j, depth = i, 0
        while j < n:
            c = gbody[j]
            if c == '\\':
                j += 2
                continue
            if c == '{':
                depth += 1
            elif c == '}':
                depth -= 1
            elif depth == 0 and (c == '[' or c.isspace()):
                break
            j += 1
        cells = {'a': _strip_group(gbody[i:j])}
        if j < n and gbody[j] == '[':
            k = _match_bracket(gbody, j)
            if k == -1:
                k = n - 1
            for idx, p in enumerate(_split_top(gbody[j + 1:k], '/')):
                cells[chr(ord('b') + idx)] = _strip_group(p.strip())
                max_letters = max(max_letters, idx + 2)
            j = k + 1
            d = j
            while d < n and not gbody[d].isspace():
                d += 1
            if '+' in gbody[j:d]:            # '+' after a word = line break
                gloss.breaks.add(len(gloss.columns) + 1)
            j = d
        gloss.columns.append({'cells': cells})
        i = j
    gloss.letters = [chr(ord('a') + k) for k in range(max_letters)]
    return gloss


def _parse_body_blocks(body, st, xdoc, label_value, top_raw='', parent_tag=None):
    """Split example text into blocks: ('text', s) | ('gloss', GlossNode) | ('sub', ExNode).
    \\label commands found in the text are registered with `label_value`."""
    blocks, pos = [], 0
    tok = re.compile(r'\\begingl(?![A-Za-z@])|\\pex(?![A-Za-z@])')
    endgl = re.compile(r'\\endgl(?![A-Za-z@])')

    def add_text(t):
        def reg(m):
            xdoc.labels.setdefault(m.group(1).strip(), label_value)
            return ''
        t = _EXPEX_LABEL_RE.sub(reg, t).strip()
        if t:
            blocks.append(('text', t))

    while True:
        m = next(_unescaped_matches(tok, body, pos), None)
        if m is None:
            break
        add_text(body[pos:m.start()])
        if m.group(0).startswith('\\begingl'):
            e = endgl.search(body, m.end())
            gend = e.start() if e else len(body)
            blocks.append(('gloss', _parse_gloss(body[m.end():gend], st)))
            pos = e.end() if e else len(body)
        else:
            end = _find_xe(body, m.end())
            sub = _build_example('pex', body[m.end():end], st, xdoc, None,
                                 top_raw=top_raw, nested=True, parent_tag=parent_tag)
            blocks.append(('sub', sub))
            pos = min(end + len('\\xe'), len(body))
    add_text(body[pos:])

    # a lone \ljudge{*} in front of a gloss belongs to the gloss's first line
    jre = re.compile(r'^\\l?judge\s*\{(.*)\}$', re.DOTALL)
    merged, k = [], 0
    while k < len(blocks):
        b = blocks[k]
        if (b[0] == 'text' and k + 1 < len(blocks) and blocks[k + 1][0] == 'gloss'
                and jre.match(b[1])):
            blocks[k + 1][1].judge = jre.match(b[1]).group(1)
        else:
            merged.append(b)
        k += 1
    return merged


def _build_example(kind, body, settings, xdoc, excnt, top_raw='', nested=False,
                   parent_tag=None):
    """Parse one example body (text between the opener and \\xe)."""
    node = ExNode(kind)
    opts, tag, i = _parse_header(body, 0)
    st = settings.copy()
    st.apply(opts)
    node.tag = tag or opts.get('tag') or parent_tag or None
    rest = body[i:]

    if not nested:
        n_noexno = len(re.findall(r'\\noexno(?![A-Za-z@])', rest))
        rest = re.sub(r'\\noexno(?![A-Za-z@])', '', rest)
        special = opts.get('exno')
        node.excnt = excnt
        node.raw = special if special is not None else st.raw_exno(excnt)
        node.top_raw = node.raw
        node.number_text = '' if kind == 'exdisplay' else st.format_exno(node.raw)
        node.step = 0 if special is not None else 1 - n_noexno
        if node.tag:
            xdoc.tags[node.tag] = node.raw
    else:
        node.top_raw = top_raw
        node.raw = top_raw

    if kind == 'pex':
        pre, part_texts = _split_parts(rest)
        node.preamble = (_parse_body_blocks(pre, st, xdoc, node.raw, node.top_raw, node.tag)
                         if pre.strip() else [])
        if not part_texts:                   # \pex without \a: behaves like \ex
            node.kind = 'ex'
            node.blocks, node.preamble = node.preamble, []
            return node
        idx = 0
        for ptext in part_texts:
            popts, ptag, j = _parse_header(ptext, 0)
            special_label = popts.get('label')
            if special_label is None:
                raw_label = st.raw_label(idx)
                idx += 1
            else:
                raw_label = special_label
            full = st.fullref(node.top_raw, raw_label)
            ptag = ptag or popts.get('tag')
            if ptag:
                xdoc.tags[(node.tag + '.' if node.tag else '.') + ptag] = (raw_label, full)
            blocks = _parse_body_blocks(ptext[j:], st, xdoc, full, node.top_raw, node.tag)
            node.parts.append(PartNode(st.format_label(raw_label), blocks))
    else:
        node.blocks = _parse_body_blocks(rest, st, xdoc, node.raw, node.top_raw, node.tag)
    return node


def collect_expex_preamble_settings(full_content):
    """Read \\lingset{...} commands from the preamble (before \\begin{document})."""
    settings = ExpexSettings()
    doc_start = full_content.find(r'\begin{document}')
    if doc_start == -1:
        return settings
    preamble = full_content[:doc_start]
    for m in _EXPEX_LINGSET_RE.finditer(preamble):
        k = _match_brace(preamble, m.end() - 1)
        if k != -1:
            settings.apply(_parse_keyvals(preamble[m.end():k]))
    return settings


def extract_expex_examples(content, settings=None, verbose=False):
    """Cut expex examples out of `content`.

    Returns (new_content, ExpexDocument). Every top-level example is replaced by
    a line "__EXAMPLE_n__"; references to examples in the surrounding text
    (\\getref, \\lastx, \\ref{...}) are replaced by the example numbers.
    """
    xdoc = ExpexDocument(settings)
    if not _EXPEX_USED_RE.search(content) and not _EXPEX_LINGSET_RE.search(content):
        return content, xdoc

    segments = []                    # ('text', str, excnt) | ('ex', index)
    pos, excnt = 0, 1
    tok = re.compile(_EX_OPEN_RE.pattern + '|' + _EXPEX_LINGSET_RE.pattern)
    while True:
        m = next(_unescaped_matches(tok, content, pos), None)
        if m is None:
            break
        segments.append(('text', content[pos:m.start()], excnt))
        if m.group(0).startswith('\\lingset'):
            k = _match_brace(content, m.end() - 1)
            if k == -1:
                pos = m.end()
                continue
            xdoc.settings.apply(_parse_keyvals(content[m.end():k]))
            pos = k + 1
            continue
        end = _find_xe(content, m.end())
        if end >= len(content) and verbose:
            print(f"  ⚠ Warning: \\{m.group(1)} without matching \\xe")
        node = _build_example(m.group(1), content[m.end():end], xdoc.settings, xdoc, excnt)
        segments.append(('ex', len(xdoc.examples)))
        xdoc.examples.append(node)
        if verbose:
            kind = node.kind + (f" ({len(node.parts)} parts)" if node.parts else '')
            print(f"  Example {node.number_text or '(unnumbered)'}: {kind}")
        excnt += node.step
        pos = min(end + len('\\xe'), len(content))
    segments.append(('text', content[pos:], excnt))

    out = []
    for si, seg in enumerate(segments):
        if seg[0] == 'ex':
            out.append(f"\n__EXAMPLE_{seg[1]}__\n")
        else:
            text = xdoc.resolve_refs(seg[1], seg[2])
            # blank lines around an example would turn into empty paragraphs
            if si + 1 < len(segments):
                text = text.rstrip()
            if si > 0:
                text = text.lstrip()
            text = re.sub(r'\\(gathertags|refproofing|keepexcntlocal)(?![A-Za-z@])', '', text)
            out.append(text)
    return ''.join(out), xdoc


# ----------------------------------------------------------------------------
# inline LaTeX -> formatted runs
# ----------------------------------------------------------------------------

_COMBINING = {"'": '\u0301', '`': '\u0300', '^': '\u0302', '"': '\u0308', '~': '\u0303',
              '=': '\u0304', '.': '\u0307', 'u': '\u0306', 'v': '\u030C', 'H': '\u030B',
              'c': '\u0327', 'k': '\u0328', 'r': '\u030A', 'b': '\u0331', 'd': '\u0323'}
_SPACING_ACCENT = {"'": '\u00b4', '`': '`', '^': '^', '"': '\u00a8', '~': '~', '=': '\u00af'}
_SPECIAL_LETTERS = {'ss': 'ß', 'o': 'ø', 'O': 'Ø', 'ae': 'æ', 'AE': 'Æ', 'oe': 'œ',
                    'OE': 'Œ', 'aa': 'å', 'AA': 'Å', 'l': 'ł', 'L': 'Ł', 'i': 'ı', 'j': 'ȷ',
                    'ldots': '…', 'dots': '…', 'textendash': '–', 'textemdash': '—',
                    'textbackslash': '\\', 'textasciitilde': '~', 'textquoteleft': '‘',
                    'textquoteright': '’', 'textquotedblleft': '“',
                    'textquotedblright': '”', 'textless': '<', 'textgreater': '>',
                    'textbar': '|', 'textdegree': '°', 'textasteriskcentered': '*',
                    'S': '§', 'P': '¶', 'LaTeX': 'LaTeX', 'TeX': 'TeX', 'tspace': '\u2003',
                    'tspacea': '\u2003', 'tspaceb': '\u2003', 'tspacec': '\u2003',
                    'thinspace': '\u2009', 'enspace': '\u2002', 'quad': '\u2003',
                    'qquad': '\u2003'}
_DROP_COMMANDS = {'glstrut', 'strut', 'noindent', 'centering', 'relax', 'par', 'hfil',
                  'hfill', 'null', 'tiny', 'scriptsize', 'footnotesize', 'small',
                  'normalsize', 'large', 'Large', 'LARGE', 'huge', 'Huge', 'sffamily',
                  'sf', 'nolinebreak', 'linebreak', 'newline', 'allowbreak'}
# commands taking one argument whose content is kept, with the format they set
_ARG_FORMAT = {'textit': ('+', 'italic'), 'textsl': ('+', 'italic'),
               'emph': ('~', 'italic'), 'textbf': ('+', 'bold'),
               'textsc': ('+', 'smallcaps'), 'texttt': ('+', 'mono'),
               'textrm': ('clear', None), 'textnormal': ('clear', None),
               'textup': ('-', 'italic'), 'textmd': ('-', 'bold'),
               'textsubscript': ('+', 'sub'), 'textsuperscript': ('+', 'sup')}
_DECLARATIONS = {'it': ('+', 'italic'), 'itshape': ('+', 'italic'),
                 'sl': ('+', 'italic'), 'slshape': ('+', 'italic'),
                 'em': ('~', 'italic'), 'bf': ('+', 'bold'), 'bfseries': ('+', 'bold'),
                 'sc': ('+', 'smallcaps'), 'scshape': ('+', 'smallcaps'),
                 'tt': ('+', 'mono'), 'ttfamily': ('+', 'mono'),
                 'rm': ('clear', None), 'rmfamily': ('clear', None),
                 'normalfont': ('clear', None), 'upshape': ('-', 'italic'),
                 'mdseries': ('-', 'bold')}


def _typo(s):
    """TeX ligatures for dashes and quotes (applied to plain text only)."""
    s = s.replace('---', '—').replace('--', '–')
    s = s.replace('``', '“').replace("''", '”')
    s = s.replace('`', '‘').replace("'", '’')
    return s


def _apply_accent(cmd, arg):
    arg = arg.replace('\\i', 'i').replace('\\j', 'j')
    if not arg:
        return _SPACING_ACCENT.get(cmd, '')
    return unicodedata.normalize('NFC', arg[0] + _COMBINING[cmd] + arg[1:])


def _set_flag(flags, mode, flag):
    if mode == 'clear':
        flags.difference_update({'italic', 'bold', 'smallcaps', 'mono'})
    elif mode == '+':
        flags.add(flag)
    elif mode == '-':
        flags.discard(flag)
    elif mode == '~':                        # \emph toggles italics
        flags.symmetric_difference_update({flag})


def _inline(s, flags, runs):
    """Recursive worker for parse_inline(); appends to `runs`."""
    i, n = 0, len(s)
    buf = []

    def flush():
        if buf:
            runs.append(('text', ''.join(buf), frozenset(flags)))
            buf.clear()

    while i < n:
        c = s[i]
        if c == '$':                         # inline math
            close = '$$' if s.startswith('$$', i) else '$'
            k = i + len(close)
            while True:
                k = s.find(close, k)
                if k == -1 or s[k - 1] != '\\':
                    break
                k += 1
            if k != -1:
                flush()
                runs.append(('math', s[i + len(close):k], frozenset(flags)))
                i = k + len(close)
                continue
            buf.append(c)
            i += 1
        elif c == '{':
            k = _match_brace(s, i)
            if k == -1:
                i += 1
                continue
            flush()
            _inline(s[i + 1:k], set(flags), runs)
            i = k + 1
        elif c == '}':
            i += 1
        elif c == '~':
            buf.append('\u00a0')
            i += 1
        elif c == '\\':
            i = _inline_command(s, i, flags, runs, buf, flush)
        else:
            m = re.compile(r'[^\\{}$~]+').match(s, i)
            buf.append(_typo(m.group(0)))
            i = m.end()
    flush()


def _read_arg(s, i):
    """Read one mandatory argument at s[i:] (after optional whitespace)."""
    i = _skip_ws(s, i)
    if i < len(s) and s[i] == '{':
        k = _match_brace(s, i)
        if k != -1:
            return s[i + 1:k], k + 1
    return None, i


def _inline_command(s, i, flags, runs, buf, flush):
    """Handle the control sequence at s[i] == '\\'. Returns the new index."""
    n = len(s)
    if i + 1 >= n:
        return n
    c = s[i + 1]
    if not c.isalpha():
        j = i + 2
        if c in _COMBINING and c not in ('u', 'v', 'H', 'c', 'k', 'r', 'b', 'd'):
            arg, k = _read_arg(s, j)
            if arg is None:
                if j < n and s[j] == '\\' and s[j + 1:j + 2] in ('i', 'j'):
                    arg, k = s[j:j + 2], j + 2
                elif j < n:
                    arg, k = s[j], j + 1
                else:
                    arg, k = '', j
            buf.append(_apply_accent(c, arg))
            return k
        if c in '&%$#_{}':
            buf.append(c)
        elif c == ' ':
            buf.append(' ')
        elif c in ',;:!':
            buf.append('\u2009' if c == ',' else ' ' if c != '!' else '')
        elif c == '\\':
            buf.append('\n')
        elif c == '(':                       # \( ... \)
            k = s.find('\\)', j)
            if k != -1:
                flush()
                runs.append(('math', s[j:k], frozenset(flags)))
                return k + 2
        return j

    m = re.compile(r'[A-Za-z]+\*?').match(s, i + 1)
    name = m.group(0).rstrip('*')
    j = _skip_ws(s, m.end())                 # TeX skips spaces after control words

    if name in _COMBINING:                   # \c c, \v{s}, \u a, \H o ...
        arg, k = _read_arg(s, j)
        if arg is None and j < n:
            arg, k = s[j], j + 1
        if arg is not None:
            buf.append(_apply_accent(name, arg))
            return k
    if name in _SPECIAL_LETTERS:
        buf.append(_SPECIAL_LETTERS[name])
        return j
    if name in _DECLARATIONS:
        flush()
        _set_flag(flags, *_DECLARATIONS[name])
        return j
    if name in _DROP_COMMANDS:
        return j
    if name in _ARG_FORMAT:
        arg, k = _read_arg(s, j)
        if arg is not None:
            flush()
            new = set(flags)
            mode, flag = _ARG_FORMAT[name]
            _set_flag(new, mode, flag)
            _inline(arg, new, runs)
            return k
        return j
    if name in ('label', 'hspace', 'vspace', 'phantom', 'hphantom', 'vphantom',
                'glsetup', 'index'):
        _, k = _read_arg(s, j)
        if name == 'hspace':
            buf.append(' ')
        return k
    if name in ('cite', 'citep', 'citet', 'ref'):
        arg, k = _read_arg(s, j)
        if arg is not None:
            buf.append('[' + arg + ']')
            return k
        return j
    if name in ('textcolor', 'colorbox'):
        _, k = _read_arg(s, j)               # colour name
        arg, k2 = _read_arg(s, k)
        if arg is not None:
            flush()
            _inline(arg, set(flags), runs)
            return k2
        return k
    if name in ('ljudge', 'judge', 'rightcomment', 'trailingcitation', 'rightcite'):
        arg, k = _read_arg(s, j)
        if arg is not None:
            flush()
            if name in ('rightcomment', 'trailingcitation', 'rightcite'):
                buf.append('\u2003')
            _inline(arg, set(flags), runs)
            if name == 'judge':
                buf.append('\u2009')
            return k
        return j
    # unknown command: keep the content of a following {argument}
    arg, k = _read_arg(s, j)
    if arg is not None:
        flush()
        _inline(arg, set(flags), runs)
        return k
    return j


def parse_inline(s, base_flags=()):
    """Turn a LaTeX fragment into runs: ('text', str, flags) or ('math', latex, flags)."""
    runs = []
    _inline(s, set(base_flags), runs)
    merged = []
    for r in runs:
        if merged and r[0] == 'text' and merged[-1][0] == 'text' and merged[-1][2] == r[2]:
            merged[-1] = ('text', merged[-1][1] + r[1], r[2])
        elif r[0] == 'text' and not r[1]:
            continue
        else:
            merged.append(r)
    return merged


# ----------------------------------------------------------------------------
# Word output helpers
# ----------------------------------------------------------------------------

class _RenderCtx:
    def __init__(self, xdoc, excnt, verbose=False):
        self.xdoc = xdoc
        self.excnt = excnt
        self.verbose = verbose


def _style_run(run, flags):
    if 'italic' in flags:
        run.font.italic = True
    if 'bold' in flags:
        run.font.bold = True
    if 'smallcaps' in flags:
        run.font.small_caps = True
    if 'sub' in flags:
        run.font.subscript = True
    if 'sup' in flags:
        run.font.superscript = True
    if 'mono' in flags:
        run.font.name = 'Courier New'


def _append_inline_omml(paragraph, omml):
    """Insert an inline OMML equation as a child of the paragraph (schema-valid)."""
    m = re.search(r'<m:oMath(?:\s[^>]*)?>.*</m:oMath>', omml, re.DOTALL)
    inner = m.group(0) if m else omml
    if 'xmlns:m' not in inner:
        inner = inner.replace('<m:oMath', '<m:oMath xmlns:m='
                              '"http://schemas.openxmlformats.org/officeDocument/2006/math"', 1)
    paragraph._p.append(parse_xml(inner))


def _add_inline(paragraph, latex, base_flags, rc):
    """Write a LaTeX fragment into a paragraph (text formatting + math -> OMML)."""
    latex = rc.xdoc.resolve_refs(latex, rc.excnt)
    for kind, payload, flags in parse_inline(latex, base_flags):
        if kind == 'text':
            _style_run(paragraph.add_run(payload), flags)
            continue
        omml = latex_to_omml(payload.strip())
        try:
            if not omml:
                raise ValueError('conversion failed')
            _append_inline_omml(paragraph, omml)
        except Exception as e:
            _style_run(paragraph.add_run(payload), flags)
            if rc.verbose:
                print(f"  ✗ Could not convert math in example (${payload}$): {e}")


def _tight(paragraph, before=0):
    pf = paragraph.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0


def _spacer(paragraph, points):
    from docx.enum.text import WD_LINE_SPACING
    pf = paragraph.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(points)


def _est_width(latex, rc):
    """Estimated rendered width (inches) of a LaTeX fragment."""
    chars = 0.0
    for kind, payload, _ in parse_inline(rc.xdoc.resolve_refs(latex, rc.excnt)):
        chars += len(payload) * (1.0 if kind == 'text' else EXPEX_MATH_CHAR_FACTOR)
    return chars * EXPEX_CHAR_WIDTH_IN


class _CellWriter:
    """Adds paragraphs and nested tables to a cell without stray blank paragraphs."""

    def __init__(self, cell):
        self.cell = cell
        self.fresh = cell.paragraphs[0]      # the empty paragraph every cell starts with
        self.free = None                     # empty paragraph left behind by add_table

    def paragraph(self, before=0):
        if self.fresh is not None:
            p, self.fresh = self.fresh, None
        elif self.free is not None:
            p, self.free = self.free, None
        else:
            p = self.cell.add_paragraph()
        _tight(p, before)
        return p

    def table(self, rows, cols):
        if self.fresh is not None:
            self.cell._tc.remove(self.fresh._p)
            self.fresh = None
        if self.free is not None:            # two tables in a row: keep a thin spacer
            _spacer(self.free, 3)
            self.free = None
        tbl = self.cell.add_table(rows, cols)
        self.free = self.cell.paragraphs[-1]
        return tbl

    def finish(self):
        if self.free is not None:
            _spacer(self.free, 2)
            self.free = None


def _configure_gloss_table(tbl):
    """Borderless table with a right-hand gap after each column."""
    tblPr = tbl._tbl.tblPr
    mar = OxmlElement('w:tblCellMar')
    for side, width in (('left', 0), ('right', int(EXPEX_WORD_GAP_IN * 1440))):
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:w'), str(width))
        el.set(qn('w:type'), 'dxa')
        mar.append(el)
    look = tblPr.find(qn('w:tblLook'))
    if look is not None:
        look.addprevious(mar)
    else:
        tblPr.append(mar)


def _no_wrap(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(OxmlElement('w:noWrap'))


def _write_text_block(cw, text, rc, before=0):
    for para in re.split(r'\n\s*\n', text):
        t = re.sub(r'\s+', ' ', para).strip()
        if t:
            _add_inline(cw.paragraph(before), t, (), rc)


def _render_gloss(cw, gloss, avail_in, rc):
    """Render a gloss as one or more aligned word-column tables (wrapping long glosses)."""
    if gloss.preamble:
        _add_inline(cw.paragraph(), gloss.preamble, (), rc)
    letters, cols = gloss.letters, gloss.columns
    if cols and letters:
        widths = []
        for ci, col in enumerate(cols):
            w = 0.0
            for l in letters:
                txt = col['cells'].get(l, '')
                if ci == 0 and l == 'a' and gloss.judge:
                    txt = gloss.judge + txt
                w = max(w, _est_width(txt, rc))
            widths.append(max(w, 0.15) + EXPEX_WORD_GAP_IN)

        chunks, cur, cur_w = [], [], 0.0
        for ci, w in enumerate(widths):      # greedy line filling
            if cur and (ci in gloss.breaks or cur_w + w > avail_in):
                chunks.append(cur)
                cur, cur_w = [], 0.0
            cur.append(ci)
            cur_w += w
        if cur:
            chunks.append(cur)

        for chunk in chunks:
            tbl = cw.table(len(letters), len(chunk))
            _configure_gloss_table(tbl)
            # Auto-fit layout: the grid widths are only preferences, so a column whose text
            # is wider than estimated grows (noWrap) instead of breaking the word.
            tbl.autofit = True
            for ti, ci in enumerate(chunk):
                tbl.columns[ti].width = Inches(widths[ci])
            for ti, ci in enumerate(chunk):
                for ri, l in enumerate(letters):
                    cell = tbl.cell(ri, ti)
                    cell.width = Inches(widths[ci])
                    _no_wrap(cell)
                    txt = cols[ci]['cells'].get(l, '')
                    if ci == 0 and l == 'a' and gloss.judge:
                        txt = gloss.judge + txt
                    p = cell.paragraphs[0]
                    _tight(p)
                    _add_inline(p, txt, gloss.line_flags.get(l, ()), rc)
    for k, ft in enumerate(gloss.ft):
        if ft:
            _add_inline(cw.paragraph(before=2), ft, (), rc)


def _flatten_example(node, prefix, rows):
    """Turn an ExNode into table rows: {'num', 'labels', 'blocks'}."""
    if node.kind != 'pex' or not node.parts:
        rows.append({'num': '', 'labels': list(prefix),
                     'blocks': node.blocks if node.kind != 'pex' else node.preamble})
        return
    if node.preamble:
        rows.append({'num': '', 'labels': list(prefix), 'blocks': node.preamble})
    for part in node.parts:
        own = list(prefix) + [part.label_text]
        cont = list(prefix) + ['']
        cur, used = [], False
        for blk in part.blocks:
            if blk[0] != 'sub':
                cur.append(blk)
                continue
            if cur:
                rows.append({'num': '', 'labels': list(cont if used else own), 'blocks': cur})
                used, cur = True, []
            tmp = []
            _flatten_example(blk[1], cont, tmp)
            if tmp and not used:
                tmp[0]['labels'][len(prefix)] = part.label_text
                used = True
            rows.extend(tmp)
        if cur or not used:
            rows.append({'num': '', 'labels': list(cont if used else own), 'blocks': cur})


def render_expex_example(doc, node, xdoc, verbose=False):
    """Add one parsed expex example to the Word document as a borderless table."""
    rc = _RenderCtx(xdoc, node.excnt, verbose)
    section = doc.sections[-1]
    text_w = (section.page_width - section.left_margin - section.right_margin) / 914400.0

    rows = []
    _flatten_example(node, [], rows)
    rows[0]['num'] = node.number_text
    depth = max(len(r['labels']) for r in rows)
    ncols = 2 + depth

    num_w = max([0.45] + [_est_width(r['num'], rc) + 0.15 for r in rows])
    label_w = []
    for lvl in range(depth):
        label_w.append(max([0.3] + [_est_width(r['labels'][lvl], rc) + 0.15
                                    for r in rows if len(r['labels']) > lvl]))
    widths = [num_w] + label_w + [max(text_w - num_w - sum(label_w), 1.5)]

    _spacer(doc.add_paragraph(), EXPEX_SPACER_PT)
    tbl = doc.add_table(rows=len(rows), cols=ncols)
    tbl.autofit = False
    for ci, w in enumerate(widths):
        tbl.columns[ci].width = Inches(w)
    for ri, row in enumerate(rows):
        tr = tbl.rows[ri]._tr
        tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))
        cells = list(tbl.rows[ri].cells)
        for ci, w in enumerate(widths):
            cells[ci].width = Inches(w)

        def put(cell, text):
            p = cell.paragraphs[0]
            _tight(p)
            if text:
                _add_inline(p, text, (), rc)

        put(cells[0], row['num'])
        for li, lab in enumerate(row['labels']):
            put(cells[1 + li], lab)
        k = len(row['labels'])
        content = cells[1 + k]
        span_w = sum(widths[1 + k:])
        if 1 + k < ncols - 1:
            content = content.merge(cells[ncols - 1])
        cw = _CellWriter(content)
        for kind, payload in row['blocks']:
            if kind == 'text':
                _write_text_block(cw, payload, rc)
            elif kind == 'gloss':
                _render_gloss(cw, payload, span_w - 0.2, rc)
        cw.finish()

    spacer = doc.add_paragraph()
    _spacer(spacer, EXPEX_SPACER_PT)
    if verbose:
        print(f"✓ Added example {node.number_text or '(unnumbered)'} "
              f"({len(rows)} row{'s' if len(rows) != 1 else ''})")
    return tbl


# ============================================================================
# MAIN PIPELINE
# ============================================================================

def latex_to_word(latex_file, output_file=None, verbose=True):
    """
    Convert a LaTeX file to a Word document with LaTeX equations rendered as OMML.
    
    This is the main entry point for the conversion pipeline.
    
    Args:
        latex_file (str): Path to the LaTeX file to convert
        output_file (str): Path for the output Word document (default: latex_file with .docx extension)
        verbose (bool): If True, print progress information
    
    Returns:
        str: Path to the created Word document
    """
    # Determine output path if not provided
    if output_file is None:
        base_path = Path(latex_file).stem
        output_file = f"{base_path}.docx"
    
    if verbose:
        print(f"Reading LaTeX file: {latex_file}")
    
    # Read the LaTeX file
    with open(latex_file, 'r', encoding='utf-8') as f:
        latex_content = f.read()
    
    if verbose:
        print(f"Removing comments...")
    
    # Remove LaTeX comments
    latex_content = remove_latex_comments(latex_content)
    
    # expex: \\lingset settings live in the preamble, so read them before it is skipped
    expex_settings = collect_expex_preamble_settings(latex_content)
    
    if verbose:
        print(f"Skipping preamble (if any)...")
    
    # Skip preamble
    latex_content = skip_latex_preamble(latex_content)
    
    # expex: cut out numbered examples (before equation extraction, so that math
    # inside glosses is handled by the example renderer and not by the equation list)
    latex_content, expex_doc = extract_expex_examples(latex_content, expex_settings, verbose=verbose)
    if verbose and expex_doc.examples:
        print(f"Found {len(expex_doc.examples)} expex examples")
    
    # Strip \resizebox commands BEFORE extracting equations
    # This is critical because \resizebox{...}{...}{$...$} will confuse the equation extractor
    if verbose and r'\resizebox' in latex_content:
        print(f"Stripping \\resizebox commands...")
    
    def strip_resizebox_early(text):
        """Strip \resizebox{...}{...}{content} and remove $ delimiters."""
        iteration = 0
        while r'\resizebox{' in text:
            iteration += 1
            if iteration > 100:
                break
            match = re.search(r'\\resizebox\{', text)
            if not match:
                break
            
            pos = match.end() - 1
            for _ in range(2):
                if pos >= len(text) or text[pos] != '{':
                    break
                brace_count = 1
                pos += 1
                while pos < len(text) and brace_count > 0:
                    if text[pos] == '{' and (pos == 0 or text[pos-1] != '\\'):
                        brace_count += 1
                    elif text[pos] == '}' and (pos == 0 or text[pos-1] != '\\'):
                        brace_count -= 1
                    pos += 1
            
            if pos >= len(text) or text[pos] != '{':
                break
            
            start_content = pos + 1
            brace_count = 1
            pos = start_content
            while pos < len(text) and brace_count > 0:
                if text[pos] == '{' and (pos == 0 or text[pos-1] != '\\'):
                    brace_count += 1
                elif text[pos] == '}' and (pos == 0 or text[pos-1] != '\\'):
                    brace_count -= 1
                pos += 1
            
            if brace_count != 0:
                break
            
            content_inner = text[start_content:pos-1].strip()
            if content_inner.startswith('$') and content_inner.endswith('$'):
                content_inner = content_inner[1:-1]
            
            text = text[:match.start()] + content_inner + text[pos:]
        
        return text
    
    latex_content = strip_resizebox_early(latex_content)
    
    if verbose:
        print(f"Extracting LaTeX equations...")
    
    # Extract equations
    equations_from_tex = extract_latex_equations(latex_content)
    
    if verbose:
        print(f"\nFound {len(equations_from_tex)} LaTeX equations\n")
        for i, (eq, is_display, label) in enumerate(equations_from_tex, 1):
            mode = "DISPLAY" if is_display else "INLINE"
            eq_preview = eq.replace('\n', ' ')[:60]
            label_str = f" [Label: {label}]" if label else ""
            print(f"{i}. [{mode}]{label_str} {eq_preview}{'...' if len(eq) > 60 else ''}")
    
    # Convert all extracted equations to OMML
    tex_omml_equations = convert_equations_to_omml(equations_from_tex, verbose=verbose)
    
    # Create the Word document. Equations that failed to convert (e.g. texmath is
    # missing) are kept as plain LaTeX text rather than blocking the whole document.
    n_converted = sum(1 for omml, *_ in tex_omml_equations if omml)
    output_path = create_word_doc_from_latex(
        latex_content, 
        tex_omml_equations,
        output_path=output_file,
        verbose=verbose,
        expex_doc=expex_doc
    )
    if verbose:
        if not equations_from_tex:
            print(f"\n✓ Successfully created Word document (no equations to convert)")
        elif n_converted == len(equations_from_tex):
            print(f"\n✓ Successfully created Word document with {n_converted} equations")
        else:
            print(f"\n✓ Successfully created Word document ({n_converted}/{len(equations_from_tex)} "
                  f"equations converted; the rest were kept as plain LaTeX text)")
    return output_path


# ============================================================================
# SCRIPT ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    import sys
    import argparse

    parser = argparse.ArgumentParser(
        prog="latex_to_word.py",
        description="Convert a LaTeX (.tex) file to a Word (.docx) document, "
                    "with equations rendered as OMML and expex examples as tables.",
    )
    parser.add_argument("latex_file", nargs="?", default=None,
                        help="path to the .tex file to convert "
                             "(if omitted, you will be prompted)")
    parser.add_argument("-o", "--output", dest="output_file", default=None,
                        help="output .docx path (default: <latex_file>.docx)")
    parser.add_argument("-v", "--verbose", action="store_true",
                        help="print progress information while converting")
    args = parser.parse_args()

    verbose_tag = args.verbose
    latex_file = args.latex_file or input("Enter the path to the LaTeX file: ").strip()
    output_file = args.output_file

    # Resolve the path (handle both relative and absolute paths)
    latex_path = Path(latex_file)
    
    # If relative path, check current directory first, then parent directories
    if not latex_path.is_absolute():
        # Try current directory
        if not latex_path.exists():
            # Try common parent directories
            alt_paths = [
                Path.cwd() / latex_file,
                Path.cwd().parent / latex_file,
                Path.home() / latex_file,
            ]
            for alt_path in alt_paths:
                if alt_path.exists():
                    latex_path = alt_path
                    break
    
    # Check if file exists
    if not latex_path.exists():
        print(f"❌ Error: File '{latex_file}' not found")
        print(f"   Checked: {Path.cwd() / latex_file}")
        sys.exit(1)
    
    # Convert to absolute path for processing
    latex_file = str(latex_path.resolve())

    try:
        result = latex_to_word(latex_file, output_file, verbose=verbose_tag)
        if result:
            print(f"\n✅ Conversion completed successfully!")
            print(f"Output file: {result}")
        else:
            print("\n❌ Conversion failed")
    except Exception as e:
        print(f"\n❌ Error during conversion: {e}")
        import traceback
        traceback.print_exc()
