// SPDX-License-Identifier: Apache-2.0
//
// editor.ts - dependency-free syntax-highlighted code editor component.
//
// Uses the classic "highlighted overlay" technique: a transparent <textarea>
// sits exactly on top of a <pre><code> layer. On each input event the textarea
// value is HTML-escaped, tokenized, and written back into the highlight layer.
// Scroll position is mirrored so the two layers track each other perfectly.

type Mode = "sql" | "python";

// ---------------------------------------------------------------------------
// Token colours are CSS classes defined in demo.css.
// ---------------------------------------------------------------------------

const SQL_KEYWORDS = new Set(
  (
    "SELECT FROM WHERE GROUP BY ORDER HAVING JOIN LEFT RIGHT INNER OUTER ON AS AND OR NOT NULL " +
    "LIMIT DISTINCT UNION INSERT UPDATE DELETE INTO VALUES CREATE TABLE WITH CASE WHEN THEN " +
    "ELSE END SUM AVG COUNT MIN MAX DESC ASC ALL CROSS FULL NATURAL EXCEPT INTERSECT TRUE FALSE"
  )
    .split(" ")
    .filter(Boolean),
);

const PYTHON_KEYWORDS = new Set(
  (
    "def return import from for in if elif else while class with as None True False " +
    "and or not lambda print pass break continue raise try except finally yield async await"
  )
    .split(" ")
    .filter(Boolean),
);

// ---------------------------------------------------------------------------
// HTML-escape helpers (ASCII-safe, no fancy chars).
// ---------------------------------------------------------------------------

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// ---------------------------------------------------------------------------
// Tokenizer - returns highlighted HTML.
// ---------------------------------------------------------------------------

function highlight(raw: string, mode: Mode): string {
  // We tokenize character by character via a single pass regex.
  // Order matters: comments > strings > numbers > keywords > plain text.
  if (mode === "sql") {
    return highlightSql(raw);
  }
  return highlightPython(raw);
}

// Shared span wrappers.
function span(cls: string, content: string): string {
  return `<span class="${cls}">${content}</span>`;
}

// SQL tokenizer.
function highlightSql(raw: string): string {
  // Pattern groups (in order of priority):
  //   1. line comment  --...
  //   2. single-quoted string  '...' (no escape handling for simplicity)
  //   3. double-quoted identifier  "..."
  //   4. number literal
  //   5. identifier / keyword (letters + _)
  //   6. anything else (operators, punctuation, whitespace)
  const TOKEN = /--[^\n]*|'[^']*'|"[^"]*"|\b\d+(?:\.\d+)?\b|[A-Za-z_][A-Za-z0-9_]*|[\s\S]/g;
  let out = "";
  let m: RegExpExecArray | null;
  while ((m = TOKEN.exec(raw)) !== null) {
    const tok = m[0];
    if (tok.startsWith("--")) {
      out += span("tok-comment", escapeHtml(tok));
    } else if (tok.startsWith("'") || tok.startsWith('"')) {
      out += span("tok-string", escapeHtml(tok));
    } else if (/^\d/.test(tok)) {
      out += span("tok-number", escapeHtml(tok));
    } else if (/^[A-Za-z_]/.test(tok)) {
      const upper = tok.toUpperCase();
      if (SQL_KEYWORDS.has(upper)) {
        out += span("tok-keyword", escapeHtml(tok));
      } else {
        out += escapeHtml(tok);
      }
    } else {
      out += escapeHtml(tok);
    }
  }
  return out;
}

// Python tokenizer.
function highlightPython(raw: string): string {
  // Pattern groups:
  //   1. block comment  #...
  //   2. triple double-quoted string """..."""
  //   3. triple single-quoted string '''...'''
  //   4. double-quoted string "..."
  //   5. single-quoted string '...'
  //   6. number literal
  //   7. identifier / keyword
  //   8. anything else
  const TOKEN =
    /#[^\n]*|"""[\s\S]*?"""|'''[\s\S]*?'''|"[^"\n]*"|'[^'\n]*'|\b\d+(?:\.\d+)?\b|[A-Za-z_][A-Za-z0-9_]*|[\s\S]/g;
  let out = "";
  let m: RegExpExecArray | null;
  while ((m = TOKEN.exec(raw)) !== null) {
    const tok = m[0];
    if (tok.startsWith("#")) {
      out += span("tok-comment", escapeHtml(tok));
    } else if (
      tok.startsWith('"""') ||
      tok.startsWith("'''") ||
      tok.startsWith('"') ||
      tok.startsWith("'")
    ) {
      out += span("tok-string", escapeHtml(tok));
    } else if (/^\d/.test(tok)) {
      out += span("tok-number", escapeHtml(tok));
    } else if (/^[A-Za-z_]/.test(tok)) {
      if (PYTHON_KEYWORDS.has(tok)) {
        out += span("tok-keyword", escapeHtml(tok));
      } else {
        out += escapeHtml(tok);
      }
    } else {
      out += escapeHtml(tok);
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

export interface CodeEditorHandle {
  /** The root element to insert into the DOM. */
  el: HTMLElement;
  /** Return the current editor text. */
  getValue(): string;
  /** Programmatically set the editor text and re-highlight. */
  setValue(v: string): void;
  /** Switch syntax mode and re-highlight. */
  setMode(m: Mode): void;
}

export interface CodeEditorOptions {
  value: string;
  mode: Mode;
  onInput?: (v: string) => void;
}

/**
 * Create a syntax-highlighted code editor backed by a plain <textarea>.
 * Returns a handle with the wrapper element and imperative methods.
 */
export function createCodeEditor(opts: CodeEditorOptions): CodeEditorHandle {
  let currentMode: Mode = opts.mode;

  // Container - positioned so the highlight layer and textarea can overlap.
  const container = document.createElement("div");
  container.className = "code-editor";

  // Highlight layer: a <pre> with a nested <code> element.
  const pre = document.createElement("pre");
  pre.className = "code-editor__highlight";
  pre.setAttribute("aria-hidden", "true");
  const code = document.createElement("code");
  pre.appendChild(code);

  // Textarea - transparent background, sits on top of the highlight layer.
  const textarea = document.createElement("textarea");
  textarea.className = "code-editor__textarea";
  textarea.value = opts.value;
  textarea.spellcheck = false;
  // autocorrect / autocapitalize are non-standard but help on mobile Safari.
  // autocomplete is set via setAttribute to avoid AutoFill type constraints.
  textarea.setAttribute("autocomplete", "off");
  textarea.setAttribute("autocorrect", "off");
  textarea.setAttribute("autocapitalize", "off");

  container.appendChild(pre);
  container.appendChild(textarea);

  // Mirror content into the highlight layer.
  function sync(): void {
    // Append a trailing newline so the last line in the textarea is always
    // tall enough (browsers eat the final newline in <pre> otherwise).
    code.innerHTML = highlight(textarea.value, currentMode) + "\n";
  }

  // Mirror scroll so the highlight layer tracks the textarea exactly.
  function syncScroll(): void {
    pre.scrollTop = textarea.scrollTop;
    pre.scrollLeft = textarea.scrollLeft;
  }

  textarea.addEventListener("input", () => {
    sync();
    opts.onInput?.(textarea.value);
  });

  textarea.addEventListener("scroll", syncScroll);

  // Initial render.
  sync();

  return {
    el: container,

    getValue(): string {
      return textarea.value;
    },

    setValue(v: string): void {
      textarea.value = v;
      sync();
    },

    setMode(m: Mode): void {
      currentMode = m;
      sync();
    },
  };
}
