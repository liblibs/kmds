(() => {
  'use strict';

  // 1) Config
  const VBA_KEYWORDS = [
    'Dim','Set','If','Then','Else','ElseIf','End If','For','Next','To','Step',
    'Do','Loop','While','Wend','Function','Sub','End Function','End Sub',
    'Exit Function','Exit Sub','With','End With','Select','Case','End Select',
    'Public','Private','Const','Static','Option','Explicit','ByVal','ByRef'
  ];

  const VBA_FUNCTIONS = [
    'MsgBox','InputBox','CStr','CLng','CInt','CDbl','Trim','LTrim','RTrim',
    'Len','Mid','Left','Right','Replace','Date','Now','Time','Format','Round'
  ];

  const rx = new RegExp([
    "(?<comment>'[^\\n\\r]*)", // single-line comment
    '(?<string>"(?:[^"]|"")*")', // string literal
    '(?<number>\\b\\d+(?:\\.\\d+)?\\b)', // numbers
    `(?<function>\\b(?:${VBA_FUNCTIONS.join('|')})\\b)(?=\\s*\\()`, // functions
    `(?<keyword>\\b(?:${VBA_KEYWORDS.join('|')})\\b)`, // keywords
    '(?<operator>[+\\-*/&=<>]+)', // operators
    '(?<boolean>\\bTrue\\b|\\bFalse\\b)' // booleans
  ].join('|'), 'gi');

  function escapeHtml(s) {
    return s.replace(/[&<>\"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
  }

  function highlightVBA(code) {
    let out = '';
    let last = 0;
    for (const m of code.matchAll(rx)) {
      const i = m.index;
      out += escapeHtml(code.slice(last, i));
      const g = m.groups || {};
      const txt = escapeHtml(m[0]);
      if (g.comment) out += `<span class="vba-comment">${txt}</span>`;
      else if (g.string) out += `<span class="vba-string">${txt}</span>`;
      else if (g.number) out += `<span class="vba-number">${txt}</span>`;
      else if (g.function) out += `<span class="vba-function">${txt}</span>`;
      else if (g.keyword) out += `<span class="vba-keyword">${txt}</span>`;
      else if (g.operator) out += `<span class="vba-operator">${txt}</span>`;
      else if (g.boolean) out += `<span class="vba-boolean">${txt}</span>`;
      else out += txt;
      last = i + m[0].length;
    }
    out += escapeHtml(code.slice(last));
    return out;
  }

  function processBlock(el) {
    el.innerHTML = highlightVBA(el.textContent);
  }

  function applyToAll() {
    document.querySelectorAll('pre.vba > code, code.language-vba, code[data-language="vba"]').forEach(processBlock);
  }

  document.addEventListener('DOMContentLoaded', applyToAll);

  window.VbaHighlighterLite = { highlight: highlightVBA, refresh: applyToAll };
})();
