(() => {
  'use strict';

  // === 1) Config — extend these as needed ===
  const SQL_KEYWORDS = [
    'SELECT','FROM','WHERE','GROUP','BY','HAVING','ORDER','ASC','DESC',
    'JOIN','LEFT','RIGHT','FULL','INNER','OUTER','CROSS','APPLY','ON','AS',
    'TOP','DISTINCT','UNION','ALL','EXCEPT','INTERSECT',
    'INSERT','INTO','VALUES','UPDATE','SET','DELETE',
    'CREATE','ALTER','DROP','TABLE','VIEW','PROCEDURE','FUNCTION','TRIGGER','INDEX',
    'PRIMARY','KEY','FOREIGN','REFERENCES','DEFAULT','IDENTITY',
    'DECLARE','BEGIN','END','COMMIT','ROLLBACK','TRANSACTION','WITH','NOLOCK',
    'OFFSET','FETCH','OVER','PARTITION',
    'CASE','WHEN','THEN','ELSE','END',
    'AND','OR','NOT','NULL','IS','LIKE','IN','EXISTS','BETWEEN'
  ];

  const SQL_FUNCTIONS = [
    // Common T-SQL functions (add/remove as needed)
    'SUBSTRING','LTRIM','RTRIM','ISNULL','COALESCE','REPLACE',
    'LEFT','RIGHT','LEN','DATALENGTH',
    'CAST','CONVERT','TRY_CAST','TRY_CONVERT',
    'UPPER','LOWER',
    'DATEADD','DATEDIFF','DATENAME','DATEPART',
    'GETDATE','GETUTCDATE','SYSDATETIME','NEWID',
    'ROW_NUMBER','COUNT','SUM','AVG','MIN','MAX','CLng','CInt','CDbl','CStr'
  ];

  const CLAUSES = [
    'LEFT OUTER JOIN','RIGHT OUTER JOIN','FULL OUTER JOIN',
    'LEFT JOIN','RIGHT JOIN','FULL JOIN',
    'INNER JOIN','OUTER APPLY','CROSS APPLY','JOIN',
    'GROUP BY','ORDER BY','UNION ALL','UNION','EXCEPT','INTERSECT',
    'SELECT','FROM','WHERE','HAVING'
  ].sort((a, b) => b.length - a.length); // longest first

  // Precompile clause regexes to detect at current index, ignoring case & flexible spaces
  const CLAUSE_PATTERNS = CLAUSES.map(c => ({
    clause: c,
    rx: new RegExp('^\\b' + c.split(' ').map(w => w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('\\s+') + '\\b', 'i')
  }));

  // === 2) Helpers ===
  function escapeHtml(s) {
    return s.replace(/[&<>"']/g, ch => (
      { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]
    ));
  }

  // Insert linebreaks (before major clauses) & safe comma-breaks, respecting strings/comments
  function prettySql(sql) {
    // Normalize newlines to \n
    const src = String(sql).replace(/\r\n?/g, '\n');

    let out = '';
    let i = 0;

    // state
    let inSgl = false, inDbl = false, inBrk = false;
    let inLine = false, inBlock = false;

    while (i < src.length) {
      const ch = src[i];
      const next = src[i + 1];

      // Handle end of line comment
      if (inLine) {
        out += ch;
        if (ch === '\n') inLine = false;
        i++;
        continue;
      }

      // Handle block comment
      if (inBlock) {
        out += ch;
        if (ch === '*' && next === '/') {
          out += '/';
          i += 2;
          inBlock = false;
        } else {
          i++;
        }
        continue;
      }

      // Start comments (outside strings/brackets)
      if (!inSgl && !inDbl && !inBrk) {
        if (ch === '-' && next === '-') {
          out += ch + next;
          i += 2;
          inLine = true;
          continue;
        }
        if (ch === '/' && next === '*') {
          out += ch + next;
          i += 2;
          inBlock = true;
          continue;
        }
      }

      // Strings / bracket identifiers
      if (!inDbl && !inBrk && ch === "'" ) {
        // Toggle single-quote string, handle escaped ''
        out += ch; i++;
        inSgl = !inSgl;
        // stay in string until closing unescaped '
        while (inSgl && i < src.length) {
          const c = src[i], n = src[i + 1];
          out += c; i++;
          if (c === "'" && n === "'") { out += n; i++; } // escaped ''
          else if (c === "'") inSgl = false;
        }
        continue;
      }
      if (!inSgl && !inBrk && ch === '"') {
        out += ch; i++;
        inDbl = !inDbl;
        while (inDbl && i < src.length) {
          const c = src[i];
          out += c; i++;
          if (c === '"') inDbl = false;
        }
        continue;
      }
      if (!inSgl && !inDbl && ch === '[') {
        out += ch; i++;
        inBrk = true;
        while (inBrk && i < src.length) {
          const c = src[i];
          out += c; i++;
          if (c === ']') inBrk = false;
        }
        continue;
      }

      // Outside strings/comments/brackets → handle clauses & commas
      if (!inSgl && !inDbl && !inBrk) {
        // Try to match a clause here
        const slice = src.slice(i);
        let matched = null;
        for (const { clause, rx } of CLAUSE_PATTERNS) {
          const m = rx.exec(slice);
          if (m && m.index === 0) {
            matched = { text: m[0], clause };
            break;
          }
        }
        if (matched) {
          // Insert line break before clause (unless at start or just after newline)
          if (out.length && out[out.length - 1] !== '\n') {
            out = out.replace(/[ \t]+$/, ''); // trim trailing spaces
            out += '\n';
          }
          out += matched.clause.toUpperCase();
          i += matched.text.length;
          continue;
        }

        // Safe comma break: add newline after comma not followed by ')'
        if (ch === ',') {
          out += ',';
          // look ahead (skip spaces/tabs) but not newlines
          let j = i + 1;
          while (j < src.length && (src[j] === ' ' || src[j] === '\t')) j++;
          if (src[j] !== ')') {
            out += '\n  ';
          } else {
            out += ' ';
          }
          i++;
          continue;
        }
      }

      // Default: copy char
      out += ch;
      i++;
    }

    // Collapse excessive spaces per line and add simple indentation based on parens
    const IND = '  ';
    const lines = out.split('\n').map(s => s.trim().replace(/\s{2,}/g, ' '));
    let depth = 0;
    const indented = lines.map(line => {
      // reduce indent if line starts with closing paren
      const leadingClose = /^\)/.test(line) ? 1 : 0;
      const indentLevel = Math.max(0, depth - leadingClose);
      const outLine = (line ? IND.repeat(indentLevel) + line : '');
      // update depth from paren counts (approx; ignores parens in strings—which we avoided in scanning for breaks)
      const opens = (line.match(/\(/g) || []).length;
      const closes = (line.match(/\)/g) || []).length;
      depth = Math.max(0, depth + opens - closes);
      return outLine;
    });

    return indented.join('\n').replace(/\s+\n/g, '\n').trim();
  }

  // === 3) Highlighter ===
  function buildHighlighter() {
    const fnAlt = SQL_FUNCTIONS.map(x => x.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|');
    const kwAlt = SQL_KEYWORDS.map(x => x.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|');

    const rx = new RegExp([
      // Comments & strings come first
      '(?<lineComment>--[^\\n\\r]*)',
      '(?<blockComment>/\\*[\\s\\S]*?\\*/)',
      "(?<string>N?'(?:''|[^'])*')",
      '(?<bracketIdent>\\[[^\\]]+\\])',
      '(?<quotedIdent>"[^"]*")',
      '(?<number>\\b\\d+(?:\\.\\d+)?\\b)',
      `(?<function>\\b(?:${fnAlt})\\b)(?=\\s*\\()`,
      `(?<keyword>\\b(?:${kwAlt})\\b)`,
      '(?<operator>\\*|[=<>!]+|[()+\\-/%.,;])',
      '(?<boolean>\\bTRUE\\b|\\bFALSE\\b)'
    ].join('|'), 'gim');

    return function highlightSql(sql) {
      let out = '';
      let last = 0;
      for (const m of sql.matchAll(rx)) {
        const i = m.index;
        out += escapeHtml(sql.slice(last, i));
        const g = m.groups || {};
        const txt = m[0];
        const esc = escapeHtml(txt);
        if (g.lineComment || g.blockComment) out += `<span class="sql-comment">${esc}</span>`;
        else if (g.string) out += `<span class="sql-string">${esc}</span>`;
        else if (g.bracketIdent || g.quotedIdent) out += `<span class="sql-identifier">${esc}</span>`;
        else if (g.number) out += `<span class="sql-number">${esc}</span>`;
        else if (g.function) out += `<span class="sql-function">${esc}</span>`;
        else if (g.keyword) out += `<span class="sql-keyword">${esc.toUpperCase()}</span>`;
        else if (g.operator) out += `<span class="sql-operator">${esc}</span>`;
        else if (g.boolean) out += `<span class="sql-boolean">${esc}</span>`;
        else out += esc;
        last = i + txt.length;
      }
      out += escapeHtml(sql.slice(last));
      return out;
    };
  }

  const highlightSql = buildHighlighter();

  // === 4) Runner: find blocks and process ===
  function processBlock(codeEl) {
    const raw = codeEl.textContent;
    // const formatted = prettySql(raw);            // Commented out as I want to preserve original formatting
    // codeEl.innerHTML = highlightSql(formatted);  // Commented out as I want to preserve original formatting
    
    codeEl.innerHTML = highlightSql(raw);
  }

  function applyToAll() {
    const selectors = [
      'pre.sql > code',           // recommended
      'code.language-sql',        // alt
      'code[data-language="sql"]' // alt
    ];
    const seen = new Set();
    selectors.forEach(sel => {
      document.querySelectorAll(sel).forEach(el => {
        if (!seen.has(el)) {
          seen.add(el);
          processBlock(el);
        }
      });
    });
  }

  document.addEventListener('DOMContentLoaded', applyToAll);

  // Optional: expose small API
  window.SqlFormatterLite = {
    format: (sqlText) => highlightSql(prettySql(sqlText)),
    refresh: applyToAll
  };
})();