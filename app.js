const TABLE_STYLES = {
  simpleClean: {
    label: 'シンプル1: クリーン',
    borderWidth: 1,
    stripe: false,
    roundedCells: false,
    shadow: false,
    outerFrame: false,
  },
  simpleStriped: {
    label: 'シンプル2: ストライプ',
    borderWidth: 1,
    stripe: true,
    roundedCells: false,
    shadow: false,
    outerFrame: false,
  },
  simpleDense: {
    label: 'シンプル3: 高密度',
    borderWidth: 0.75,
    stripe: false,
    roundedCells: false,
    shadow: false,
    outerFrame: false,
  },
  cardSoft: {
    label: '装飾1: カード風',
    borderWidth: 0,
    stripe: true,
    roundedCells: true,
    shadow: true,
    outerFrame: false,
  },
  modernGlass: {
    label: '装飾2: モダン',
    borderWidth: 0,
    stripe: false,
    roundedCells: true,
    shadow: true,
    outerFrame: false,
  },
};

const INPUT_FORMATS = {
  auto: 'auto',
  csv: 'csv',
  markdown: 'markdown',
};

let selectedInputFormat = INPUT_FORMATS.auto;

const COLOR_THEMES = {
  grayMist: {
    label: 'グレー1: Mist',
    bg: '#ffffff',
    text: '#111827',
    border: '#d1d5db',
    headerBg: '#f3f4f6',
    headerText: '#111827',
    stripeBg: '#f9fafb',
    firstColBg: '#f8fafc',
    accent: '#6b7280',
  },
  graySlate: {
    label: 'グレー2: Slate',
    bg: '#ffffff',
    text: '#1f2937',
    border: '#cbd5e1',
    headerBg: '#e2e8f0',
    headerText: '#0f172a',
    stripeBg: '#f1f5f9',
    firstColBg: '#eef2f7',
    accent: '#475569',
  },
  blueSky: {
    label: '青1: Sky',
    bg: '#ffffff',
    text: '#0b2541',
    border: '#bfd7ea',
    headerBg: '#dbeafe',
    headerText: '#0c4a6e',
    stripeBg: '#eff6ff',
    firstColBg: '#e0f2fe',
    accent: '#0284c7',
  },
  blueNavy: {
    label: '青2: Navy',
    bg: '#ffffff',
    text: '#1e293b',
    border: '#c7d2fe',
    headerBg: '#e0e7ff',
    headerText: '#1e1b4b',
    stripeBg: '#eef2ff',
    firstColBg: '#e2e8f0',
    accent: '#4338ca',
  },
  greenLeaf: {
    label: '緑: Leaf',
    bg: '#ffffff',
    text: '#1f2937',
    border: '#b7e4c7',
    headerBg: '#dcfce7',
    headerText: '#14532d',
    stripeBg: '#f0fdf4',
    firstColBg: '#e8f5e9',
    accent: '#16a34a',
  },
};

const els = {
  csvFile: document.getElementById('csvFile'),
  formatAutoBtn: document.getElementById('formatAutoBtn'),
  formatCsvBtn: document.getElementById('formatCsvBtn'),
  formatMdBtn: document.getElementById('formatMdBtn'),
  formatHint: document.getElementById('formatHint'),
  csvText: document.getElementById('csvText'),
  aspectRatio: document.getElementById('aspectRatio'),
  dpi: document.getElementById('dpi'),
  fontFamily: document.getElementById('fontFamily'),
  fontSize: document.getElementById('fontSize'),
  currentFontSize: document.getElementById('currentFontSize'),
  hasHeaderRow: document.getElementById('hasHeaderRow'),
  hasHeaderCol: document.getElementById('hasHeaderCol'),
  headerAlign: document.getElementById('headerAlign'),
  bodyAlign: document.getElementById('bodyAlign'),
  tableStyle: document.getElementById('tableStyle'),
  colorTheme: document.getElementById('colorTheme'),
  renderBtn: document.getElementById('renderBtn'),
  downloadBtn: document.getElementById('downloadBtn'),
  canvas: document.getElementById('previewCanvas'),
};

function initSelect(select, source) {
  Object.entries(source).forEach(([value, conf], idx) => {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = conf.label;
    if (idx === 0) option.selected = true;
    select.append(option);
  });
}

function normalizeRows(rows) {
  if (!rows.length) return [];
  const maxCols = rows.reduce((max, r) => Math.max(max, r.length), 0);
  return rows.map((r) => [...r, ...Array(maxCols - r.length).fill('')]);
}

function parseDelimited(text, delimiter) {
  const rows = [];
  let row = [];
  let value = '';
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    const next = text[i + 1];

    if (ch === '"') {
      if (inQuotes && next === '"') {
        value += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === delimiter && !inQuotes) {
      row.push(value.trim());
      value = '';
    } else if ((ch === '\n' || ch === '\r') && !inQuotes) {
      if (ch === '\r' && next === '\n') i += 1;
      row.push(value.trim());
      value = '';
      if (row.some((cell) => cell !== '')) rows.push(row);
      row = [];
    } else {
      value += ch;
    }
  }

  if (value.length || row.length) {
    row.push(value.trim());
    if (row.some((cell) => cell !== '')) rows.push(row);
  }

  return normalizeRows(rows);
}

function parseCsv(text) {
  return parseDelimited(text, ',');
}

function isMarkdownSeparatorRow(cells) {
  if (!cells.length) return false;
  return cells.every((cell) => /^:?-{3,}:?$/.test(cell.replace(/\s+/g, '')));
}

function parseMarkdownTable(text) {
  const lines = String(text).replace(/\r\n?/g, '\n').split('\n');
  const tableLines = lines.filter((line) => line.includes('|') && line.trim() !== '');

  const parsedRows = tableLines
    .map((line) => {
      let row = line.trim();
      if (row.startsWith('|')) row = row.slice(1);
      if (row.endsWith('|')) row = row.slice(0, -1);
      return row.split('|').map((cell) => cell.trim());
    })
    .filter((cells) => cells.some((cell) => cell !== ''));

  if (parsedRows.length >= 2 && isMarkdownSeparatorRow(parsedRows[1])) {
    parsedRows.splice(1, 1);
  }

  return normalizeRows(parsedRows);
}

function detectInputFormat(text) {
  const normalized = String(text || '').trim();
  if (!normalized) return INPUT_FORMATS.csv;

  const lines = normalized.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  const pipeLines = lines.filter((line) => line.includes('|'));

  if (pipeLines.length >= 2) {
    const second = pipeLines[1] || '';
    if (second.replace(/\s+/g, '').match(/^\|?:?-{3,}:?(\|:?-{3,}:?)+\|?$/)) return INPUT_FORMATS.markdown;
    if (pipeLines.every((line) => line.includes('|'))) return INPUT_FORMATS.markdown;
  }

  return INPUT_FORMATS.csv;
}

function getActiveInputFormat(text) {
  if (selectedInputFormat !== INPUT_FORMATS.auto) return selectedInputFormat;
  return detectInputFormat(text);
}

function parseByFormat(text) {
  const normalizedText = normalizeEscapedNewlines(text);
  const format = getActiveInputFormat(normalizedText);
  const rows = format === INPUT_FORMATS.markdown ? parseMarkdownTable(normalizedText) : parseCsv(normalizedText);
  return { rows, format };
}

function normalizeEscapedNewlines(text) {
  const source = String(text ?? '');
  const hasRealNewline = /\r|\n/.test(source);
  if (hasRealNewline || !source.includes('\\n')) return source;

  return source.replace(/\\r\\n/g, '\n').replace(/\\n/g, '\n').replace(/\\r/g, '\n');
}

function parseMarkdownStrong(text) {
  const source = String(text ?? '');
  const parts = source.split(/(\*\*[^*]+\*\*)/g).filter((part) => part.length > 0);

  if (!parts.length) return [{ text: source, strong: false }];

  return parts.map((part) => {
    const isStrong = /^\*\*[^*]+\*\*$/.test(part);
    return {
      text: isStrong ? part.slice(2, -2) : part,
      strong: isStrong,
    };
  });
}

function getCellDisplayValue(raw, format) {
  if (format !== INPUT_FORMATS.markdown) {
    return {
      text: raw ?? '',
      segments: [{ text: raw ?? '', strong: false }],
    };
  }

  const segments = parseMarkdownStrong(raw);
  return {
    text: segments.map((segment) => segment.text).join(''),
    segments,
  };
}

function measureStyledText(ctx, text, strong, fontSize) {
  ctx.font = `${strong ? '600' : '400'} ${fontSize}px ${els.fontFamily.value}`;
  return ctx.measureText(text).width;
}

function breakLongStyledToken(ctx, token, maxWidth, fontSize) {
  const chunks = [];
  let rest = token.text;

  while (rest.length) {
    let i = 1;
    while (i <= rest.length && measureStyledText(ctx, rest.slice(0, i), token.strong, fontSize) <= maxWidth) i += 1;
    const cut = Math.max(1, i - 1);
    chunks.push({ text: rest.slice(0, cut), strong: token.strong });
    rest = rest.slice(cut);
  }

  return chunks;
}

function wrapStyledText(ctx, segments, maxWidth, fontSize) {
  const lines = [];
  const src = segments.length ? segments : [{ text: '', strong: false }];

  src.forEach((segment, segIdx) => {
    const paragraphs = String(segment.text).replace(/\r\n?/g, '\n').split('\n');

    paragraphs.forEach((paragraph, pIdx) => {
      if (!lines.length) lines.push([]);

      const tokens = tokenizeText(paragraph).map((token) => ({ text: token, strong: !!segment.strong }));
      if (!tokens.length) {
        if (pIdx < paragraphs.length - 1) lines.push([]);
        return;
      }

      for (const token of tokens) {
        let current = lines[lines.length - 1];
        const currentWidth = current.reduce(
          (sum, part) => sum + measureStyledText(ctx, part.text, part.strong, fontSize),
          0
        );
        const tokenWidth = measureStyledText(ctx, token.text, token.strong, fontSize);

        if (!current.length || currentWidth + tokenWidth <= maxWidth) {
          current.push(token);
          continue;
        }

        if (tokenWidth <= maxWidth) {
          lines.push([token]);
          continue;
        }

        const chunks = breakLongStyledToken(ctx, token, maxWidth, fontSize);
        for (const chunk of chunks) {
          current = lines[lines.length - 1];
          const width = current.reduce(
            (sum, part) => sum + measureStyledText(ctx, part.text, part.strong, fontSize),
            0
          );
          const chunkWidth = measureStyledText(ctx, chunk.text, chunk.strong, fontSize);

          if (!current.length || width + chunkWidth <= maxWidth) current.push(chunk);
          else lines.push([chunk]);
        }
      }

      if (pIdx < paragraphs.length - 1) lines.push([]);
    });
  });

  return lines.length ? lines : [[]];
}

function setFormatMode(mode) {
  selectedInputFormat = mode;
  const activeFromText = getActiveInputFormat(els.csvText.value);

  els.formatAutoBtn.classList.toggle('active', mode === INPUT_FORMATS.auto);
  els.formatCsvBtn.classList.toggle('active', mode === INPUT_FORMATS.csv);
  els.formatMdBtn.classList.toggle('active', mode === INPUT_FORMATS.markdown);

  const labelMap = {
    [INPUT_FORMATS.auto]: `現在: 自動判定 (${activeFromText === INPUT_FORMATS.markdown ? 'Markdown表' : 'CSV'})`,
    [INPUT_FORMATS.csv]: '現在: CSV (手動指定)',
    [INPUT_FORMATS.markdown]: '現在: Markdown表 (手動指定)',
  };
  els.formatHint.textContent = labelMap[mode];
}

function readRatio() {
  const [w, h] = els.aspectRatio.value.split(':').map(Number);
  return w / h;
}

function calcCanvasSize() {
  const dpi = Number(els.dpi.value) || 1200;
  const ratio = readRatio();
  const baseWidthAt360 = 1920;
  const width = Math.max(640, Math.round((baseWidthAt360 * dpi) / 360));
  const height = Math.max(360, Math.round(width / ratio));
  return { width, height };
}

function getFontSize(tableW, tableH, rowCount, colCount) {
  const value = els.fontSize.value;
  if (value !== 'auto') return Number(value);

  const perCellW = tableW / Math.max(1, colCount);
  const perCellH = tableH / Math.max(1, rowCount);
  const fromWidth = Math.round(perCellW / 8.5);
  const fromHeight = Math.round(perCellH * 0.38);
  return Math.max(10, Math.min(156, Math.min(fromWidth, fromHeight)));
}

function updateCurrentFontSizeLabel(fontSize) {
  const modeLabel = els.fontSize.value === 'auto' ? '自動算出' : '固定';
  els.currentFontSize.textContent = `${fontSize}px (${modeLabel})`;
}

function tokenizeText(text) {
  if (!text) return [];

  if (typeof Intl !== 'undefined' && typeof Intl.Segmenter === 'function') {
    const segmenter = new Intl.Segmenter('ja', { granularity: 'word' });
    return Array.from(segmenter.segment(text), ({ segment }) => segment);
  }

  return Array.from(text);
}

function breakLongToken(ctx, token, maxWidth) {
  const broken = [];
  let rest = token;

  while (rest.length) {
    let i = 1;
    while (i <= rest.length && ctx.measureText(rest.slice(0, i)).width <= maxWidth) i += 1;
    const cut = Math.max(1, i - 1);
    broken.push(rest.slice(0, cut));
    rest = rest.slice(cut);
  }

  return broken;
}

function wrapText(ctx, text, maxWidth) {
  if (!text) return [''];

  const lines = [];
  const paragraphs = String(text).replace(/\r\n?/g, '\n').split('\n');

  for (const paragraph of paragraphs) {
    if (!paragraph) {
      lines.push('');
      continue;
    }

    if (ctx.measureText(paragraph).width <= maxWidth) {
      lines.push(paragraph);
      continue;
    }

    const tokens = tokenizeText(paragraph);
    let line = '';

    for (const token of tokens) {
      const candidate = `${line}${token}`;
      if (ctx.measureText(candidate).width <= maxWidth) {
        line = candidate;
        continue;
      }

      if (line) {
        lines.push(line);
        line = '';
      }

      if (ctx.measureText(token).width <= maxWidth) {
        line = token;
        continue;
      }

      lines.push(...breakLongToken(ctx, token, maxWidth));
    }

    if (line) lines.push(line);
  }

  return lines.length ? lines : [''];
}

function estimateColumnWeights(rows) {
  const colCount = rows[0].length;
  const weights = Array(colCount).fill(0);

  for (let c = 0; c < colCount; c += 1) {
    let total = 0;
    for (let r = 0; r < rows.length; r += 1) {
      const len = (rows[r][c] ?? '').replace(/\s+/g, ' ').trim().length;
      total += Math.max(1, len);
    }
    weights[c] = total;
  }

  return weights;
}

function estimateColumnMinWidths(ctx, rows, paddingX) {
  const colCount = rows[0].length;
  const minWidths = Array(colCount).fill(0);

  for (let c = 0; c < colCount; c += 1) {
    let longestTokenWidth = 0;

    for (let r = 0; r < rows.length; r += 1) {
      const text = (rows[r][c] ?? '').replace(/\s+/g, ' ').trim();
      if (!text) continue;

      const tokens = tokenizeText(text).filter((token) => token.trim());
      for (const token of tokens) {
        longestTokenWidth = Math.max(longestTokenWidth, ctx.measureText(token).width);
      }
    }

    minWidths[c] = Math.max(80, longestTokenWidth + paddingX * 2 + 2);
  }

  return minWidths;
}

function calculateColumnWidths(tableW, rows, ctx, paddingX) {
  const colCount = rows[0].length;
  const minColWidth = Math.max(80, tableW * 0.08);
  const minByToken = estimateColumnMinWidths(ctx, rows, paddingX);
  const weights = estimateColumnWeights(rows);
  const totalWeight = weights.reduce((a, b) => a + b, 0) || colCount;

  const raw = weights.map((w) => (tableW * w) / totalWeight);
  let widths = raw.map((w, idx) => Math.max(minColWidth, minByToken[idx], w));
  let sum = widths.reduce((a, b) => a + b, 0);

  if (sum > tableW) {
    const flexIdx = widths.map((w, idx) => [w, idx]).sort((a, b) => b[0] - a[0]);
    let overflow = sum - tableW;
    for (const [, idx] of flexIdx) {
      if (overflow <= 0) break;
      const strictMin = Math.max(minColWidth, minByToken[idx]);
      const canShrink = widths[idx] - strictMin;
      if (canShrink <= 0) continue;
      const delta = Math.min(canShrink, overflow);
      widths[idx] -= delta;
      overflow -= delta;
    }
  }

  sum = widths.reduce((a, b) => a + b, 0);
  const adjust = tableW - sum;
  widths[colCount - 1] += adjust;
  return widths;
}

function getAlignedX(x, cellW, paddingX, textWidth, align) {
  if (align === 'center') return x + cellW / 2 - textWidth / 2;
  if (align === 'right') return x + cellW - paddingX - textWidth;
  return x + paddingX;
}

function render() {
  const csvSource = els.csvText.value.trim();
  const { rows, format } = parseByFormat(csvSource);

  if (!rows.length) {
    alert('CSVデータが空です。');
    return;
  }

  const style = TABLE_STYLES[els.tableStyle.value];
  const theme = COLOR_THEMES[els.colorTheme.value];
  const hasHeaderRow = els.hasHeaderRow.checked;
  const hasHeaderCol = els.hasHeaderCol.checked;

  const { width, height } = calcCanvasSize();
  els.canvas.width = width;
  els.canvas.height = height;

  const ctx = els.canvas.getContext('2d');
  ctx.fillStyle = theme.bg;
  ctx.fillRect(0, 0, width, height);

  const margin = Math.round(Math.min(width, height) * 0.025);
  const tableX = margin;
  const tableY = margin;
  const tableW = width - margin * 2;
  const tableH = height - margin * 2;

  if (style.shadow) {
    ctx.save();
    ctx.fillStyle = 'rgba(15, 23, 42, 0.12)';
    roundRect(ctx, tableX + 8, tableY + 10, tableW, tableH, 14);
    ctx.fill();
    ctx.restore();
  }

  const rowCount = rows.length;
  const colCount = rows[0].length;
  let fontSize = getFontSize(tableW, tableH, rowCount, colCount);

  let lineHeight = 0;
  let paddingX = 0;
  let paddingY = 0;
  let colWidths = [];
  let wrappedCells = [];
  let rowHeights = [];

  do {
    lineHeight = Math.round(fontSize * 1.35);
    paddingX = Math.max(10, Math.round(fontSize * 0.65));
    paddingY = Math.max(8, Math.round(fontSize * 0.4));
    colWidths = calculateColumnWidths(tableW, rows, ctx, paddingX);

    wrappedCells = rows.map((row, r) =>
      row.map((raw, c) => {
        const { text, segments } = getCellDisplayValue(raw, format);
        const isHeader = (hasHeaderRow && r === 0) || (hasHeaderCol && c === 0);
        const weight = isHeader ? '600' : '400';
        ctx.font = `${weight} ${fontSize}px ${els.fontFamily.value}`;
        return {
          lines:
            format === INPUT_FORMATS.markdown
              ? wrapStyledText(ctx, segments, Math.max(24, colWidths[c] - paddingX * 2), fontSize)
              : wrapText(ctx, text, Math.max(24, colWidths[c] - paddingX * 2)).map((line) => [
                  { text: line, strong: false },
                ]),
        };
      })
    );

    rowHeights = wrappedCells.map((cells) => {
      const maxLines = Math.max(...cells.map((cell) => cell.lines.length));
      return maxLines * lineHeight + paddingY * 2;
    });

    const contentHeight = rowHeights.reduce((a, b) => a + b, 0);
    if (contentHeight <= tableH || fontSize <= 8) break;
    fontSize -= 1;
  } while (true);

  ctx.textBaseline = 'alphabetic';

  let y = tableY;
  for (let r = 0; r < rowCount; r += 1) {
    let x = tableX;
    const rowH = rowHeights[r];

    for (let c = 0; c < colCount; c += 1) {
      const isHeader = (hasHeaderRow && r === 0) || (hasHeaderCol && c === 0);
      let cellBg = theme.bg;
      if (isHeader) cellBg = theme.headerBg;
      else if (hasHeaderCol && c === 0) cellBg = theme.firstColBg;
      else if (style.stripe && r % 2 === 1) cellBg = theme.stripeBg;

      const cellW = colWidths[c];

      if (style.roundedCells) {
        roundRect(ctx, x + 1, y + 1, cellW - 2, rowH - 2, 8);
        ctx.fillStyle = cellBg;
        ctx.fill();
      } else {
        ctx.fillStyle = cellBg;
        ctx.fillRect(x, y, cellW, rowH);
      }

      if (style.borderWidth > 0) {
        ctx.lineWidth = style.borderWidth;
        ctx.strokeStyle = theme.border;
        ctx.strokeRect(x, y, cellW, rowH);
      }

      const cellText = wrappedCells[r][c];
      const lines = cellText.lines;
      const align = isHeader ? els.headerAlign.value : els.bodyAlign.value;
      ctx.fillStyle = isHeader ? theme.headerText : theme.text;

      const lineCount = lines.length;
      const textBlockH = lineCount * lineHeight;
      let textY = y + (rowH - textBlockH) / 2 + fontSize;

      lines.forEach((lineParts) => {
        const textWidth = lineParts.reduce(
          (sum, part) => sum + measureStyledText(ctx, part.text, isHeader || part.strong, fontSize),
          0
        );
        const drawX = getAlignedX(x, cellW, paddingX, textWidth, align);
        let cursorX = drawX;

        lineParts.forEach((part) => {
          const strong = isHeader || part.strong;
          ctx.font = `${strong ? '600' : '400'} ${fontSize}px ${els.fontFamily.value}`;
          ctx.fillText(part.text, cursorX, textY);
          cursorX += ctx.measureText(part.text).width;
        });

        textY += lineHeight;
      });

      x += cellW;
    }

    y += rowH;
  }

  updateCurrentFontSizeLabel(fontSize);

  if (style.outerFrame) {
    ctx.lineWidth = 2;
    ctx.strokeStyle = theme.accent;
    roundRect(ctx, tableX, tableY, tableW, tableH, 12);
    ctx.stroke();
  }
}

function roundRect(ctx, x, y, width, height, radius) {
  const r = Math.min(radius, width / 2, height / 2);
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + width, y, x + width, y + height, r);
  ctx.arcTo(x + width, y + height, x, y + height, r);
  ctx.arcTo(x, y + height, x, y, r);
  ctx.arcTo(x, y, x + width, y, r);
  ctx.closePath();
}

function downloadPng() {
  render();
  const link = document.createElement('a');
  link.href = els.canvas.toDataURL('image/png');
  link.download = `table_${Date.now()}.png`;
  link.click();
}

els.csvFile.addEventListener('change', async (event) => {
  const [file] = event.target.files;
  if (!file) return;
  const text = await file.text();
  els.csvText.value = text;
  setFormatMode(selectedInputFormat);
  render();
});

els.formatAutoBtn.addEventListener('click', () => {
  setFormatMode(INPUT_FORMATS.auto);
  render();
});
els.formatCsvBtn.addEventListener('click', () => {
  setFormatMode(INPUT_FORMATS.csv);
  render();
});
els.formatMdBtn.addEventListener('click', () => {
  setFormatMode(INPUT_FORMATS.markdown);
  render();
});
els.csvText.addEventListener('input', () => {
  if (selectedInputFormat === INPUT_FORMATS.auto) setFormatMode(INPUT_FORMATS.auto);
});

els.renderBtn.addEventListener('click', render);
els.downloadBtn.addEventListener('click', downloadPng);

initSelect(els.tableStyle, TABLE_STYLES);
initSelect(els.colorTheme, COLOR_THEMES);
els.csvText.value =
  '項目,値,備考\n売上,1200,前月比 +10%\n利益,320,改善傾向\n顧客満足度,4.4,アンケート結果\n長い注記,この列は分量に応じて幅が広がり、必要に応じて折り返して表示されます,見切れ防止';
setFormatMode(INPUT_FORMATS.auto);
render();
