const TABLE_STYLES = {
  simpleClean: {
    label: 'シンプル1: クリーン',
    borderWidth: 1,
    stripe: false,
    roundedCells: false,
    shadow: false,
    outerFrame: true,
  },
  simpleStriped: {
    label: 'シンプル2: ストライプ',
    borderWidth: 1,
    stripe: true,
    roundedCells: false,
    shadow: false,
    outerFrame: true,
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
    outerFrame: true,
  },
};

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
  csvText: document.getElementById('csvText'),
  aspectRatio: document.getElementById('aspectRatio'),
  dpi: document.getElementById('dpi'),
  fontFamily: document.getElementById('fontFamily'),
  fontSize: document.getElementById('fontSize'),
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

function parseCsv(text) {
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
    } else if (ch === ',' && !inQuotes) {
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

  const maxCols = rows.reduce((max, r) => Math.max(max, r.length), 0);
  return rows.map((r) => [...r, ...Array(maxCols - r.length).fill('')]);
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
  return Math.max(10, Math.min(60, Math.min(fromWidth, fromHeight)));
}

function wrapText(ctx, text, maxWidth) {
  if (!text) return [''];

  const normalized = text.replace(/\s+/g, ' ').trim();
  if (!normalized) return [''];

  if (ctx.measureText(normalized).width <= maxWidth) return [normalized];

  const chunks = normalized.split(' ');
  const lines = [];
  let line = '';

  const pushBrokenWord = (word) => {
    let rest = word;
    while (rest.length) {
      let i = 1;
      while (i <= rest.length && ctx.measureText(rest.slice(0, i)).width <= maxWidth) i += 1;
      const cut = Math.max(1, i - 1);
      lines.push(rest.slice(0, cut));
      rest = rest.slice(cut);
    }
  };

  chunks.forEach((word) => {
    const candidate = line ? `${line} ${word}` : word;
    if (ctx.measureText(candidate).width <= maxWidth) {
      line = candidate;
      return;
    }

    if (line) {
      lines.push(line);
      line = '';
    }

    if (ctx.measureText(word).width <= maxWidth) {
      line = word;
      return;
    }

    pushBrokenWord(word);
  });

  if (line) lines.push(line);
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

      const tokens = text.split(' ');
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
  const rows = parseCsv(csvSource);

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

  const margin = Math.round(width * 0.05);
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
        const isHeader = (hasHeaderRow && r === 0) || (hasHeaderCol && c === 0);
        ctx.font = `${isHeader ? '600' : '400'} ${fontSize}px ${els.fontFamily.value}`;
        return wrapText(ctx, raw ?? '', Math.max(24, colWidths[c] - paddingX * 2));
      })
    );

    rowHeights = wrappedCells.map((cells) => {
      const maxLines = Math.max(...cells.map((lines) => lines.length));
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

      const lines = wrappedCells[r][c];
      const align = isHeader ? els.headerAlign.value : els.bodyAlign.value;
      ctx.fillStyle = isHeader ? theme.headerText : theme.text;
      ctx.font = `${isHeader ? '600' : '400'} ${fontSize}px ${els.fontFamily.value}`;

      const lineCount = lines.length;
      const textBlockH = lineCount * lineHeight;
      let textY = y + (rowH - textBlockH) / 2 + fontSize;

      lines.forEach((line) => {
        const textWidth = ctx.measureText(line).width;
        const drawX = getAlignedX(x, cellW, paddingX, textWidth, align);
        ctx.fillText(line, drawX, textY);
        textY += lineHeight;
      });

      x += cellW;
    }

    y += rowH;
  }

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
  render();
});

els.renderBtn.addEventListener('click', render);
els.downloadBtn.addEventListener('click', downloadPng);

initSelect(els.tableStyle, TABLE_STYLES);
initSelect(els.colorTheme, COLOR_THEMES);
els.csvText.value =
  '項目,値,備考\n売上,1200,前月比 +10%\n利益,320,改善傾向\n顧客満足度,4.4,アンケート結果\n長い注記,この列は分量に応じて幅が広がり、必要に応じて折り返して表示されます,見切れ防止';
render();
