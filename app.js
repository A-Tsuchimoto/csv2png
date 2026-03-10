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
  hasHeaderRow: document.getElementById('hasHeaderRow'),
  hasHeaderCol: document.getElementById('hasHeaderCol'),
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
  const dpi = Number(els.dpi.value) || 180;
  const ratio = readRatio();
  const baseWidthAt180 = 1920;
  const width = Math.max(640, Math.round((baseWidthAt180 * dpi) / 180));
  const height = Math.max(360, Math.round(width / ratio));
  return { width, height };
}

function fitText(ctx, text, maxWidth) {
  if (ctx.measureText(text).width <= maxWidth) return text;
  let out = text;
  while (out.length > 1 && ctx.measureText(`${out}…`).width > maxWidth) {
    out = out.slice(0, -1);
  }
  return `${out}…`;
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
  const colWidths = Array(colCount).fill(tableW / colCount);
  const rowH = tableH / rowCount;

  const fontSize = Math.max(12, Math.round(width / 70));
  const paddingX = Math.round(fontSize * 0.7);

  ctx.textBaseline = 'middle';
  ctx.font = `${fontSize}px ${els.fontFamily.value}`;

  let y = tableY;
  for (let r = 0; r < rowCount; r += 1) {
    let x = tableX;
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

      const raw = rows[r][c] ?? '';
      const text = fitText(ctx, raw, cellW - paddingX * 2);
      ctx.fillStyle = isHeader ? theme.headerText : theme.text;
      ctx.font = `${isHeader ? '600' : '400'} ${fontSize}px ${els.fontFamily.value}`;
      ctx.fillText(text, x + paddingX, y + rowH / 2);

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
els.csvText.value = '項目,値,備考\n売上,1200,前月比 +10%\n利益,320,改善傾向\n顧客満足度,4.4,アンケート結果';
render();
