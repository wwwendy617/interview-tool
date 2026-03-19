const express = require('express');
const fs = require('fs');
const path = require('path');
const { v4: uuidv4 } = require('uuid');
const ExcelJS = require('exceljs');

const app = express();
const PORT = 3000;
const DATA_FILE = path.join(__dirname, 'data', 'interviews.json');

app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// --- Data helpers ---
function readData() {
  const raw = fs.readFileSync(DATA_FILE, 'utf-8');
  return JSON.parse(raw);
}

function writeData(data) {
  fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), 'utf-8');
}

// --- API ---

// Get all interviews
app.get('/api/interviews', (req, res) => {
  res.json(readData());
});

// Get single interview
app.get('/api/interviews/:id', (req, res) => {
  const data = readData();
  const item = data.find(d => d.id === req.params.id);
  if (!item) return res.status(404).json({ error: 'Not found' });
  res.json(item);
});

// Create interview
app.post('/api/interviews', (req, res) => {
  const data = readData();
  const interview = {
    id: uuidv4(),
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    ...req.body
  };
  data.push(interview);
  writeData(data);
  res.json(interview);
});

// Update interview
app.put('/api/interviews/:id', (req, res) => {
  const data = readData();
  const idx = data.findIndex(d => d.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'Not found' });
  data[idx] = { ...data[idx], ...req.body, updatedAt: new Date().toISOString() };
  writeData(data);
  res.json(data[idx]);
});

// Delete interview
app.delete('/api/interviews/:id', (req, res) => {
  let data = readData();
  data = data.filter(d => d.id !== req.params.id);
  writeData(data);
  res.json({ success: true });
});

// --- Export ---

const QUESTION_LABELS = [
  'Q1 AI营销使用现状',
  'Q2 AI带来的变化与价值',
  'Q3 AI内容效果与风险',
  'Q4 采购决策流程',
  'Q5 AI营销预算来源',
  'Q6 服务商评估标准',
  'Q7 自建vs外采',
  'Q8 品牌重要性认知',
  'Q9 消费者洞察工具价值',
  'Q10 品牌策略付费意愿',
  'Q11 未来展望',
  'Q12 给服务商的建议'
];

function buildRows(interviews) {
  return interviews.map(iv => ({
    '受访者': iv.name || '',
    '职位': iv.title || '',
    '公司': iv.company || '',
    '行业': iv.industry || '',
    '访谈日期': iv.date || '',
    '访谈时长': iv.duration || '',
    'AI成熟度': iv.maturity || '',
    ...Object.fromEntries(QUESTION_LABELS.map((label, i) => [label, iv.answers?.[i] || ''])),
    '核心发现1': iv.findings?.[0] || '',
    '核心发现2': iv.findings?.[1] || '',
    '核心发现3': iv.findings?.[2] || '',
    'P12校验信号': iv.p12Signal || '',
    '对ICC的启示': iv.iccInsight || '',
    '意外洞察': iv.surprise || '',
    '后续跟进': iv.followUp || ''
  }));
}

// Export Excel
app.get('/api/export/xlsx', async (req, res) => {
  const data = readData();
  const ids = req.query.ids ? req.query.ids.split(',') : data.map(d => d.id);
  const filtered = data.filter(d => ids.includes(d.id));
  const rows = buildRows(filtered);

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('访谈记录');

  if (rows.length > 0) {
    const columns = Object.keys(rows[0]);
    sheet.columns = columns.map(key => ({
      header: key,
      key,
      width: key.startsWith('Q') ? 40 : 18
    }));

    // Style header row
    sheet.getRow(1).font = { bold: true };
    sheet.getRow(1).fill = {
      type: 'pattern', pattern: 'solid',
      fgColor: { argb: 'FF4472C4' }
    };
    sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };

    rows.forEach(row => sheet.addRow(row));

    // Wrap text for Q columns
    sheet.eachRow((row, rowNum) => {
      if (rowNum > 1) {
        row.alignment = { wrapText: true, vertical: 'top' };
      }
    });
  }

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename=interviews_${Date.now()}.xlsx`);
  await workbook.xlsx.write(res);
  res.end();
});

// Export CSV
app.get('/api/export/csv', (req, res) => {
  const data = readData();
  const ids = req.query.ids ? req.query.ids.split(',') : data.map(d => d.id);
  const filtered = data.filter(d => ids.includes(d.id));
  const rows = buildRows(filtered);

  if (rows.length === 0) {
    return res.status(200).send('');
  }

  const columns = Object.keys(rows[0]);
  const escapeCsv = (val) => {
    const s = String(val).replace(/"/g, '""');
    return s.includes(',') || s.includes('"') || s.includes('\n') ? `"${s}"` : s;
  };

  const csvLines = [
    '\uFEFF' + columns.map(escapeCsv).join(','), // BOM for Excel Chinese support
    ...rows.map(row => columns.map(col => escapeCsv(row[col])).join(','))
  ];

  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.setHeader('Content-Disposition', `attachment; filename=interviews_${Date.now()}.csv`);
  res.send(csvLines.join('\n'));
});

app.listen(PORT, () => {
  console.log(`访谈工具已启动: http://localhost:${PORT}`);
});
