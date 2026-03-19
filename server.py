#!/usr/bin/env python3
"""品牌方AI营销采购访谈工具 - Python后端 (Supabase 版)"""

import json
import os
import uuid
import csv
import io
import zipfile
import xml.etree.ElementTree as ET
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
from urllib.request import Request, urlopen
from urllib.error import HTTPError
from datetime import datetime

PORT = int(os.environ.get('PORT', '3000'))
PUBLIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'public')

# Supabase 配置
SUPABASE_URL = os.environ.get('SUPABASE_URL', '')
SUPABASE_KEY = os.environ.get('SUPABASE_KEY', '')
SUPABASE_TABLE = 'interviews'

QUESTION_LABELS = [
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
]

# ============ Supabase REST API 封装 ============

def _supabase_headers(prefer=None):
    headers = {
        'apikey': SUPABASE_KEY,
        'Authorization': f'Bearer {SUPABASE_KEY}',
        'Content-Type': 'application/json',
    }
    if prefer:
        headers['Prefer'] = prefer
    return headers

def _supabase_request(method, path, body=None, headers_extra=None):
    url = f'{SUPABASE_URL}/rest/v1/{path}'
    data = json.dumps(body, ensure_ascii=False).encode('utf-8') if body else None
    headers = _supabase_headers()
    if headers_extra:
        headers.update(headers_extra)
    req = Request(url, data=data, headers=headers, method=method)
    try:
        with urlopen(req) as resp:
            raw = resp.read().decode('utf-8')
            return json.loads(raw) if raw else None
    except HTTPError as e:
        error_body = e.read().decode('utf-8')
        print(f'Supabase error: {e.code} {error_body}', flush=True)
        raise

def db_read_all():
    """读取所有访谈记录"""
    rows = _supabase_request('GET', f'{SUPABASE_TABLE}?select=*&order=created_at.desc')
    return [row['data'] for row in rows] if rows else []

def db_read_one(interview_id):
    """读取单条记录"""
    rows = _supabase_request('GET', f'{SUPABASE_TABLE}?interview_id=eq.{interview_id}&select=*')
    if rows and len(rows) > 0:
        return rows[0]['data']
    return None

def db_insert(interview):
    """插入一条记录"""
    row = {
        'interview_id': interview['id'],
        'data': interview,
    }
    _supabase_request('POST', SUPABASE_TABLE, body=row, headers_extra={'Prefer': 'return=minimal'})

def db_update(interview_id, interview):
    """更新一条记录"""
    row = {
        'data': interview,
    }
    _supabase_request('PATCH', f'{SUPABASE_TABLE}?interview_id=eq.{interview_id}', body=row, headers_extra={'Prefer': 'return=minimal'})

def db_delete(interview_id):
    """删除一条记录"""
    _supabase_request('DELETE', f'{SUPABASE_TABLE}?interview_id=eq.{interview_id}')

# ============ 导出相关 ============

def build_rows(interviews):
    rows = []
    for iv in interviews:
        row = {
            '受访者': iv.get('name', ''),
            '职位': iv.get('title', ''),
            '公司': iv.get('company', ''),
            '行业': iv.get('industry', ''),
            '访谈日期': iv.get('date', ''),
            '访谈时长': iv.get('duration', ''),
            'AI成熟度': iv.get('maturity', ''),
        }
        answers = iv.get('answers', [])
        for i, label in enumerate(QUESTION_LABELS):
            row[label] = answers[i] if i < len(answers) else ''
        findings = iv.get('findings', [])
        for i in range(3):
            row[f'核心发现{i+1}'] = findings[i] if i < len(findings) else ''
        row['P12校验信号'] = iv.get('p12Signal', '')
        row['对ICC的启示'] = iv.get('iccInsight', '')
        row['意外洞察'] = iv.get('surprise', '')
        row['后续跟进'] = iv.get('followUp', '')
        rows.append(row)
    return rows

def generate_xlsx(rows):
    """Generate a minimal .xlsx file using zipfile + XML (no dependencies)."""
    if not rows:
        rows = [{'': ''}]

    columns = list(rows[0].keys())

    def col_letter(idx):
        result = ''
        while True:
            result = chr(65 + idx % 26) + result
            idx = idx // 26 - 1
            if idx < 0:
                break
        return result

    def escape_xml(s):
        return str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')

    shared_strings = []
    ss_index = {}
    def get_ss_idx(val):
        val = str(val)
        if val not in ss_index:
            ss_index[val] = len(shared_strings)
            shared_strings.append(val)
        return ss_index[val]

    for col in columns:
        get_ss_idx(col)
    for row in rows:
        for col in columns:
            get_ss_idx(str(row.get(col, '')))

    sheet_rows = []
    cells = []
    for ci, col in enumerate(columns):
        ref = f'{col_letter(ci)}1'
        cells.append(f'<c r="{ref}" t="s" s="1"><v>{get_ss_idx(col)}</v></c>')
    sheet_rows.append(f'<row r="1">{"".join(cells)}</row>')

    for ri, row in enumerate(rows, start=2):
        cells = []
        for ci, col in enumerate(columns):
            ref = f'{col_letter(ci)}{ri}'
            val = str(row.get(col, ''))
            cells.append(f'<c r="{ref}" t="s" s="0"><v>{get_ss_idx(val)}</v></c>')
        sheet_rows.append(f'<row r="{ri}">{"".join(cells)}</row>')

    dim_end = f'{col_letter(len(columns)-1)}{len(rows)+1}'

    sheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="A1:{dim_end}"/>
<cols>{''.join(f'<col min="{i+1}" max="{i+1}" width="{40 if col.startswith("Q") else 18}" customWidth="1"/>' for i, col in enumerate(columns))}</cols>
<sheetData>{"".join(sheet_rows)}</sheetData>
</worksheet>'''

    ss_items = ''.join(f'<si><t>{escape_xml(s)}</t></si>' for s in shared_strings)
    ss_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(shared_strings)}" uniqueCount="{len(shared_strings)}">
{ss_items}
</sst>'''

    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2">
<font><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
</fonts>
<fills count="3">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill>
</fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="1"><xf/></cellStyleXfs>
<cellXfs count="2">
<xf fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
<xf fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
</cellXfs>
</styleSheet>'''

    workbook_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="访谈记录" sheetId="1" r:id="rId1"/></sheets>
</workbook>'''

    rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

    wb_rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

    content_types_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>'''

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', content_types_xml)
        zf.writestr('_rels/.rels', rels_xml)
        zf.writestr('xl/workbook.xml', workbook_xml)
        zf.writestr('xl/_rels/workbook.xml.rels', wb_rels_xml)
        zf.writestr('xl/worksheets/sheet1.xml', sheet_xml)
        zf.writestr('xl/sharedStrings.xml', ss_xml)
        zf.writestr('xl/styles.xml', styles_xml)
    return buf.getvalue()

def generate_csv(rows):
    if not rows:
        return ''
    columns = list(rows[0].keys())
    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=columns)
    writer.writeheader()
    for row in rows:
        writer.writerow(row)
    return '\ufeff' + output.getvalue()


class InterviewHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=PUBLIC_DIR, **kwargs)

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path

        if path == '/api/interviews':
            data = db_read_all()
            self._json_response(data)
        elif path.startswith('/api/interviews/') and '/export' not in path:
            interview_id = path.split('/')[-1]
            item = db_read_one(interview_id)
            if item:
                self._json_response(item)
            else:
                self._json_response({'error': 'Not found'}, 404)
        elif path == '/api/export/xlsx':
            params = parse_qs(parsed.query)
            ids = params.get('ids', [''])[0].split(',') if 'ids' in params else None
            data = db_read_all()
            filtered = [d for d in data if ids is None or d['id'] in ids]
            rows = build_rows(filtered)
            xlsx_data = generate_xlsx(rows)
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', f'attachment; filename=interviews_{int(datetime.now().timestamp())}.xlsx')
            self.send_header('Content-Length', str(len(xlsx_data)))
            self.end_headers()
            self.wfile.write(xlsx_data)
        elif path == '/api/export/csv':
            params = parse_qs(parsed.query)
            ids = params.get('ids', [''])[0].split(',') if 'ids' in params else None
            data = db_read_all()
            filtered = [d for d in data if ids is None or d['id'] in ids]
            rows = build_rows(filtered)
            csv_data = generate_csv(rows).encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type', 'text/csv; charset=utf-8')
            self.send_header('Content-Disposition', f'attachment; filename=interviews_{int(datetime.now().timestamp())}.csv')
            self.send_header('Content-Length', str(len(csv_data)))
            self.end_headers()
            self.wfile.write(csv_data)
        else:
            super().do_GET()

    def do_POST(self):
        if self.path == '/api/interviews':
            body = self._read_body()
            interview = {
                'id': str(uuid.uuid4()),
                'createdAt': datetime.now().isoformat(),
                'updatedAt': datetime.now().isoformat(),
                **body
            }
            db_insert(interview)
            self._json_response(interview)
        else:
            self._json_response({'error': 'Not found'}, 404)

    def do_PUT(self):
        if self.path.startswith('/api/interviews/'):
            interview_id = self.path.split('/')[-1]
            body = self._read_body()
            existing = db_read_one(interview_id)
            if existing:
                updated = {**existing, **body, 'updatedAt': datetime.now().isoformat()}
                db_update(interview_id, updated)
                self._json_response(updated)
            else:
                self._json_response({'error': 'Not found'}, 404)
        else:
            self._json_response({'error': 'Not found'}, 404)

    def do_DELETE(self):
        if self.path.startswith('/api/interviews/'):
            interview_id = self.path.split('/')[-1]
            db_delete(interview_id)
            self._json_response({'success': True})
        else:
            self._json_response({'error': 'Not found'}, 404)

    def _read_body(self):
        length = int(self.headers.get('Content-Length', 0))
        raw = self.rfile.read(length)
        return json.loads(raw.decode('utf-8'))

    def _json_response(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format, *args):
        import sys
        sys.stderr.write("%s - - [%s] %s\n" % (self.client_address[0], self.log_date_time_string(), format % args))
        sys.stderr.flush()


if __name__ == '__main__':
    if not SUPABASE_URL or not SUPABASE_KEY:
        print('警告: 未配置 SUPABASE_URL 或 SUPABASE_KEY 环境变量', flush=True)
        print('请设置环境变量后重新启动', flush=True)
    print(f'Starting server on port {PORT}...', flush=True)
    print(f'Supabase: {SUPABASE_URL}', flush=True)
    server = HTTPServer(('0.0.0.0', PORT), InterviewHandler)
    print(f'Listening on 0.0.0.0:{PORT}', flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\n服务已停止')
        server.server_close()
