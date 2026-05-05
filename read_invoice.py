import zipfile, xml.etree.ElementTree as ET, sys, json

path = r'C:\Users\wdxyh\.openclaw\media\inbound\å_é_å_ç_æ_è_å_¼å_ºç_æ---cb1ab75c-865e-402e-a469-c2fb38396531.xlsx'
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

def get_cell_value(c):
    is_el = c.find('{%s}is/{%s}t' % (NS, NS))
    v_el = c.find('{%s}v' % NS)
    if is_el not in (None, '') and hasattr(is_el, 'text') and is_el.text:
        return is_el.text
    if v_el is not None:
        return v_el.text or ''
    return ''

def col_to_num(col):
    num = 0
    for ch in col:
        num = num * 26 + (ord(ch) - ord('A') + 1)
    return num

def cell_ref_to_col(ref):
    col = ''
    num = ''
    for ch in ref:
        if ch.isalpha():
            col += ch
        else:
            num += ch
    return col, int(num)

invoices_in = []  # 进项发票
invoices_out = []  # 销项发票

with zipfile.ZipFile(path) as z:
    for si, sheet_path in enumerate(['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml']):
        ws_xml = z.read(sheet_path)
        ws = ET.fromstring(ws_xml)
        rows_data = {}
        for row in ws.findall('.//{%s}row' % NS):
            rn = int(row.get('r'))
            rows_data[rn] = {}
            for c in row.findall('{%s}c' % NS):
                ref = c.get('r')
                rows_data[rn][ref] = get_cell_value(c)
        
        if si == 0:
            headers = {k: v for k, v in rows_data.get(1, {}).items() if v}
        else:
            headers2 = {k: v for k, v in rows_data.get(1, {}).items() if v}
        
        for rn in sorted(rows_data.keys()):
            if rn == 1:
                continue
            row = rows_data[rn]
            if si == 0:
                invoices_in.append(row)
            else:
                invoices_out.append(row)

print('IN count:', len(invoices_in))
print('OUT count:', len(invoices_out))
print()
print('=== 进项发票 Sheet1 headers ===')
print(headers)
print()
print('=== 销项发票 Sheet2 headers ===')
print(headers2)
print()

# Print first few rows of each
print('=== 进项前3行 ===')
for i, row in enumerate(invoices_in[:3]):
    print(i+1, ':', {k: v for k, v in row.items() if v})

print()
print('=== 销项前3行 ===')
for i, row in enumerate(invoices_out[:3]):
    print(i+1, ':', {k: v for k, v in row.items() if v})

# Try to find key financial data
print()
print('=== 进项发票金额分析 ===')
total_in = 0
for row in invoices_in:
    q = row.get('Q', '')  # 金额？
    t = row.get('T', '')  # 价税合计？
    if q:
        try: total_in += float(q)
        except: pass
    if t:
        try: pass  # total already has it
        except: pass

print('Total in Q column:', total_in)

# Output full data for analysis
out_data = {
    'in_headers': headers,
    'out_headers': headers2,
    'in_count': len(invoices_in),
    'out_count': len(invoices_out),
    'in_first3': [{k: v for k, v in row.items() if v} for row in invoices_in[:5]],
    'out_first3': [{k: v for k, v in row.items() if v} for row in invoices_out[:5]],
}
print('JSON:', json.dumps(out_data, ensure_ascii=False, indent=2))
