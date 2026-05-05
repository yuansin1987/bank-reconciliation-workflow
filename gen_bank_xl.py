import zipfile, io
from collections import defaultdict

def xl_esc(s):
    return str(s).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')

def col_letter(n):
    s = ''
    while n > 0:
        n, r = divmod(n-1, 26)
        s = chr(65+r) + s
    return s

def make_xlsx(out_path, sheets):
    buf = io.BytesIO()
    all_strings = []
    def si(s):
        s = str(s)
        if s not in all_strings:
            all_strings.append(s)
        return all_strings.index(s)
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>']
        for i in range(len(sheets)):
            ct.append('<Override PartName="/xl/worksheets/sheet' + str(i+1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
        ct += ['<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.workbook.main+xml"/>','<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>','<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>','</Types>']
        zf.writestr('[Content_Types].xml', ''.join(ct))
        zf.writestr('_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        wb = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>']
        for i, sd in enumerate(sheets):
            wb.append('<sheet name="' + xl_esc(sd[0]) + '" sheetId="' + str(i+1) + '" r:id="rId' + str(i+1) + '"/>')
        wb += ['</sheets>', '</workbook>']
        zf.writestr('xl/workbook.xml', ''.join(wb))
        wr = ['<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        for i in range(len(sheets)):
            wr.append('<Relationship Id="rId' + str(i+1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + str(i+1) + '.xml"/>')
        wr += ['<Relationship Id="rId999" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>','<Relationship Id="rId998" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>','</Relationships>']
        zf.writestr('xl/_rels/workbook.xml.rels', ''.join(wr))
        styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts><font><sz val="10"/><name val="微软雅黑"/></font><font><sz val="10"/><b/><name val="微软雅黑"/></font></fonts><fills><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFC000"/></patternFill></fill></fills><borders><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="center"/></xf><xf numFmtId="2" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"><alignment horizontal="right"/></xf><xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="center"/></xf></cellXfs></styleSheet>'
        zf.writestr('xl/styles.xml', styles)
        for si_idx, sd in enumerate(sheets, 1):
            sname, headers, rows = sd[0], sd[1], sd[2]
            ws_rows = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>']
            ws_rows.append('<row r="1">')
            for ci, h in enumerate(headers, 1):
                cl = col_letter(ci)
                ws_rows.append('<c r="' + cl + '1" t="s" s="1"><v>' + str(si(h)) + '</v></c>')
            ws_rows.append('</row>')
            for ri, row in enumerate(rows, 2):
                ws_rows.append('<row r="' + str(ri) + '">')
                for ci, val in enumerate(row, 1):
                    cl = col_letter(ci)
                    if isinstance(val, float) and val != 0:
                        ws_rows.append('<c r="' + cl + str(ri) + '" s="2"><v>' + ('%.2f' % val) + '</v></c>')
                    elif isinstance(val, int) and val != 0:
                        ws_rows.append('<c r="' + cl + str(ri) + '" s="2"><v>' + str(val) + '</v></c>')
                    elif val in (0, '0'):
                        ws_rows.append('<c r="' + cl + str(ri) + '" s="2"><v>0</v></c>')
                    elif val not in (None, ''):
                        ws_rows.append('<c r="' + cl + str(ri) + '" t="s" s="0"><v>' + str(si(val)) + '</v></c>')
                    else:
                        ws_rows.append('<c r="' + cl + str(ri) + '" s="0"/>')
                ws_rows.append('</row>')
            ws_rows += ['</sheetData>', '</worksheet>']
            zf.writestr('xl/worksheets/sheet' + str(si_idx) + '.xml', ''.join(ws_rows))
        ss_count = str(len(all_strings))
        ss_xml = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + ss_count + '" uniqueCount="' + ss_count + '">']
        for s in all_strings:
            ss_xml.append('<si><t xml:space="preserve">' + xl_esc(s) + '</t></si>')
        ss_xml.append('</sst>')
        zf.writestr('xl/sharedStrings.xml', ''.join(ss_xml))
    with open(out_path, 'wb') as f:
        f.write(buf.getvalue())
    print('Written:', out_path, len(buf.getvalue()), 'bytes')

bank_data = [
    ('20260302','对公提回贷','三体系+五星售后年审费用',13000.00,'武汉赛思教学设备有限公司'),
    ('20260302','对公转账出','退款',-4750.00,'武汉甲牛技术有限公司'),
    ('20260303','对公提回贷','预付款',2400.00,'湖北楚食安供应链科技有限公司'),
    ('20260303','对公转账出','服务费（瑞钰祥）',-2670.00,'创璟众创空间（武汉）有限公司'),
    ('20260304','汇入汇款','认证费用',2250.00,'武汉星驰激光科技有限公司'),
    ('20260304','汇入汇款','证书年审费',8000.00,'湖北熠橙科技有限公司'),
    ('20260304','汇入汇款','认证年审费',10000.00,'武汉环投胜韵通工程项目管理有限公司'),
    ('20260305','汇入汇款','体系认证费',6000.00,'湖北本杰建设工程有限公司'),
    ('20260305','对公提回贷','',11000.00,'武汉能科工程技术研究有限公司'),
    ('20260305','对公提回贷','无附言',10000.00,'湖北世强建设工程有限公司'),
    ('20260305','汇入汇款','报名费',14000.00,'湖北朗固建设工程有限公司'),
    ('20260306','汇入汇款','ISO9001年审费',4000.00,'武汉思创智能装备有限公司'),
    ('20260306','汇入汇款','三体系审核费',10000.00,'湖北屯仓管业科技发展有限公司'),
    ('20260306','汇入汇款','咨询费',9000.00,'湖北长港水泥制品有限公司'),
    ('20260309','汇入汇款','体系尾款',6000.00,'南京国网南自工程有限公司'),
    ('20260310','对公提回贷','转账汇款',12000.00,'武汉东城垸新型墙体有限公司'),
    ('20260310','汇入汇款','ISO认证费',6000.00,'湖北毓秀建筑工程有限公司'),
    ('20260310','汇入汇款','iso体系认证',5000.00,'长沙城北贸易有限责任公司'),
    ('20260310','汇入汇款','T19001体系再认证费用',5000.00,'荆州市钜荣智能科技有限公司'),
    ('20260310','对公提回贷','认证费',13000.00,'宜昌楚邦建设工程有限公司'),
    ('20260311','对公转账出','服务费（德希）',-895.00,'武汉璞然知识产权代理有限公司'),
    ('20260311','对公转账出','服务费（能科）',-2670.00,'武汉盛凡空间信息技术有限公司'),
    ('20260311','汇入汇款','',12750.00,'苏州照明科技有限公司'),
    ('20260311','对公提回贷','',9500.00,'立信大华工程咨询有限责任公司'),
    ('20260312','汇入汇款','转账',6000.00,'安徽子啸建设工程有限公司'),
    ('20260312','对公提回贷','',3500.00,'武汉云克隆科技股份有限公司'),
    ('20260312','汇入汇款','认证费',9000.00,'湖北昌盛家具有限公司'),
    ('20260312','对公转账出','服务费（瑞海达）',-2350.00,'武汉艾思欧科技有限公司'),
    ('20260312','对公转账出','服务费（襄阳四凯）',-1034.00,'合肥好助手商务咨询服务有限公司'),
    ('20260313','对公提回贷','安陆市鸿兴达包装支付款',10000.00,'安陆市鸿兴达包装制品有限公司'),
    ('20260313','存现现金','',24000.00,'湖北流浪舱文旅集团有限公司'),
    ('20260313','汇入汇款','鉴证咨询服务认证费',10000.00,'湖北鼎祥钢铁炉料有限公司'),
    ('20260313','对公提回贷','无附言',5000.00,'宜昌市江星电气机械有限公司'),
    ('20260316','汇入汇款','易方嘉ISO三标认证续费',11000.00,'湖北易方嘉建设工程有限公司'),
    ('20260316','对公提回贷','',10500.00,'湖北勇业商贸有限公司'),
    ('20260316','对公转账出','服务费（流浪舱）',-11750.00,'武汉韵恒科技咨询有限公司'),
    ('20260316','对公转账出','服务费（俊楠）',-3800.00,'资质录（武汉）信息技术有限公司'),
    ('20260316','工资','2月工资',-34082.97,'代发业务资金过渡专户'),
    ('20260316','汇入汇款','服务费',13000.00,'湖北金利源大悟纺织服装有限公司'),
    ('20260316','对公提回贷','',13000.00,'湖北沛林生态工程有限公司'),
    ('20260317','税款','实时缴税',-14596.11,'暂收款'),
    ('20260318','汇入汇款','服务费',13000.00,'武汉乐为物业管理有限公司'),
    ('20260318','汇入汇款','体系审核费用',4800.00,'湖北初和汽车部件有限公司'),
    ('20260318','汇入汇款','服务费',10000.00,'湖北永阳光明电力有限公司'),
    ('20260319','对公提回贷','认证尾款',2400.00,'湖北楚食安供应链科技有限公司'),
    ('20260319','对公提回贷','转账汇款',8000.00,'武汉有教科技有限公司'),
    ('20260320','对公提回贷','转账汇款',6500.00,'武汉中坤商品混凝土有限公司'),
    ('20260320','汇入汇款','',14000.00,''),
    ('20260320','汇入汇款','年审费用',10000.00,'沃克豪斯（武汉）新材料科技有限公司'),
    ('20260320','汇入汇款','服务费',10000.00,'武汉鑫沃斯机器人工程有限公司'),
    ('20260321','账户结息','收息',61.71,'应付利息-单位活期存款利息'),
    ('20260323','汇入汇款','2026年体系认证费用',15000.00,'湖北爱道威建设工程有限公司'),
    ('20260323','汇入汇款','五星售后证书费用',2000.00,'江西景兴智能科技有限公司'),
    ('20260323','汇入汇款','iso安全管理体系认证费',5000.00,'武汉洪体游泳馆运营管理有限公司'),
    ('20260324','对公转账出','退款',-7000.00,'武汉咨元信号技术有限公司'),
    ('20260324','对公转账出','服务费',-10351.48,'张尤佳'),
    ('20260324','对公转账出','服务费（咨元信号）',-570.00,'武汉谊笙成达企业服务有限公司'),
    ('20260324','汇入汇款','认证费',6000.00,'湖北本杰建设工程有限公司'),
    ('20260324','汇入汇款','服务费',12000.00,'锐城科技股份有限公司'),
    ('20260325','对公提回贷','2026年ISO证书审核费用',12500.00,'湖北亿能电力建设有限公司'),
    ('20260325','汇入汇款','IOS年审费',3500.00,'武汉博莱恩智能装备有限公司'),
    ('20260325','汇入汇款','咨询服务费',10000.00,'湖北恒平汽车运输有限公司'),
    ('20260325','汇入汇款','代工',10000.00,'荆州市正泰电气销售有限公司'),
    ('20260326','汇入汇款','iso体系认证',5000.00,'长沙城北贸易有限责任公司'),
    ('20260326','汇入汇款','ISO年审费用',3500.00,'武汉博锐特机电设备有限公司'),
    ('20260326','汇入汇款','',12750.00,'苏州照明科技有限公司'),
    ('20260327','对公提回贷','ISO年审费',8000.00,'湖北科姆森科技服务有限公司'),
    ('20260327','对公转账出','服务费（科姆森）',-2820.00,'武汉一诺前景企业服务有限公司'),
    ('20260327','汇入汇款','ISO认证',15000.00,'宜昌城大建设有限公司'),
    ('20260330','对公提回贷','',10000.00,'武汉林源装饰工程有限公司'),
    ('20260330','汇入汇款','采购款证书年检',8000.00,'湖北韦中工贸有限公司'),
    ('20260331','对公提回贷','鉴证咨询服务认证费',7000.00,'湖北慧联科建设工程有限公司'),
    ('20260331','对公提回贷','无附言',4000.00,'黄梅华胜建材有限公司'),
    ('20260331','对公转账出','差旅报销',-2645.00,'陈曼'),
    ('20260331','对公转账出','差旅报销',-689.00,'罗瑜'),
    ('20260331','对公转账出','差旅报销',-1122.00,'宋智超'),
    ('20260331','对公转账出','差旅报销',-2905.95,'宋丹丹'),
    ('20260331','对公转账出','服务费（苏州照明）',-11280.00,'苏州景宽企业管理有限公司'),
    ('20260331','对公转账出','服务费（五峰立达、宜昌楚邦）',-4700.00,'宜昌质飞咨询有限公司'),
    ('20260331','对公转账出','服务费（星驰激光）',-1880.00,'武汉风生水起科技有限公司'),
]

fees_bank = [
    ('20260302','网银费用','跨行异地手续费',8.76,''),
    ('20260302','网银费用','跨行本地普通手续费',2.65,''),
    ('20260302','网银费用','网上企业银行服务费',25.00,''),
    ('20260311','对公转账出','服务费（德希）',895.00,'武汉璞然知识产权代理有限公司'),
    ('20260311','对公转账出','服务费（能科）',2670.00,'武汉盛凡空间信息技术有限公司'),
    ('20260312','对公转账出','服务费（瑞海达）',2350.00,'武汉艾思欧科技有限公司'),
    ('20260312','对公转账出','服务费（襄阳四凯）',1034.00,'合肥好助手商务咨询服务有限公司'),
    ('20260316','对公转账出','服务费（流浪舱）',11750.00,'武汉韵恒科技咨询有限公司'),
    ('20260316','对公转账出','服务费（俊楠）',3800.00,'资质录（武汉）信息技术有限公司'),
    ('20260324','对公转账出','服务费',10351.48,'张尤佳'),
    ('20260324','对公转账出','服务费（咨元信号）',570.00,'武汉谊笙成达企业服务有限公司'),
    ('20260327','对公转账出','服务费（科姆森）',2820.00,'武汉一诺前景企业服务有限公司'),
]

inv_amts = [24000.0,16000.0,15000.0,14000.0,13000.0,12750.0,12500.0,12000.0,11000.0,10500.0,10000.0,9500.0,9000.0,8000.0,7000.0,5000.0,4800.0,4000.0,3500.0,2400.0]
inv_pos = set(inv_amts)

credits_all = [(d,t,desc,a,c) for d,t,desc,a,c in bank_data if a > 0]
debits_all  = [(d,t,desc,abs(a),c) for d,t,desc,a,c in bank_data if a < 0]

bank_by_co = defaultdict(lambda: {'total':0.0,'count':0,'dates':[]})
for d,t,desc,a,c in credits_all:
    if c and c not in ('应付利息-单位活期存款利息',):
        bank_by_co[c]['total'] += a
        bank_by_co[c]['count'] += 1
        bank_by_co[c]['dates'].append(d)

inv_count = 52
inv_total_sum = 430100.00

s1h = ['序号','公司名称','当月收款金额','收款笔数','收款日期','匹配发票金额','发票张数','备注']
s1r = []
n = 1
for co, v in sorted(bank_by_co.items(), key=lambda x: x[1]['total'], reverse=True):
    amt_key = round(v['total'], 2)
    is_matched = amt_key in inv_pos
    note = '已匹配' if is_matched else '未匹配到发票'
    s1r.append([n, co, v['total'], v['count'], ','.join(sorted(set(v['dates']))), v['total'] if is_matched else '', '1' if is_matched else '', note])
    n += 1

s2h = ['日期','业务类型','摘要','金额','对方账户']
s2r = sorted(fees_bank, key=lambda x: x[0])

s3h = ['日期','收款金额','对方公司名称','摘要']
s3r = sorted([(d,a,c,desc) for d,t,desc,a,c in credits_all if round(a,2) not in inv_pos and c and c not in ('应付利息-单位活期存款利息',)], key=lambda x: x[0])

s4h = ['日期','业务类型','摘要','金额','对方账户']
s4r = sorted(debits_all, key=lambda x: x[0])

sheets = [
    ('Sheet1-发票收款对比', s1h, s1r),
    ('Sheet2-银行手续费明细', s2h, s2r),
    ('Sheet3-收款未开票', s3h, s3r),
    ('Sheet4-其他支出明细', s4h, s4r),
]

out_path = r'C:\Users\wdxyh\Desktop\银行对账与发票核对_2026年03月_v2.xlsx'
make_xlsx(out_path, sheets)
bank_total = sum(v['total'] for v in bank_by_co.values())
matched_n = sum(1 for co,v in bank_by_co.items() if round(v['total'],2) in inv_pos)
unmatched_n = sum(1 for co,v in bank_by_co.items() if round(v['total'],2) not in inv_pos)
print('Invoice:', inv_count, '张', inv_total_sum, 'Bank receipt:', bank_total, 'Matched:', matched_n, 'Unmatched:', unmatched_n)
