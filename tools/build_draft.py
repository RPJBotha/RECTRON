from pathlib import Path
import csv, collections, datetime, json, re, zipfile, xml.etree.ElementTree as ET
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table,TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.workbook.properties import CalcProperties
ROOT=Path('/workspace/attachments'); OUT=Path('/workspace/RECTRON/output'); OUT.mkdir(exist_ok=True)
def csvrows(name):
 p=next(ROOT.glob('*/'+name)); return list(csv.DictReader(p.open(encoding='utf-8-sig',newline='')))
def num(x):
 try: return float(str(x).replace('$','').replace(',','').strip())
 except: return None
def positive(x):
 v=num(x); return v if v is not None and v>0 else None
def key(x): return str(x or '').strip()
soh=csvrows('SOH.csv'); po=csvrows('PO.csv'); orders=csvrows('OPEN ENT.csv'); sales=csvrows('SALES.csv'); receipts=csvrows('STUDIO USD.csv')
cip=list(openpyxl.load_workbook(ROOT/'c3f0fe55-77c9-43bc-a6ca-a9c45af0531e/CIP LIST.xlsx',data_only=True).active.values)
pl=openpyxl.load_workbook(next(ROOT.glob('*/DJI-PL.xlsx')),data_only=True)
w=openpyxl.Workbook();w.remove(w.active);w.calculation=CalcProperties(calcId=191029,fullCalcOnLoad=True)
cache={}; tables=0
def sheet(name,headers,rows):
 global tables
 s=w.create_sheet(name);s.append(headers)
 for r in rows:s.append(list(r))
 s.freeze_panes='C2';s.auto_filter.ref=s.dimensions;s.sheet_view.showGridLines=False
 for c in s[1]:c.fill=PatternFill('solid',fgColor='073349');c.font=Font(color='FFFFFF',bold=True);c.alignment=Alignment(wrap_text=True)
 s.row_dimensions[1].height=42
 for col in s.columns:
  letter=col[0].column_letter;s.column_dimensions[letter].width=min(42,max(14,len(str(col[0].value))+2))
 if s.max_row>1:
  tables+=1;t=Table(displayName='Data'+str(tables),ref=s.dimensions);t.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=True);s.add_table(t)
 return s
def formula(s,r,c,f,value):
 s.cell(r,c,f);cache[(s.title,s.cell(r,c).coordinate)]=value
notes=[
('Purpose','Internal draft • all monetary selling prices and commercial costs exclude VAT unless explicitly labelled.'),
('Workflow','Review Settings, Product Master mappings and Review flags; enter blue override cells; open in Excel with automatic calculation.'),
('Refresh limitation','This draft embeds the supplied reports. Power Query / automated refresh is not configured. Rerunning the builder recreates the draft and does not preserve manual edits; save an edited copy first.'),
('Pricing','S=cost/0.90; A=cost/0.85; B=cost/0.80; OTHER=cost/0.75. DJI class prices are references, not replacements for these baselines.'),
('Cost selection','Manual actual ZAR cost overrides estimate; otherwise selected USD cost × cost ROE × (1 + handling). SOH and receipt costs are references pending logistics confirmation.'),
('Supplier precedence','Manual USD override > unique positive CIP portal DP > unique positive DJI-PL CIP. FCA is not automatically substituted. RRP uses CIP then DJI-PL.'),
('RRP fallback','45% target gross margin is a provisional editable assumption, not a confirmed policy. Public selling ROE is separate from cost ROE.'),
('VAT','Tax-inclusive supplier RRP is divided by 1+VAT. Missing portal tax flags are not assumed; portal RRP is ignored until resolved. VAT setting is provisional 15%.'),
('Matching','Exact unique supplier part number links SOH to CIP; EAN links DJI-PL. No fuzzy matching. Duplicate source keys require review.'),
('Scope','806 SOH products classified DJI Enterprise plus supplier products not mapped to those codes. Agriculture and Consumer remain in raw reports only.'),
('Dates','Report export timestamps are not supplied. PO ETA dates are supplier estimates. No report is assumed current solely because it was uploaded.'),
('Orders','PO remaining=max(ordered-received,0). Open customer quantities and backorders are shown separately; do not add them or add SOH backorders again.'),
('Sales','Net units and value retain credit-note signs. Trailing 90-day period ends at the editable analysis date, not the latest transaction date.'),
('Combos','Allocate USD unit costs on Combo Allocations by order/receipt. Extended component allocations must reconcile to combo total; allocations do not globally override a component.'),
('Dealer sharing','Do not share this workbook with dealers: raw reports contain customer and cost information. Export approved dealer prices to a separate values-only workbook.'),
('Draft status','Not approved for ERP upload or dealer publication. Resolve mappings, assumptions and review flags first.'),
]
sheet('Read Me',['Topic','Guidance'],notes)
settings=sheet('Settings',['Parameter','Value','Meaning'],[
('Cost ROE',16.7,'Provisional ZAR per USD; screenshot reference'),('Selling ROE',17,'Provisional ZAR per USD for supplier RRP'),('Handling',.06,'Additional landed-cost allowance; confirm scope for CIP'),('VAT',.15,'Used for displayed VAT and supplier RRP normalization'),('S margin',.10,'Rectron gross margin'),('A margin',.15,'Rectron gross margin'),('B margin',.20,'Rectron gross margin'),('OTHER margin',.25,'Rectron gross margin'),('Fallback RRP margin',.45,'PROVISIONAL gross margin; confirm versus markup'),('Analysis date',datetime.datetime(2026,11,10),'Change to the reporting date'),('Sales window days',90,'Trailing days including analysis date')])
for r in range(2,11):settings.cell(r,2).number_format='0.00' if r<4 else '0.0%'
settings['B11'].number_format='yyyy-mm-dd'
for r in range(2,13):settings.cell(r,2).font=Font(color='0070C0')
ci=collections.defaultdict(list); ei=collections.defaultdict(list)
for r in cip[1:]:
 ci[key(r[4])].append(r);ei[key(r[5])].append(r)
pi=collections.defaultdict(list)
for title in ['single','Other Product','demo']:
 for r in list(pl[title].values)[1:]:
  e=key(r[1] if title=='Other Product' else r[0])
  if not e:continue
  if title=='single': d={'ean':e,'name':r[1],'cost':positive(r[5]),'rrp':positive(r[7]),'classes':r[14:18],'source':title}
  elif title=='Other Product': d={'ean':e,'name':r[2],'cost':positive(r[6]) if key(r[0])=='CIP' else None,'rrp':positive(r[4]),'classes':[],'source':title}
  else:d={'ean':e,'name':r[1],'cost':positive(r[3]),'rrp':positive(r[4])/1.15 if positive(r[4]) else None,'classes':[],'source':title}
  pi[e].append(d)
# Prefer standard single products over other/demo; multiple records within a selected sheet remain ambiguous.
def plmatch(e):
 a=pi.get(e,[])
 for title in ['single','Other Product','demo']:
  b=[x for x in a if x['source']==title]
  if b:return b[0] if len(b)==1 else None
 return None
stock={key(r['ItemId']):r for r in soh if r['ProductLine']=='DJI Enterprise'}
products=[];used=set()
for code,s in stock.items():
 a=ci.get(key(s['Supplier Code']),[]);r=a[0] if len(a)==1 else None
 if r:used.add(key(r[5]))
 products.append((code,key(s['Supplier Code']),key(r[5]) if r else '',s['Description'],s,r,'Unique part match' if r else 'CHECK mapping'))
for e,rs in ei.items():
 if e in used:continue
 r=rs[0] if len(rs)==1 else None
 products.append(('',key(rs[0][4]),e,key(rs[0][8]),{},r,'Supplier only' if r else 'CHECK duplicate / combo'))
for e in pi:
 if e not in ei:
  d=plmatch(e);products.append(('', '',e,d['name'] if d else e,{},None,'DJI-PL only' if d else 'CHECK duplicate DJI-PL'))
# aggregate fact reports independently to avoid multiplying rows
pa=collections.defaultdict(lambda:[0,None,None]);oa=collections.defaultdict(lambda:[0,0]);sa=collections.defaultdict(lambda:[0,0]);lr={}
for r in po:
 k=key(r['ItemId']);q=max((num(r['Qty Ord']) or 0)-(num(r['Qty Recv']) or 0),0);pa[k][0]+=q
 if q and r['ETA'] and (pa[k][1] is None or r['ETA']<pa[k][1]):pa[k][1]=r['ETA'];pa[k][2]=num(r['Price'])
for r in orders:oa[key(r['ItemID'])][0]+=num(r['QtyRequired']) or 0;oa[key(r['ItemID'])][1]+=num(r['QtyBackordered']) or 0
for r in receipts:
 k=key(r['ItemID'])
 if k not in lr or r['DateCreated']>lr[k]['DateCreated']:lr[k]=r
# retain daily facts so analysis-date edits recalculate sales windows
sale_daily=collections.defaultdict(lambda:[0,0])
for r in sales:
 if r['Product Line']!='DJI Enterprise':continue
 try:d=datetime.datetime.strptime(r['Date'][:10],'%Y/%m/%d')
 except ValueError:continue
 k=(key(r['ItemID']),d);sale_daily[k][0]+=num(r['Qty']) or 0;sale_daily[k][1]+=num(r['Amount']) or 0
sd=sheet('Sales Daily',['Rectron code','Date','Net units','Revenue ExVAT'],[(k[0],k[1],v[0],v[1]) for k,v in sorted(sale_daily.items())])
for r in range(2,sd.max_row+1):sd.cell(r,2).number_format='yyyy-mm-dd'
headers=['Rectron code','DJI part number','EAN / combo ID','Description','Mapping status','Portal CIP USD','DJI-PL CIP USD','Manual USD override','Selected USD','Cost source','Actual landed ZAR override','Estimated cost ExVAT','Cost used ExVAT','SOH actual cost ref','Last receipt ZAR ref','Supplier RRP USD ExVAT','Manual public ExVAT','Public RRP ExVAT','S ExVAT','A ExVAT','B ExVAT','OTHER ExVAT','Public incl VAT','S dealer potential margin','SOH total','Open order qty','Backorder qty','Outstanding PO qty','Next ETA','Next PO USD ref','Net units sales window','Revenue sales window','Indicative stock cover days','Review','Override reason','DJI S USD ref','DJI A USD ref','DJI B USD ref','DJI OTHER USD ref']
m=sheet('MPL Internal',headers,[]); now=datetime.datetime(2026,11,10); end=sd.max_row
for i,(code,part,e,name,s,r,status) in enumerate(products,2):
 d=plmatch(e);pc=positive(r[17]) if r else None;dc=d['cost'] if d else None
 rr=None
 if r and positive(r[19]) and key(r[20])=='USD':
  if key(r[21])=='Yes':rr=positive(r[19])/1.15
  elif key(r[21]).lower() in ('no','without tax'):rr=positive(r[19])
 if rr is None and d:rr=d['rrp']
 usd=pc or dc;estimate=usd*16.7*1.06 if usd else None;pub=rr*17 if rr else estimate/.55 if estimate else None
 flags=[]
 if not code:flags.append('NO RECTRON CODE')
 if status.startswith('CHECK'):flags.append(status)
 if not usd:flags.append('MISSING COST')
 if not rr:flags.append('FALLBACK RRP' if estimate else 'MISSING RRP')
 if d is None and e in pi:flags.append('DJI-PL DUPLICATE')
 q=positive(s.get('SOH_Tot')) or 0; p=pa[code];o=oa[code];latest=lr.get(code,{})
 daily=[(k,v) for k,v in sale_daily.items() if k[0]==code and now-datetime.timedelta(days=89)<=k[1]<=now]
 units=sum(v[0] for k,v in daily); revenue=sum(v[1] for k,v in daily)
 vals=[code,part,e,name,status,pc,dc,None,None,None,None,None,None,num(s.get('ActualCost')),num(latest.get('Price')),rr,None,None,None,None,None,None,None,None,num(s.get('SOH_Tot')),o[0],o[1],p[0],p[1],p[2],None,None,None,None,None]+list(d['classes'] if d else [])
 m.append(vals)
 f=lambda c,expr,v:formula(m,i,c,expr,v)
 f(9,f'=IF(ISNUMBER(H{i}),H{i},IF(ISNUMBER(F{i}),F{i},IF(ISNUMBER(G{i}),G{i},"")))',usd)
 f(10,f'=IF(ISNUMBER(H{i}),"Manual",IF(ISNUMBER(F{i}),"CIP portal",IF(ISNUMBER(G{i}),"DJI-PL CIP","MISSING")))','CIP portal' if pc else 'DJI-PL CIP' if dc else 'MISSING')
 f(12,f'=IF(ISNUMBER(I{i}),ROUND(I{i}*Settings!$B$2*(1+Settings!$B$4),2),"")',round(estimate,2) if estimate else None)
 f(13,f'=IF(ISNUMBER(K{i}),K{i},L{i})',round(estimate,2) if estimate else None)
 f(18,f'=IF(ISNUMBER(Q{i}),Q{i},IF(ISNUMBER(P{i}),ROUND(P{i}*Settings!$B$3,2),IF(ISNUMBER(M{i}),ROUND(M{i}/(1-Settings!$B$10),2),"")))',round(pub,2) if pub else None)
 for c,sr in [(19,6),(20,7),(21,8),(22,9)]:
  margin={19:.1,20:.15,21:.2,22:.25}[c];f(c,f'=IF(ISNUMBER(M{i}),ROUND(M{i}/(1-Settings!$B${sr}),2),"")',round(round(estimate,2)/(1-margin),2) if estimate else None)
 f(23,f'=IF(ISNUMBER(R{i}),ROUND(R{i}*(1+Settings!$B$5),2),"")',round(round(pub,2)*1.15,2) if pub else None)
 f(24,f'=IFERROR((R{i}-S{i})/R{i},"")',(pub-estimate/.9)/pub if pub and estimate else None)
 bounds=f"'Sales Daily'!$B$2:$B${end},\">=\"&Settings!$B$11-Settings!$B$12+1,'Sales Daily'!$B$2:$B${end},\"<=\"&Settings!$B$11"
 for col,src,val in [(31,'C',units),(32,'D',revenue)]:f(col,f'=IF(A{i}="","",SUMIFS(\'Sales Daily\'!${src}$2:${src}${end},\'Sales Daily\'!$A$2:$A${end},A{i},{bounds}))',val if code else None)
 f(33,f'=IFERROR(IF(AE{i}>0,Y{i}/(AE{i}/Settings!$B$12),""),"")',q/(units/90) if units>0 else None)
 static='; '.join(flags)
 f(34,f'="{static}"&IF(AND(ISNUMBER(R{i}),ISNUMBER(V{i}),R{i}<V{i}),"; PUBLIC BELOW OTHER","")&IF(OR(AND(ISNUMBER(H{i}),H{i}<=0),AND(ISNUMBER(K{i}),K{i}<=0),AND(ISNUMBER(Q{i}),Q{i}<=0)),"; INVALID OVERRIDE","")&IF(AND(OR(ISNUMBER(H{i}),ISNUMBER(K{i}),ISNUMBER(Q{i})),AI{i}=""),"; ADD OVERRIDE REASON","")',static+('; PUBLIC BELOW OTHER' if pub and estimate and pub<estimate/.75 else ''))
 for c in [8,11,17,35]:m.cell(i,c).fill=PatternFill('solid',fgColor='E5F2FF');m.cell(i,c).font=Font(color='0070C0')
 for c in list(range(6,24))+[30,32,33]:m.cell(i,c).number_format='#,##0.00;[Red](#,##0.00);–'
 m.cell(i,24).number_format='0.0%'
for col in ['H','K','Q']:
 dv=DataValidation(type='decimal',operator='greaterThan',formula1=0,allow_blank=True);dv.showErrorMessage=True;dv.error='Enter a positive value, or clear the cell.';m.add_data_validation(dv);dv.add(f'{col}2:{col}{m.max_row}')
m.auto_filter.ref=m.dimensions
m.conditional_formatting.add(f'X2:X{m.max_row}',CellIsRule(operator='lessThan',formula=['0'],fill=PatternFill('solid',fgColor='FFC7CE')))
# Combo allocations preloaded from DJI list, with manually entered unit allocations.
combo=[]
for r in list(pl['combo'].values)[1:]:
 if not r[0]:continue
 if positive(r[6]):total=positive(r[6]);comboid=key(r[0]);comboname=r[1]
 elif r[2]:combo.append([comboid,comboname,'',key(r[2]),r[4],num(r[3]),total,None,None,None,''])
a=sheet('Combo Allocations',['Combo ID','Combo name','PO / receipt reference','Component EAN','Component name','Quantity','Combo total USD','Manual USD unit allocation','Extended USD','Unallocated USD','Notes'],combo)
for i in range(2,a.max_row+1):
 formula(a,i,9,f'=IF(ISNUMBER(H{i}),F{i}*H{i},"")',None)
 formula(a,i,10,f'=G{i}-SUMIFS(I$2:I${a.max_row},A$2:A${a.max_row},A{i},C$2:C${a.max_row},C{i})',a.cell(i,7).value)
 for c in [3,8,11]:a.cell(i,c).fill=PatternFill('solid',fgColor='E5F2FF');a.cell(i,c).font=Font(color='0070C0')
review=sheet('Review',['Review item','Result'],[
 ('Products in draft',len(products)),('Enterprise SOH products',len(stock)),('Review flags','Filter Review column on MPL Internal; blank means no automatic flag, not approval.'),('Manual mapping','Mappings are draft exact matches. EAN and part numbers stored as text.'),('Cost approval','Confirm 6% allowance, ROEs, 45% fallback gross margin and use of actual logistics cost.'),('Refresh','Embedded snapshot only; automated refresh is a next phase.'),('Dealer export','Separate dealer-safe values-only export not yet generated; draft requires approval first.')])
sheet('Source CIP',list(cip[0]),cip[1:])
for s in pl:
 rows=list(s.values);heads=[str(v or 'Unused '+str(i+1)) for i,v in enumerate(rows[0])];sheet('DJI '+s.title,heads,rows[1:])
for title,rs in [('Source SOH',soh),('Source PO',po),('Source Open Orders',orders),('Source Sales',sales),('Source Receipts',receipts)]:sheet(title,list(rs[0]),[list(r.values()) for r in rs])
# Protect formula-looking source text from execution; only formulas registered above are permitted.
for s in w:
 for row in s:
  for c in row:
   if c.data_type=='f' and (s.title,c.coordinate) not in cache:c.data_type='s'
for s in w:
 s.sheet_properties.pageSetUpPr.fitToPage=True
 s.sheet_properties.outlinePr.summaryRight=False
 s.sheet_properties.tabColor='229EAD' if s.title in ['MPL Internal','Settings','Combo Allocations'] else '718096'
w.active=w.sheetnames.index('Read Me');target=OUT/'DJI_Enterprise_MPL_DRAFT.xlsx';w.save(target)
# Supply initial cached formula results, with Excel recalculation enabled for edits.
ns={'x':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'};temp=target.with_suffix('.tmp')
with zipfile.ZipFile(target) as src,zipfile.ZipFile(temp,'w',zipfile.ZIP_DEFLATED) as dst:
 for item in src.infolist():
  data=src.read(item.filename)
  match=re.fullmatch(r'xl/worksheets/sheet(\d+).xml',item.filename)
  if match:
   sn=w.sheetnames[int(match[1])-1];root=ET.fromstring(data)
   for c in root.findall('.//x:c',ns):
    k=(sn,c.attrib['r'])
    if k not in cache:continue
    v=c.find('x:v',ns)
    if v is None:v=ET.SubElement(c,'{'+ns['x']+'}v')
    val=cache[k]
    if isinstance(val,str):c.set('t','str');v.text=val
    elif val is None:c.attrib.pop('t',None);v.text=None
    else:c.attrib.pop('t',None);v.text=str(val)
   data=ET.tostring(root,encoding='utf-8',xml_declaration=True)
  dst.writestr(item,data)
temp.replace(target)
# Structural and pricing checks against saved workbook.
f=openpyxl.load_workbook(target,data_only=False);v=openpyxl.load_workbook(target,data_only=True)
assert len(f['MPL Internal']['A'])-1==len(products)
checked=0
for row in v['MPL Internal'].iter_rows(min_row=2):
 cost=row[12].value
 if isinstance(cost,(float,int)) and cost>0:
  for idx,margin in [(18,.1),(19,.15),(20,.2),(21,.25)]:assert abs(row[idx].value-round(cost/(1-margin),2))<.011
  checked+=1
assert sum(c.data_type=='f' for s in f for row in s for c in row)==len(cache)
print(json.dumps({'file':str(target),'products':len(products),'priced_products_checked':checked,'formula_cells':len(cache),'sheets':len(w.sheetnames),'bytes':target.stat().st_size},indent=2))
