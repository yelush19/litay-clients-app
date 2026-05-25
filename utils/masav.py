import io
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter, quote_sheetname

GREEN = '4CAF50'

def format_date(dt):
    return f"{dt.day:02d}/{dt.month:02d}/{dt.year}"

def parse_amount(s):
    import re
    c = re.sub(r"[\$,\s]","",str(s or ""))
    c = re.sub(r"[()]","",c)
    try: return abs(float(c))
    except: return 0.0

BLUE  = '1565C0'

# ===== MASAV PARSER + BUILDER =====
def read_vendor_index(index_json):
    return index_json if isinstance(index_json, dict) else {}

def parse_masav(file_bytes, vendor_lookup, bank_coa):
    rows,batches,errors,unmatched=[],{},[],[]
    try:
        wb=load_workbook(io.BytesIO(file_bytes),data_only=True)
        ws=wb.active
        header=[str(c.value or '') for c in list(ws.iter_rows(min_row=1,max_row=1))[0]]
        def col(name): return next((i for i,h in enumerate(header) if name in h),None)
        cv=col('שם ספק'); ca=col('סכום'); chp=col('ח.פ'); cd=col('תאריך')
        cpd=col('שולם'); cb=col('Batch'); ci=col('חשבונית')
        cbk=col('בנק'); cbr=col('מספר סניף'); cac=col('מספר חשבון')
        for ri,row in enumerate(ws.iter_rows(min_row=2,max_row=ws.max_row,values_only=True),2):
            vendor=str(row[cv] or '').strip()
            if not vendor or vendor=='סך הכל לתשלום': continue
            paid=row[cpd]
            if paid is not True and str(paid).upper() not in ('TRUE','1','כן'): continue
            raw_date=row[cd]
            dt=raw_date if isinstance(raw_date,datetime) else None
            if not dt:
                for fmt in ('%d/%m/%Y','%Y-%m-%d','%d.%m.%Y'):
                    try: dt=datetime.strptime(str(raw_date),fmt); break
                    except: pass
            if not dt: errors.append(f"שורה {ri}: תאריך לא תקין"); continue
            try: amount=float(str(row[ca] or '').replace(',','').replace('₪','').strip())
            except: errors.append(f"שורה {ri}: סכום לא תקין"); continue
            hp_raw=row[chp]
            hp=str(int(hp_raw)) if isinstance(hp_raw,float) else str(hp_raw or '').strip()
            invoice=str(row[ci] or '') if ci is not None else ''
            batch=str(row[cb] or 'ללא batch') if cb is not None else 'ללא batch'
            bank_d=f"{row[cbk] or ''}-{row[cbr] or ''}-{row[cac] or ''}" if cbk is not None else ''
            vendor_coa=vendor_lookup.get(hp,'')
            if not vendor_coa:
                unmatched.append({'vendor':vendor,'hp':hp,'amount':amount,
                                  'date':format_date(dt)})
            rec={'date':dt,'date_fmt':format_date(dt),'vendor':vendor,'amount':amount,
                 'invoice':invoice,'hp':hp,'vendor_coa':vendor_coa,'bank_coa':bank_coa,
                 'bank_detail':bank_d,'batch':batch}
            rows.append(rec)
            if batch not in batches: batches[batch]={'total':0,'count':0}
            batches[batch]['total']+=amount; batches[batch]['count']+=1
    except Exception as e:
        errors.append(f"שגיאה: {e}")
    rows.sort(key=lambda r:(r['batch'],r['date']))
    return rows,batches,errors,unmatched

def build_masav_excel(rows,batches):
    wb=Workbook(); ws=wb.active; ws.title='MASAV'
    headers=['תאריך','תיאור','חובה','זכות','אסמכתא 1','אסמכתא 2','ח"ן חובה','ח"ן זכות','Batch','פרטי בנק ספק']
    ws.append(headers)
    hf=PatternFill(start_color=GREEN,end_color=GREEN,fill_type="solid")
    hfont=Font(bold=True,color="FFFFFF")
    for cell in ws[1]: cell.fill=hf; cell.font=hfont; cell.alignment=Alignment(horizontal="right")
    for col,w in zip("ABCDEFGHIJ",[12,35,14,14,22,14,12,12,10,26]):
        ws.column_dimensions[col].width=w
    current_batch=None; batch_start=2
    def add_summary(bkey,start):
        end=ws.max_row
        if end<start: return
        ws.append(['',f'סך הכל batch {bkey}',f'=SUM(C{start}:C{end})',f'=SUM(D{start}:D{end})','','','','','',''])
        r=ws.max_row
        for c in (3,4): ws.cell(r,c).number_format='#,##0.00'
        sf=Font(bold=True); sfill=PatternFill(start_color='E8F5E9',end_color='E8F5E9',fill_type='solid')
        for c in range(1,11): ws.cell(r,c).font=sf; ws.cell(r,c).fill=sfill
    for t in rows:
        if t['batch']!=current_batch:
            if current_batch is not None: add_summary(current_batch,batch_start)
            current_batch=t['batch']; batch_start=ws.max_row+1
        ws.append([t['date_fmt'],t['vendor'],t['amount'],t['amount'],
                   t['invoice'],t['hp'],t['vendor_coa'],t['bank_coa'],t['batch'],t['bank_detail']])
        r=ws.max_row
        for c in (3,4): ws.cell(r,c).number_format='#,##0.00'
        for c in (5,6,7,8,9):
            cell=ws.cell(r,c); cell.value=str(cell.value or ''); cell.number_format='@'
        if not t['vendor_coa']:
            ws.cell(r,7).fill=PatternFill(start_color='FFF9C4',end_color='FFF9C4',fill_type='solid')
    if current_batch is not None: add_summary(current_batch,batch_start)
    tr=ws.max_row+1; total=sum(t['amount'] for t in rows)
    ws.append(['','סה״כ כולל',total,total,'','','','','',''])
    for c in (3,4): ws.cell(tr,c).number_format='#,##0.00'
    tf=Font(bold=True,color='FFFFFF'); tfill=PatternFill(start_color=BLUE,end_color=BLUE,fill_type='solid')
    for c in range(1,11): ws.cell(tr,c).font=tf; ws.cell(tr,c).fill=tfill
    if rows:
        ref=f"{quote_sheetname('MASAV')}!$A$2:$J${ws.max_row}"
        wb.defined_names.add(DefinedName('MASAV',attr_text=ref))
    no_coa=sum(1 for r in rows if not r['vendor_coa'])
    return wb,{'total_rows':len(rows),'total_amount':total,
               'no_coa':no_coa,'coa_covered':len(rows)-no_coa}