import re
import pdfplumber
import pandas as pd
from io import BytesIO
from openpyxl import Workbook


def parse_bank_account(account_str):
    parts = account_str.strip().split('-')
    if len(parts) == 3:
        return parts[0], parts[1], parts[2]
    return '', '', account_str


def extract_pdf_summary(pdf_file) -> dict:
    """חולץ נתוני סיכום מה-PDF לאימות"""
    import re
    summary = {"total": 0.0, "fund_count": 0, "funds": []}
    try:
        with pdfplumber.open(pdf_file) as pdf:
            all_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    all_text += text + "\n"

        # חלץ סה"כ לסליקה
        total_match = re.search(r'(\d{1,3}(?:,\d{3})*\.\d{2})\s*(?:00\s*)?(?:\d{1,3}(?:,\d{3})*\.\d{2}\s*){1,2}\s*(?:0\.00\s*)?$',
                                all_text, re.MULTILINE)
        # חפש בשורת סיכום
        lines = all_text.split('\n')
        for line in lines:
            amounts = re.findall(r'(\d{1,3}(?:,\d{3})*\.\d{2})', line)
            if amounts and len(amounts) >= 3:
                try:
                    total = float(amounts[0].replace(',',''))
                    if 1000 < total < 10000000:
                        summary["total"] = total
                except: pass

        # ספור שורות עם חשבון בנק
        accounts = re.findall(r'\d{2}-\d{3}-\d+', all_text)
        summary["fund_count"] = len(set(accounts))
        summary["funds"] = list(set(accounts))
    except Exception as e:
        summary["error"] = str(e)
    return summary


def extract_data_from_payit_pdf(pdf_file):
    data = []
    fund_names = {}
    with pdfplumber.open(pdf_file) as pdf:
        all_text = ""
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_text += text + "\n"
        lines = all_text.split('\n')
        for line in lines:
            name_pattern = r'([\u0590-\u05FF\s]+)\s*-\s*(\d+)\s*$'
            match = re.search(name_pattern, line)
            if match:
                reversed_name = match.group(1).strip()
                fund_number = match.group(2)
                correct_name = reversed_name[::-1]
                correct_name = re.sub(r'\s+קרן\s+(פנסיה|השתלמות)\s*$', '', correct_name)
                fund_names[fund_number] = correct_name.strip()
        current_fund_name = None
        for line in lines:
            if any(keyword in line for keyword in [
                'רזחהל כ"הס', 'םולשתל כ"הס', 'הקילס רושיא',
                'קיסעמ', 'שדוחל', 'סוטטס', 'דומע', 'ןגומ עדימ'
            ]):
                continue
            account_match = re.search(r'(\d{2})-(\d{3})-(\d+)', line)
            if not account_match:
                name_match = re.search(r'([\u0590-\u05FF\s]+)\s*-\s*(\d+)\s*$', line)
                if name_match:
                    fund_number = name_match.group(2)
                    current_fund_name = fund_names.get(fund_number)
                continue
            account = account_match.group(0)
            amounts = re.findall(r'(\d{1,2}),(\d{3})\.(\d{2})', line)
            total_amount = 0
            if amounts:
                total_str = f"{amounts[0][0]}{amounts[0][1]}.{amounts[0][2]}"
                total_amount = float(total_str)
            else:
                simple_amounts = re.findall(r'(?<!\d)(\d{1,4})\.(\d{2})(?!\d)', line)
                if simple_amounts:
                    total_amount = float(f"{simple_amounts[0][0]}.{simple_amounts[0][1]}")
            if total_amount == 0:
                continue
            fund_name = current_fund_name
            inline_fund_match = re.search(r'[\u0590-\u05FF\s]+-\s*(\d+)', line)
            if inline_fund_match:
                fund_number = inline_fund_match.group(1)
                fund_name = fund_names.get(fund_number, current_fund_name)
            if not fund_name:
                clean_line = re.sub(r'\d{2}-\d{3}-\d+', '', line)
                clean_line = re.sub(r'\d{1,2},\d{3}\.\d{2}', '', clean_line)
                clean_line = re.sub(r'\d+\.\d{2}', '', clean_line)
                clean_line = re.sub(r'תיאקנב הרבעה', '', clean_line)
                hebrew_words = re.findall(r'[\u0590-\u05FF]+', clean_line)
                if hebrew_words:
                    reversed_words = [word[::-1] for word in hebrew_words]
                    fund_name = ' '.join(reversed_words[:4]).strip()
                    fund_name = re.sub(r'\s+קרן\s+(פנסיה|השתלמות)\s*$', '', fund_name)
            if not fund_name:
                fund_name = f"קופה {account}"
            data.append({'account': account, 'fund_name': fund_name.strip(), 'amount': str(total_amount)})
    return data


def create_mizrahi_excel(data) -> bytes:
    rows = []
    for item in data:
        bank, branch, account = parse_bank_account(item['account'])
        rows.append({
            'שם המוטב': item['fund_name'],
            'בנק': bank,
            'מספר סניף': branch,
            'מספר חשבון': account,
            'מהות העברה': 'תשלום לספק',
            'סכום (₪)': float(item['amount'])
        })
    df = pd.DataFrame(rows)
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()
