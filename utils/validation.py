"""
validation.py — אימות תקינות לכל הממירים
כלל: אין הורדת קובץ עם שגיאות קריטיות
"""

from dataclasses import dataclass, field
from typing import List


@dataclass
class ValidationResult:
    name: str           # שם הבדיקה
    passed: bool        # עבר / נכשל
    critical: bool      # קריטי = חוסם הורדה
    message: str        # הסבר
    expected: str = ""  # ערך צפוי
    actual: str = ""    # ערך בפועל


@dataclass
class ValidationReport:
    results: List[ValidationResult] = field(default_factory=list)

    def add(self, name, passed, critical, message, expected="", actual=""):
        self.results.append(ValidationResult(name, passed, critical, message, expected, actual))

    @property
    def ok(self) -> bool:
        """האם ניתן להוריד — אין שגיאות קריטיות"""
        return all(r.passed for r in self.results if r.critical)

    @property
    def has_warnings(self) -> bool:
        return any(not r.passed for r in self.results if not r.critical)

    @property
    def critical_failures(self):
        return [r for r in self.results if r.critical and not r.passed]

    @property
    def warnings(self):
        return [r for r in self.results if not r.critical and not r.passed]


# ===== MASAV =====

def validate_masav(rows: list, batches: dict) -> ValidationReport:
    report = ValidationReport()

    # 1. יש שורות
    report.add(
        "שורות נתונים",
        passed=len(rows) > 0,
        critical=True,
        message="לא נמצאו תנועות" if not rows else f"{len(rows)} תנועות",
    )
    if not rows:
        return report

    total = sum(r['amount'] for r in rows)

    # 2. כיסוי ח"ן
    no_coa = [r for r in rows if not r.get('vendor_coa')]
    report.add(
        'כיסוי ח"ן ספקים',
        passed=len(no_coa) == 0,
        critical=True,
        message=f"{len(no_coa)} ספקים ללא ח\"ן" if no_coa else "כל הספקים מזוהים",
        expected="0 ספקים חסרים",
        actual=f"{len(no_coa)} חסרים: {', '.join(r['vendor'] for r in no_coa[:3])}" if no_coa else "✅",
    )

    # 3. סה"כ תואם batches
    if batches:
        batch_total = sum(b['total'] for b in batches.values())
        diff = abs(total - batch_total)
        report.add(
            "סה״כ תואם batches",
            passed=diff < 0.01,
            critical=True,
            message=f"הפרש ₪{diff:,.2f}" if diff >= 0.01 else "✅ תואם",
            expected=f"₪{batch_total:,.2f}",
            actual=f"₪{total:,.2f}",
        )

    # 4. תאריכים תקינים
    invalid_dates = [r for r in rows if not r.get('date_fmt')]
    report.add(
        "תאריכים תקינים",
        passed=len(invalid_dates) == 0,
        critical=True,
        message=f"{len(invalid_dates)} שורות ללא תאריך" if invalid_dates else "✅",
    )

    # 5. אסמכתאות (אזהרה בלבד)
    empty_refs = [r for r in rows if not r.get('invoice')]
    report.add(
        "אסמכתאות מלאות",
        passed=len(empty_refs) == 0,
        critical=False,
        message=f"{len(empty_refs)} שורות ללא אסמכתא" if empty_refs else "✅",
    )

    return report


# ===== VALLEY BANK =====

def validate_valley(data: list, checks: dict) -> ValidationReport:
    report = ValidationReport()

    if not data:
        report.add("שורות נתונים", False, True, "לא נמצאו תנועות")
        return report

    # 1. איזון סכומים
    report.add(
        "איזון סכומים",
        passed=checks.get('amount_ok', False),
        critical=True,
        message="סכומים לא מאוזנים" if not checks.get('amount_ok') else "✅",
        expected=f"${checks.get('total_debits',0)+checks.get('total_credits',0):,.0f}",
        actual=f"${checks.get('total_amount',0):,.0f}",
    )

    # 2. כיסוי ח"ן
    no_coa = checks.get('no_coa', 0)
    total = checks.get('total_count', 1)
    report.add(
        'כיסוי ח"ן',
        passed=no_coa == 0,
        critical=False,
        message=f"{no_coa} תנועות ללא ח\"ן" if no_coa else "✅",
        expected=f"{total}/{total}",
        actual=f"{total-no_coa}/{total}",
    )

    return report


# ===== PAYEM =====

def validate_payem(data: list, checks: dict) -> ValidationReport:
    report = ValidationReport()

    if not data:
        report.add("שורות נתונים", False, True, "לא נמצאו תנועות")
        return report

    # 1. INC + LTD = סה"כ
    report.add(
        "INC + LTD = סה״כ",
        passed=checks.get('count_ok', False),
        critical=True,
        message="חוסר התאמה בספירת שורות" if not checks.get('count_ok') else "✅",
        expected=str(checks.get('count_data', 0)),
        actual=f"INC {checks.get('count_inc',0)} + LTD {checks.get('count_ltd',0)}",
    )

    # 2. חובה מאוזן
    report.add(
        "חובה מאוזן",
        passed=checks.get('debit_ok', False),
        critical=True,
        message="חובה לא מאוזן" if not checks.get('debit_ok') else "✅",
        expected=f"${checks.get('data_debit',0):,.0f}",
        actual=f"INC ${checks.get('inc_debit',0):,.0f} + LTD ${checks.get('ltd_debit',0):,.0f}",
    )

    # 3. זכות מאוזן
    report.add(
        "זכות מאוזן",
        passed=checks.get('credit_ok', False),
        critical=True,
        message="זכות לא מאוזן" if not checks.get('credit_ok') else "✅",
        expected=f"${checks.get('data_credit',0):,.0f}",
        actual=f"INC ${checks.get('inc_credit',0):,.0f} + LTD ${checks.get('ltd_credit',0):,.0f}",
    )

    return report


# ===== PAY-IT =====

def validate_payit(data: list, pdf_total: float = 0) -> ValidationReport:
    report = ValidationReport()

    if not data:
        report.add("שורות נתונים", False, True, "לא חולצו נתונים מה-PDF")
        return report

    extracted_total = sum(float(d['amount']) for d in data)

    # 1. כמות קופות
    report.add(
        "קופות חולצו",
        passed=len(data) > 0,
        critical=True,
        message=f"{len(data)} קופות",
    )

    # 2. אימות מול PDF
    if pdf_total > 0:
        diff = abs(extracted_total - pdf_total)
        report.add(
            "סה״כ תואם PDF",
            passed=diff < 0.01,
            critical=True,
            message=f"הפרש ₪{diff:,.2f} — נתונים חסרים!" if diff >= 0.01 else "✅ תואם",
            expected=f"₪{pdf_total:,.2f}",
            actual=f"₪{extracted_total:,.2f}",
        )

    return report


# ===== INCOME (RACHEL MOR) =====

def validate_income(invoice_rows: list, receipt_rows: list,
                    unmatched: list, unknown_cols: list) -> ValidationReport:
    report = ValidationReport()

    # 1. יש נתונים
    report.add(
        "שורות חשבוניות",
        passed=len(invoice_rows) > 0,
        critical=True,
        message=f"{len(invoice_rows)} חשבוניות" if invoice_rows else "לא נמצאו חשבוניות",
    )

    # 2. לקוחות לא מזוהים
    report.add(
        "לקוחות מזוהים",
        passed=len(unmatched) == 0,
        critical=True,
        message=f"{len(unmatched)} לקוחות לא נמצאו" if unmatched else "✅ כל הלקוחות מזוהים",
        actual=", ".join(unmatched[:3]) if unmatched else "✅",
    )

    # 3. עמודות לא מוכרות
    report.add(
        "עמודות מוכרות",
        passed=len(unknown_cols) == 0,
        critical=True,
        message=f"{len(unknown_cols)} עמודות לא מוגדרות" if unknown_cols else "✅",
        actual=", ".join(unknown_cols) if unknown_cols else "✅",
    )

    # 4. חשבוניות = קבלות
    report.add(
        "חשבוניות = קבלות",
        passed=len(invoice_rows) > 0 and len(receipt_rows) > 0,
        critical=False,
        message="חסרות קבלות" if not receipt_rows else "✅",
    )

    # 5. בדיקת איזון חובה=זכות לכל פקודה — דיוק מוחלט (cents)
    bad_balance = 0
    bad_examples = []
    for idx, row in enumerate(invoice_rows, start=1):
        if len(row) >= 9:
            try:
                debit_c   = round(float(row[6]) * 100)
                credit1_c = round(float(row[7]) * 100)
                vat_c     = round(float(row[8]) * 100)
                if debit_c != credit1_c + vat_c:
                    bad_balance += 1
                    if len(bad_examples) < 3:
                        diff = (debit_c - credit1_c - vat_c) / 100
                        bad_examples.append(f"שורה {idx}: הפרש ₪{diff:,.2f}")
            except Exception:
                bad_balance += 1
    report.add(
        "איזון חובה=זכות (דיוק מוחלט)",
        passed=bad_balance == 0,
        critical=True,
        message=f"{bad_balance} פקודות לא מאוזנות — חשבשבת ידחה את הקובץ" if bad_balance else "✅ כל הפקודות מאוזנות לאגורה",
        actual=" | ".join(bad_examples) if bad_examples else "✅",
    )

    return report
