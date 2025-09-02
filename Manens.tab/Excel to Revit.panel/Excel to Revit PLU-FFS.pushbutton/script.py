# -*- coding: utf-8 -*-
"""
EXCEL -> REVIT | PLU/FFS
Legge un unico Excel con 5 fogli (come l'export) e aggiorna sugli elementi Revit
i parametri stringa:
  - MAN_ProductCode
  - MAN_BoQ_Units

Regole di match (uguali all'export):
- Tubazioni:            Type Name + Diameter
- Isolante Tubazioni:   Type Name + Insulation Thickness + Pipe Size
- Raccordi Tubi:        Family Name + Type Name + MAN_Fittings_MaxSize (mm)
- Apparecchiature Mec:  Family Name + Type Name + MAN_Type_Code
- Generale:             Family Name + Type Name

Se in Excel il valore e vuoto -> il parametro in Revit viene impostato a "" (svuotato).
"""

__title__ = 'Excel to Revit\nPLU/FFS'
__author__ = 'Valerio Mascia'

import clr, System, re
from System import String, Array, Object
from System.Runtime.InteropServices import Marshal

# Excel Interop
clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop import Excel

# Revit
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInParameter, BuiltInCategory, FamilyInstance, Transaction
)
from Autodesk.Revit.DB.Plumbing import Pipe, PipeInsulation

# Unit conversion (Revit 2022+ / <=2021)
try:
    from Autodesk.Revit.DB import UnitUtils, UnitTypeId  # 2022+
    _HAS_UTID = True
except:
    from Autodesk.Revit.DB import UnitUtils, DisplayUnitType  # <=2021
    _HAS_UTID = False

# Dialog / UI
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import (
    OpenFileDialog, DialogResult, Form, CheckBox, Button, Label, AnchorStyles,
    FormStartPosition, FormBorderStyle
)
from System.Drawing import Point, Size, Font, FontStyle

doc = __revit__.ActiveUIDocument.Document

# -------------------------- UI ------------------------------
class RunPickerForm(Form):
    def __init__(self):
        Form.__init__(self)
        self.Text = "Excel → Revit | PLU/FFS - Seleziona cosa importare"
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.ClientSize = Size(460, 270)

        self.lbl = Label()
        self.lbl.Text = "Scegli le importazioni da eseguire (da Excel → in Revit):"
        self.lbl.Location = Point(16, 16)
        self.lbl.AutoSize = True
        self.lbl.Font = Font(self.Font, FontStyle.Bold)
        self.Controls.Add(self.lbl)

        self.chkPipe = CheckBox()
        self.chkPipe.Text = "Tubazioni (Type + Diameter)"
        self.chkPipe.Location = Point(20, 50)
        self.chkPipe.AutoSize = True
        self.chkPipe.Checked = True
        self.Controls.Add(self.chkPipe)

        self.chkIns = CheckBox()
        self.chkIns.Text = "Isolante Tubazioni (Type + Thickness + Pipe Size)"
        self.chkIns.Location = Point(20, 78)
        self.chkIns.AutoSize = True
        self.chkIns.Checked = True
        self.Controls.Add(self.chkIns)

        self.chkFit = CheckBox()
        self.chkFit.Text = "Raccordi Tubi (Family/Type + MaxSize)"
        self.chkFit.Location = Point(20, 106)
        self.chkFit.AutoSize = True
        self.chkFit.Checked = True
        self.Controls.Add(self.chkFit)

        self.chkMeq = CheckBox()
        self.chkMeq.Text = "Apparecchiature Mec (Family/Type + MAN_Type_Code)"
        self.chkMeq.Location = Point(20, 134)
        self.chkMeq.AutoSize = True
        self.chkMeq.Checked = True
        self.Controls.Add(self.chkMeq)

        self.chkGen = CheckBox()
        self.chkGen.Text = "Generale (PA / PF / Sprinklers) (Family/Type)"
        self.chkGen.Location = Point(20, 162)
        self.chkGen.AutoSize = True
        self.chkGen.Checked = True
        self.Controls.Add(self.chkGen)

        self.btnOk = Button()
        self.btnOk.Text = "OK"
        self.btnOk.Size = Size(100, 28)
        self.btnOk.Location = Point(self.ClientSize.Width - 220, 220)
        self.btnOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btnOk.DialogResult = DialogResult.OK
        self.Controls.Add(self.btnOk)

        self.btnCancel = Button()
        self.btnCancel.Text = "Annulla"
        self.btnCancel.Size = Size(100, 28)
        self.btnCancel.Location = Point(self.ClientSize.Width - 110, 220)
        self.btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btnCancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(self.btnCancel)

        self.AcceptButton = self.btnOk
        self.CancelButton = self.btnCancel

def pick_excel_path_once():
    dlg = OpenFileDialog()
    dlg.Title = "Seleziona il file Excel (PLU/FFS)"
    dlg.Filter = "Excel (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls"
    dlg.Multiselect = False
    return dlg.FileName if dlg.ShowDialog() == DialogResult.OK else None

# --------------------- Utils testo/numero -------------------
def _u(s):
    if s is None: return u""
    try:
        return s if isinstance(s, unicode) else unicode(s)
    except:
        return unicode(str(s))

def _norm_text(s):
    return _u(s).strip()

def _norm_text_strong(s):
    return u" ".join(_norm_text(s).split())

def _strip_phi(s):
    if s is None: return u""
    t = _u(s).strip()
    t = re.sub(u"[ \t]*(?:[ΦφØø⌀])$", u"", t)
    t = re.sub(u"^(?:[ΦφØø⌀])[ \t]*", u"", t)
    return t.strip()

def _number_from_text(s):
    if not s: return u""
    ss = _u(s).replace(",", ".")
    m = re.search(r'(\d+(?:\.\d+)?)', ss)
    if not m: return u""
    num = m.group(1)
    if "." in num:
        num = num.rstrip("0").rstrip(".")
    return num

def _to_float(s, default=0.0):
    try:
        if s is None or s == u"": return float(default)
        return float(_u(s).replace(",", ".").strip())
    except:
        return float(default)

def _feet_to_mm(val_ft):
    try:
        if _HAS_UTID:
            return float(UnitUtils.ConvertFromInternalUnits(float(val_ft), UnitTypeId.Millimeters))
        else:
            return float(UnitUtils.ConvertFromInternalUnits(float(val_ft), DisplayUnitType.DUT_MILLIMETERS))
    except:
        try: return float(val_ft) * 304.8
        except: return 0.0

# --------------------- Lettura da Revit (chiavi) ------------
def _category_name(elem):
    try:
        p = elem.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM)
        if p:
            vs = p.AsValueString()
            if vs: return vs
    except: pass
    try:
        if elem.Category and elem.Category.Name:
            return elem.Category.Name
    except: pass
    return ""

def _type_name_from_instance(elem):
    try:
        p = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
        if p:
            s = p.AsValueString()
            if s: return s
            tid = p.AsElementId()
            if tid and tid.IntegerValue > 0:
                t = doc.GetElement(tid)
                if t:
                    q = t.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
                    if q:
                        qs = q.AsString()
                        if qs: return qs
                    q2 = t.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                    if q2:
                        q2s = q2.AsString()
                        if q2s: return q2s
                    try: return t.Name or ""
                    except: pass
    except: pass
    return ""

def _family_name(elem):
    try:
        if hasattr(elem, "Symbol") and elem.Symbol and elem.Symbol.Family and elem.Symbol.Family.Name:
            return elem.Symbol.Family.Name
    except: pass
    try:
        p = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
        if p:
            tid = p.AsElementId()
            if tid and tid.IntegerValue > 0:
                t = doc.GetElement(tid)
                if t and hasattr(t, "Family") and t.Family and t.Family.Name:
                    return t.Family.Name
    except: pass
    return ""

# Pipe: ricavo chiave Diametro (come export)
def _diameter_key_from_pipe(elem):
    try:
        p = elem.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
        if not p: return u""
        # prendo la rappresentazione testuale (serve per intercettare "DN" ecc.)
        try:
            vs = p.AsValueString()
        except:
            vs = ""
        if not vs:
            try:
                d_ft = p.AsDouble()
                if d_ft is not None:
                    # in mm, poi tolgo zeri finali
                    d_mm = _feet_to_mm(d_ft)
                    vs = (u"%.3f" % d_mm).rstrip("0").rstrip(".")
            except:
                vs = ""
        if not vs: return u""
        # stessa normalizzazione dell'export
        s = _u(vs).strip().replace(",", ".")
        m = re.search(r'(\d+(?:\.\d+)?)', s)
        if not m: return u""
        num = m.group(1)
        if "." in num:
            num = num.rstrip("0").rstrip(".")
        return num
    except:
        return u""

# Pipe Insulation: ricavo coppia (thick, size)
def _insulation_keys(elem):
    thick_key = u""; size_key = u""
    try:
        pth = elem.get_Parameter(BuiltInParameter.RBS_INSULATION_THICKNESS_FOR_PIPE)
        if pth:
            try:
                d_ft = pth.AsDouble()
                if d_ft is not None:
                    d_mm = _feet_to_mm(d_ft)
                    thick_key = (u"%.3f" % d_mm).rstrip("0").rstrip(".")
            except:
                disp = pth.AsValueString() or ""
                thick_key = _number_from_text(disp)
    except: pass
    try:
        psz = elem.get_Parameter(BuiltInParameter.RBS_PIPE_CALCULATED_SIZE)
        if psz:
            raw = psz.AsValueString() or ""
            size_key = _number_from_text(_strip_phi(raw))
    except: pass
    return thick_key or u"0", size_key

# Fittings: chiave MaxSize mm
def _fitting_maxsize_mm(elem):
    try:
        p = elem.LookupParameter("MAN_Fittings_MaxSize")
        if not p: return 0.0
        d_ft = None
        try:
            d_ft = p.AsDouble()
        except:
            s = (p.AsString() or p.AsValueString() or "").replace(",", ".").strip()
            if not s: return 0.0
            try: d_ft = float(s)
            except: return 0.0
        return _feet_to_mm(d_ft)
    except:
        return 0.0

# Mechanical: type code
def _meq_type_code(elem):
    try:
        q = elem.LookupParameter("MAN_Type_Code")
        if not q: return u""
        s = q.AsString()
        if s: return _norm_text_strong(s)
        vs = q.AsValueString()
        if vs: return _norm_text_strong(vs)
    except: pass
    return u""

# Set parametri stringa (vuoto consente di "svuotare")
def _set_str(elem, pname, value):
    p = elem.LookupParameter(pname)
    if not p or p.IsReadOnly: return False
    try:
        p.Set(_u(value or u""))
        return True
    except:
        return False

# ---------------------- Excel helpers -----------------------
def _get_sheet(workbook, name):
    try:
        return workbook.Worksheets.Item[name]
    except:
        return None

def _headers_dict(sheet, header_row):
    last_col = sheet.Cells(header_row, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
    if last_col < 1: last_col = 1
    headers = {}
    for c in range(1, last_col + 1):
        v = sheet.Cells(header_row, c).Value2
        try:
            nm = v.strip() if isinstance(v, String) else None
        except:
            nm = None
        if nm: headers[nm] = c
    return headers

def _read_col(sheet, col, r0, r1, norm=True):
    if r1 < r0: return []
    rng = sheet.Range[sheet.Cells(r0, col), sheet.Cells(r1, col)]
    data = rng.Value2
    out = []
    if isinstance(data, System.Array) and data.Rank == 2:
        n0 = data.GetLength(0)
        for i in range(n0):
            try: val = data.GetValue(i, 0)
            except:
                try: val = data.GetValue(i+1, 1)
                except: val = None
            out.append(_norm_text(val) if norm else val)
        return out
    if isinstance(data, (tuple, list)):
        for row in data:
            try: val = row[0]
            except: val = row
            out.append(_norm_text(val) if norm else val)
        return out
    out.append(_norm_text(data) if norm else data)
    return out

def _detect_region(sheet, key_cols, start_row, empty_stop=20):
    r0 = start_row
    step = 2000
    curr = r0
    last = r0 - 1
    empty_run = 0
    max_rows = sheet.Rows.Count
    while curr <= max_rows:
        r1_try = min(max_rows, curr + step - 1)
        cols_vals = [ _read_col(sheet, c, curr, r1_try) for c in key_cols ]
        block_len = max(len(cv) for cv in cols_vals) if cols_vals else 0
        for i in range(block_len):
            any_val = False
            for cv in cols_vals:
                v = cv[i] if i < len(cv) else ""
                if v:
                    any_val = True; break
            if any_val:
                last = curr + i
                empty_run = 0
            else:
                empty_run += 1
                if empty_run >= empty_stop:
                    return (r0, last) if last >= r0 else (r0, r0-1)
        curr = r1_try + 1
    return (r0, last) if last >= r0 else (r0, r0-1)

# ------------------ IMPORT: TUBAZIONI -----------------------
SHEET_NAME_PIPE = "Tubazioni"
HEADER_ROW_PIPE = 3
START_ROW_PIPE = 5

def _build_rules_pipe(sheet):
    H = _headers_dict(sheet, HEADER_ROW_PIPE)
    required = ("Type Name", "Diameter")
    for r in required:
        if r not in H:
            print("[PIPE] Colonna mancante:", r)
            return {}
    # opzionali ma consigliate
    if "MAN_ProductCode" not in H or "MAN_BoQ_Units" not in H:
        print("[PIPE] Mancano MAN_ProductCode / MAN_BoQ_Units: nessun aggiornamento.")
        return {}

    r0, r1 = _detect_region(sheet, [H["Type Name"], H["Diameter"]], START_ROW_PIPE)
    if r1 < r0: return {}

    col_tn = _read_col(sheet, H["Type Name"], r0, r1)
    col_d  = _read_col(sheet, H["Diameter"],  r0, r1)
    col_pc = _read_col(sheet, H["MAN_ProductCode"], r0, r1)
    col_bu = _read_col(sheet, H["MAN_BoQ_Units"],   r0, r1)

    rules = {}
    n = max(len(col_tn), len(col_d), len(col_pc), len(col_bu))
    for i in range(n):
        t  = _norm_text(col_tn[i] if i < len(col_tn) else "")
        d0 = _norm_text(col_d[i]  if i < len(col_d)  else "")
        d  = _number_from_text(d0)
        if not (t or d): continue
        pc = _u(col_pc[i] if i < len(col_pc) else u"")
        bu = _u(col_bu[i] if i < len(col_bu) else u"")
        rules[(t, d)] = (pc, bu)
    return rules

def import_pipe(workbook):
    sh = _get_sheet(workbook, SHEET_NAME_PIPE)
    if not sh:
        print("[PIPE] Foglio non trovato.")
        return
    rules = _build_rules_pipe(sh)
    if not rules:
        print("[PIPE] Nessuna regola da Excel.")
        return

    elems = FilteredElementCollector(doc).OfClass(Pipe).WhereElementIsNotElementType().ToElements()
    updated = 0; miss_p = 0; not_matched = 0
    t = Transaction(doc, "Excel→Revit | PLU/FFS | Pipes")
    t.Start()
    try:
        for e in elems:
            tn = _norm_text(_type_name_from_instance(e))
            dk = _diameter_key_from_pipe(e)
            key = (tn, dk)
            if key in rules:
                pc, bu = rules[key]
                ok1 = False; ok2 = False
                att1 = _u(pc).strip() != u""
                att2 = _u(bu).strip() != u""

                if att1:
                    ok1 = _set_str(e, "MAN_ProductCode", pc)
                if att2:
                    ok2 = _set_str(e, "MAN_BoQ_Units",   bu)

                if ok1 or ok2:
                    updated += 1
                if (att1 and not ok1) or (att2 and not ok2):
                    miss_p += 1
            else:
                not_matched += 1
        t.Commit()
    except:
        t.RollBack()
        raise
    finally:
        try: Marshal.ReleaseComObject(sh)
        except: pass
    print("[PIPE] Aggiornati elementi:", updated, "| Non trovati:", not_matched, "| Parametri mancanti:", miss_p)

# ---------------- IMPORT: ISOLANTE TUBAZIONI ----------------
SHEET_NAME_INS = "Isolante Tubazioni"
HEADER_ROW_INS = 3
START_ROW_INS = 5

def _build_rules_ins(sheet):
    H = _headers_dict(sheet, HEADER_ROW_INS)
    required = ("Type Name", "Insulation Thickness", "Pipe Size")
    for r in required:
        if r not in H:
            print("[INS] Colonna mancante:", r)
            return {}
    if "MAN_ProductCode" not in H or "MAN_BoQ_Units" not in H:
        print("[INS] Mancano MAN_ProductCode / MAN_BoQ_Units: nessun aggiornamento.")
        return {}

    r0, r1 = _detect_region(sheet, [H["Type Name"], H["Insulation Thickness"], H["Pipe Size"]], START_ROW_INS)
    if r1 < r0: return {}

    col_tn = _read_col(sheet, H["Type Name"], r0, r1)
    col_th = _read_col(sheet, H["Insulation Thickness"], r0, r1)
    col_sz = _read_col(sheet, H["Pipe Size"], r0, r1)
    col_pc = _read_col(sheet, H["MAN_ProductCode"], r0, r1)
    col_bu = _read_col(sheet, H["MAN_BoQ_Units"],   r0, r1)

    rules = {}
    n = max(len(col_tn), len(col_th), len(col_sz), len(col_pc), len(col_bu))
    for i in range(n):
        t  = _norm_text(col_tn[i] if i < len(col_tn) else "")
        th = _number_from_text(col_th[i] if i < len(col_th) else "")
        sz = _number_from_text(_strip_phi(col_sz[i] if i < len(col_sz) else ""))
        if not (t or th or sz): continue
        if not th: th = u"0"
        pc = _u(col_pc[i] if i < len(col_pc) else u"")
        bu = _u(col_bu[i] if i < len(col_bu) else u"")
        rules[(t, th, sz)] = (pc, bu)
    return rules

def _is_pipe_hosted_ins(ins_elem):
    try:
        hid = ins_elem.HostElementId
        if hid and hid.IntegerValue > 0:
            host = doc.GetElement(hid)
            return isinstance(host, Pipe)
    except: pass
    return False

def import_insulation(workbook):
    sh = _get_sheet(workbook, SHEET_NAME_INS)
    if not sh:
        print("[INS] Foglio non trovato.")
        return
    rules = _build_rules_ins(sh)
    if not rules:
        print("[INS] Nessuna regola da Excel.")
        return

    elems = FilteredElementCollector(doc).OfClass(PipeInsulation).WhereElementIsNotElementType().ToElements()
    updated = 0; miss_p = 0; not_matched = 0
    t = Transaction(doc, "Excel→Revit | PLU/FFS | Pipe Insulations")
    t.Start()
    try:
        for e in elems:
            tn = _norm_text(_type_name_from_instance(e))
            th, sz = _insulation_keys(e)
            key = (tn, th or u"0", sz)
            if key in rules:
                pc, bu = rules[key]
                ok1 = False; ok2 = False
                att1 = _u(pc).strip() != u""
                att2 = _u(bu).strip() != u""

                if att1:
                    ok1 = _set_str(e, "MAN_ProductCode", pc)
                if att2:
                    ok2 = _set_str(e, "MAN_BoQ_Units",   bu)

                if ok1 or ok2:
                    updated += 1
                if (att1 and not ok1) or (att2 and not ok2):
                    miss_p += 1
            else:
                not_matched += 1
        t.Commit()
    except:
        t.RollBack()
        raise
    finally:
        try: Marshal.ReleaseComObject(sh)
        except: pass
    print("[INS] Aggiornati elementi:", updated, "| Non trovati:", not_matched, "| Parametri mancanti:", miss_p)

# ------------------ IMPORT: RACCORDI TUBI -------------------
SHEET_NAME_FIT = "Raccordi Tubi"
HEADER_ROW_FIT = 3
START_ROW_FIT = 5

def _build_rules_fit(sheet):
    H = _headers_dict(sheet, HEADER_ROW_FIT)
    required = ("Family Name", "Type Name", "MAN_Fittings_MaxSize")
    for r in required:
        if r not in H:
            print("[FIT] Colonna mancante:", r)
            return {}
    if "MAN_ProductCode" not in H or "MAN_BoQ_Units" not in H:
        print("[FIT] Mancano MAN_ProductCode / MAN_BoQ_Units: nessun aggiornamento.")
        return {}

    r0, r1 = _detect_region(sheet, [H["Family Name"], H["Type Name"], H["MAN_Fittings_MaxSize"]], START_ROW_FIT)
    if r1 < r0: return {}

    col_f  = _read_col(sheet, H["Family Name"], r0, r1)
    col_t  = _read_col(sheet, H["Type Name"],   r0, r1)
    col_m  = _read_col(sheet, H["MAN_Fittings_MaxSize"], r0, r1)
    col_pc = _read_col(sheet, H["MAN_ProductCode"], r0, r1)
    col_bu = _read_col(sheet, H["MAN_BoQ_Units"],   r0, r1)

    rules = {}
    n = max(len(col_f), len(col_t), len(col_m), len(col_pc), len(col_bu))
    for i in range(n):
        f = _norm_text_strong(col_f[i] if i < len(col_f) else "")
        t = _norm_text_strong(col_t[i] if i < len(col_t) else "")
        try:
            m = round(_to_float(col_m[i] if i < len(col_m) else 0.0), 6)
        except:
            m = 0.0
        if not (f or t or m): continue
        pc = _u(col_pc[i] if i < len(col_pc) else u"")
        bu = _u(col_bu[i] if i < len(col_bu) else u"")
        rules[(f, t, m)] = (pc, bu)
    return rules

def import_fittings(workbook):
    sh = _get_sheet(workbook, SHEET_NAME_FIT)
    if not sh:
        print("[FIT] Foglio non trovato.")
        return
    rules = _build_rules_fit(sh)
    if not rules:
        print("[FIT] Nessuna regola da Excel.")
        return

    elems = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeFitting)\
        .WhereElementIsNotElementType().ToElements()
    updated = 0; miss_p = 0; not_matched = 0
    t = Transaction(doc, "Excel→Revit | PLU/FFS | Pipe Fittings")
    t.Start()
    try:
        for e in elems:
            if not isinstance(e, FamilyInstance):
                continue
            fam = _norm_text_strong(_family_name(e))
            typ = _norm_text_strong(_type_name_from_instance(e))
            msz = round(float(_fitting_maxsize_mm(e)), 6)
            key = (fam, typ, msz)
            if key in rules:
                pc, bu = rules[key]
                ok1 = False; ok2 = False
                att1 = _u(pc).strip() != u""
                att2 = _u(bu).strip() != u""

                if att1:
                    ok1 = _set_str(e, "MAN_ProductCode", pc)
                if att2:
                    ok2 = _set_str(e, "MAN_BoQ_Units",   bu)

                if ok1 or ok2:
                    updated += 1
                if (att1 and not ok1) or (att2 and not ok2):
                    miss_p += 1
            else:
                not_matched += 1
        t.Commit()
    except:
        t.RollBack()
        raise
    finally:
        try: Marshal.ReleaseComObject(sh)
        except: pass
    print("[FIT] Aggiornati elementi:", updated, "| Non trovati:", not_matched, "| Parametri mancanti:", miss_p)

# -------------- IMPORT: APPARECCHIATURE MEC -----------------
MEQ_SHEET_NAME = "Apparecchiature Mec"
MEQ_HEADER_ROW = 3
MEQ_START_ROW = 5

def _build_rules_meq(sheet):
    H = _headers_dict(sheet, MEQ_HEADER_ROW)
    required = ("Family Name", "Type Name", "MAN_Type_Code")
    for r in required:
        if r not in H:
            print("[MEQ] Colonna mancante:", r)
            return {}
    if "MAN_ProductCode" not in H or "MAN_BoQ_Units" not in H:
        print("[MEQ] Mancano MAN_ProductCode / MAN_BoQ_Units: nessun aggiornamento.")
        return {}

    r0, r1 = _detect_region(sheet, [H["Family Name"], H["Type Name"], H["MAN_Type_Code"]], MEQ_START_ROW)
    if r1 < r0: return {}

    col_f  = _read_col(sheet, H["Family Name"], r0, r1)
    col_t  = _read_col(sheet, H["Type Name"],   r0, r1)
    col_c  = _read_col(sheet, H["MAN_Type_Code"], r0, r1)
    col_pc = _read_col(sheet, H["MAN_ProductCode"], r0, r1)
    col_bu = _read_col(sheet, H["MAN_BoQ_Units"],   r0, r1)

    rules = {}
    n = max(len(col_f), len(col_t), len(col_c), len(col_pc), len(col_bu))
    for i in range(n):
        f = _norm_text_strong(col_f[i] if i < len(col_f) else "")
        t = _norm_text_strong(col_t[i] if i < len(col_t) else "")
        c = _norm_text_strong(col_c[i] if i < len(col_c) else "")
        if not (f or t or c): continue
        pc = _u(col_pc[i] if i < len(col_pc) else u"")
        bu = _u(col_bu[i] if i < len(col_bu) else u"")
        rules[(f, t, c)] = (pc, bu)
    return rules

def import_meq(workbook):
    sh = _get_sheet(workbook, MEQ_SHEET_NAME)
    if not sh:
        print("[MEQ] Foglio non trovato.")
        return
    rules = _build_rules_meq(sh)
    if not rules:
        print("[MEQ] Nessuna regola da Excel.")
        return

    elems = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment)\
        .WhereElementIsNotElementType().ToElements()
    updated = 0; miss_p = 0; not_matched = 0
    t = Transaction(doc, "Excel→Revit | PLU/FFS | Mechanical Equipment")
    t.Start()
    try:
        for e in elems:
            fam = _norm_text_strong(_family_name(e))
            typ = _norm_text_strong(_type_name_from_instance(e))
            code = _meq_type_code(e)
            key = (fam, typ, code)
            if key in rules:
                pc, bu = rules[key]
                ok1 = False; ok2 = False
                att1 = _u(pc).strip() != u""
                att2 = _u(bu).strip() != u""

                if att1:
                    ok1 = _set_str(e, "MAN_ProductCode", pc)
                if att2:
                    ok2 = _set_str(e, "MAN_BoQ_Units",   bu)

                if ok1 or ok2:
                    updated += 1
                if (att1 and not ok1) or (att2 and not ok2):
                    miss_p += 1
            else:
                not_matched += 1
        t.Commit()
    except:
        t.RollBack()
        raise
    finally:
        try: Marshal.ReleaseComObject(sh)
        except: pass
    print("[MEQ] Aggiornati elementi:", updated, "| Non trovati:", not_matched, "| Parametri mancanti:", miss_p)

# -------------------- IMPORT: GENERALE ----------------------
GEN_SHEET_NAME = "Generale"
GEN_HEADER_ROW = 3
GEN_START_ROW = 5

def _build_rules_gen(sheet):
    H = _headers_dict(sheet, GEN_HEADER_ROW)
    required = ("Family Name", "Type Name")
    for r in required:
        if r not in H:
            print("[GEN] Colonna mancante:", r)
            return {}
    if "MAN_ProductCode" not in H or "MAN_BoQ_Units" not in H:
        print("[GEN] Mancano MAN_ProductCode / MAN_BoQ_Units: nessun aggiornamento.")
        return {}

    r0, r1 = _detect_region(sheet, [H["Family Name"], H["Type Name"]], GEN_START_ROW)
    if r1 < r0: return {}

    col_f  = _read_col(sheet, H["Family Name"], r0, r1)
    col_t  = _read_col(sheet, H["Type Name"],   r0, r1)
    col_pc = _read_col(sheet, H["MAN_ProductCode"], r0, r1)
    col_bu = _read_col(sheet, H["MAN_BoQ_Units"],   r0, r1)

    rules = {}
    n = max(len(col_f), len(col_t), len(col_pc), len(col_bu))
    for i in range(n):
        f = _norm_text_strong(col_f[i] if i < len(col_f) else "")
        t = _norm_text_strong(col_t[i] if i < len(col_t) else "")
        if not (f or t): continue
        pc = _u(col_pc[i] if i < len(col_pc) else u"")
        bu = _u(col_bu[i] if i < len(col_bu) else u"")
        rules[(f, t)] = (pc, bu)
    return rules

def import_general(workbook):
    sh = _get_sheet(workbook, GEN_SHEET_NAME)
    if not sh:
        print("[GEN] Foglio non trovato.")
        return
    rules = _build_rules_gen(sh)
    if not rules:
        print("[GEN] Nessuna regola da Excel.")
        return

    cats = (BuiltInCategory.OST_PipeAccessory,
            BuiltInCategory.OST_PlumbingFixtures,
            BuiltInCategory.OST_Sprinklers)
    elems = []
    for bic in cats:
        elems.extend(list(FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType().ToElements()))

    updated = 0; miss_p = 0; not_matched = 0
    t = Transaction(doc, "Excel→Revit | PLU/FFS | Generale (PA/PF/Sprinklers)")
    t.Start()
    try:
        for e in elems:
            if not isinstance(e, FamilyInstance):
                continue
            fam = _norm_text_strong(_family_name(e))
            typ = _norm_text_strong(_type_name_from_instance(e))
            key = (fam, typ)
            if key in rules:
                pc, bu = rules[key]
                ok1 = False; ok2 = False
                att1 = _u(pc).strip() != u""
                att2 = _u(bu).strip() != u""

                if att1:
                    ok1 = _set_str(e, "MAN_ProductCode", pc)
                if att2:
                    ok2 = _set_str(e, "MAN_BoQ_Units",   bu)

                if ok1 or ok2:
                    updated += 1
                if (att1 and not ok1) or (att2 and not ok2):
                    miss_p += 1
            else:
                not_matched += 1
        t.Commit()
    except:
        t.RollBack()
        raise
    finally:
        try: Marshal.ReleaseComObject(sh)
        except: pass
    print("[GEN] Aggiornati elementi:", updated, "| Non trovati:", not_matched, "| Parametri mancanti:", miss_p)

# ----------------------------- MAIN -------------------------
def main():
    form = RunPickerForm()
    if form.ShowDialog() != DialogResult.OK:
        return

    run_pipe = form.chkPipe.Checked
    run_ins  = form.chkIns.Checked
    run_fit  = form.chkFit.Checked
    run_meq  = form.chkMeq.Checked
    run_gen  = form.chkGen.Checked

    if not (run_pipe or run_ins or run_fit or run_meq or run_gen):
        print("Nessuna opzione selezionata. Operazione annullata.")
        return

    excel_path = pick_excel_path_once()
    if not excel_path:
        return

    excel = None; workbook = None
    try:
        excel = Excel.ApplicationClass()
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(excel_path)

        if run_pipe:
            import_pipe(workbook)
        if run_ins:
            import_insulation(workbook)
        if run_fit:
            import_fittings(workbook)
        if run_meq:
            import_meq(workbook)
        if run_gen:
            import_general(workbook)

        # chiudo Excel
        workbook.Close(False)
        excel.Quit()

    finally:
        try:
            if workbook: Marshal.ReleaseComObject(workbook)
        except: pass
        try:
            if excel: Marshal.ReleaseComObject(excel)
        except: pass

if __name__ == "__main__":
    main()