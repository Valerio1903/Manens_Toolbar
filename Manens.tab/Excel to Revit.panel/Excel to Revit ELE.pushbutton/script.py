# -*- coding: utf-8 -*-
"""
Excel -> Revit | COMBO
Importa MAN_ProductCode e MAN_BoQ_Units (istanza) da Excel agli elementi Revit
per i seguenti sheet / blocchi (stesse chiavi e filtri dell'export):
- Passerelle (Cable Trays)                  [Type Name + Size]
- Separatore passerelle (con MAN_Dividers)  [Type Name + Height mm]
- Cavidotti (Conduits)                      [Type Name + Outside Diameter mm] (esclude ThermoCable/AirSampling)
- Quadri elettrici (EEQ)                    [Family + Type + Level + Panel] (solo famiglie MAN_EEQ_PNB_SwitchBoard)
- Generale (EF/LF/LD + fittings vari)       [Family + Type] (esclude EEQ MAN_EEQ_PNB_*/MAN_SEQ e ConduitFitting Thermo/Air)
- Tubazioni (Pipe)                          [Type + Diameter]
- Raccordi Tubi (Pipe Fittings)             [Family + Type + MAN_Fittings_MaxSize mm]
- Canali Rigidi (Ducts)                     [Type + Width/Height - Diameter mm (max dim o diametro)]
- Fitting canali (Duct Fittings)            [Family + Type + MAN_Fittings_MaxSize mm]
"""

__title__  = 'Excel to Revit\nELE'
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
    FilteredElementCollector, BuiltInParameter, BuiltInCategory, FamilyInstance,
    Transaction, StorageType
)

# MEP classes
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.DB.Mechanical import Duct

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

# ==================== Util testo comuni ====================
def U(s):
    if s is None: return u""
    try: return s if isinstance(s, unicode) else unicode(s)
    except: return unicode(str(s))
def norm_text(s):  return U(s).strip()
def norm_strong(s): return u" ".join(norm_text(s).split())

def to_float_dot(s):
    if s is None: return None
    ss = U(s).replace(",", ".").strip()
    try: return float(ss)
    except: return None

# =============== Param setter robusto (istanza) ===============
PARAM_NAMES = ("MAN_ProductCode", "MAN_BoQ_Units")

def set_param_generic(param, text_val):
    """Imposta un Parameter da testo, rispettando StorageType (fallback su 0/empty)."""
    if not param: return False, "missing"
    if param.IsReadOnly: return False, "read-only"
    st = param.StorageType
    s  = U(text_val).strip()
    try:
        if st == StorageType.String:
            param.Set(s)
            return True, None
        elif st == StorageType.Integer:
            if s == u"": param.Set(0);  return True, None
            f = to_float_dot(s);  iv = int(f) if f is not None else 0
            param.Set(iv);        return True, None
        elif st == StorageType.Double:
            if s == u"": param.Set(0.0); return True, None
            f = to_float_dot(s);  fv = f if f is not None else 0.0
            param.Set(fv);        return True, None
        else:
            return False, "unsupported"
    except Exception as ex:
        return False, "exc: {}".format(ex)

def apply_two_params(elem, prod_code, boq_units, stats):
    ok_any = False
    for name, val in (("MAN_ProductCode", prod_code), ("MAN_BoQ_Units", boq_units)):
        p = None
        try: p = elem.LookupParameter(name)
        except: p = None
        if not p:
            stats["missing_param"][name] = stats["missing_param"].get(name, 0) + 1
            continue
        ok, err = set_param_generic(p, val)
        if ok:
            stats["set_count"][name] = stats["set_count"].get(name, 0) + 1
            ok_any = True
        else:
            stats["errors"][name] = stats["errors"].get(name, 0) + 1
    return ok_any

# ==================== Excel helpers (generici) ====================
def xl_headers_map(sheet, header_row):
    last_col = sheet.Cells(header_row, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
    if last_col < 1: last_col = 1
    headers = {}
    for c in range(1, last_col + 1):
        v = sheet.Cells(header_row, c).Value2
        try: nm = v.strip() if isinstance(v, String) else None
        except: nm = None
        if nm: headers[nm] = c
    return headers

def xl_read_column_block(sheet, col, r0, r1):
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
            out.append(U(val).strip())
        return out
    if isinstance(data, (tuple, list)):
        for row in data:
            try: val = row[0]
            except: val = row
            out.append(U(val).strip())
        return out
    out.append(U(data).strip())
    return out

def detect_data_region_by_cols(sheet, header_row, min_row, key_col_ids, empty_run_stop=20):
    if not key_col_ids: return (min_row, min_row-1)
    r0 = min_row
    step = 2000
    curr_start = r0
    last_data_row = r0 - 1
    empty_run = 0
    max_rows = sheet.Rows.Count
    while curr_start <= max_rows:
        r1_try = min(max_rows, curr_start + step - 1)
        cols_vals = [xl_read_column_block(sheet, c, curr_start, r1_try) for c in key_col_ids]
        block_len = max((len(v) for v in cols_vals) or [0])
        for i in range(block_len):
            present = False
            for arr in cols_vals:
                if i < len(arr) and arr[i]:
                    present = True; break
            if present:
                last_data_row = curr_start + i
                empty_run = 0
            else:
                empty_run += 1
                if empty_run >= empty_run_stop:
                    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)
        curr_start = r1_try + 1
    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)

# ==================== UI Checkboxes ====================
class RunPickerForm(Form):
    def __init__(self):
        Form.__init__(self)
        self.Text = "Excel → Revit | Seleziona cosa importare"
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.ClientSize = Size(430, 360)

        lbl = Label()
        lbl.Text = "Scegli gli import da eseguire:"
        lbl.Location = Point(16, 16)
        lbl.AutoSize = True
        lbl.Font = Font(self.Font, FontStyle.Bold)
        self.Controls.Add(lbl)

        # contatore verticale compatibile con IronPython
        self._y = 48
        self._dy = 26

        def addchk(text, checked=True):
            c = CheckBox()
            c.Text = text
            c.Location = Point(20, self._y)
            c.AutoSize = True
            c.Checked = checked
            self.Controls.Add(c)
            self._y += self._dy
            return c

        self.chkPAS  = addchk("Passerelle (Cable Trays)")
        self.chkSEP  = addchk("Separatore passerelle (Dividers>0)")
        self.chkCOND = addchk("Cavidotti (Conduits)")
        self.chkEEQ  = addchk("Quadri elettrici (EEQ)")
        self.chkGEN  = addchk("Generale (EF/LF/LD/Fixtures)")
        self.chkPIPE = addchk("Tubazioni (Pipe)")
        self.chkPFIT = addchk("Raccordi Tubi (Pipe Fittings)")
        self.chkDUCT = addchk("Canali Rigidi (Ducts)")
        self.chkDFIT = addchk("Fitting canali (Duct Fittings)")

        self.btnOk = Button(); self.btnOk.Text = "OK"; self.btnOk.Size = Size(100, 28)
        self.btnOk.Location = Point(self.ClientSize.Width - 220, 318)
        self.btnOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btnOk.DialogResult = DialogResult.OK
        self.Controls.Add(self.btnOk)

        self.btnCancel = Button(); self.btnCancel.Text = "Annulla"; self.btnCancel.Size = Size(100, 28)
        self.btnCancel.Location = Point(self.ClientSize.Width - 110, 318)
        self.btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btnCancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(self.btnCancel)

        self.AcceptButton = self.btnOk
        self.CancelButton = self.btnCancel

def pick_excel_path_once():
    dlg = OpenFileDialog()
    dlg.Title  = "Seleziona il file Excel (stesse sheet dell'export)"
    dlg.Filter = "Excel (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls"
    dlg.Multiselect = False
    return dlg.FileName if dlg.ShowDialog() == DialogResult.OK else None

# ==================== Utils per chiavi/letture specifiche ====================
# --- (riuso logiche export per parsing dimensioni/diametri) ---

# PASSERELLE
def PAS_type_name(elem):
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

def PAS_instance_size_raw(elem):
    try:
        p = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE)
        if not p: return ""
        s = p.AsString()
        if s: return norm_text(s)
        vs = p.AsValueString()
        if vs: return norm_text(vs)
        return ""
    except: return ""

def PAS_size_key_and_display(raw):
    if not raw: return "", ""
    txt = U(raw)
    txt = re.sub(u"[ΦφØø⌀ϕ]", u"", txt)
    txt = re.sub(u"[×X]", u"x", txt)
    nums = re.findall(r"(\d+(?:[.,]\d+)?)", txt)
    if len(nums) >= 2:
        a = nums[0].replace(",", "."); b = nums[1].replace(",", ".")
        def _trim(n): return n.rstrip("0").rstrip(".") if "." in n else n
        a = _trim(a); b = _trim(b)
        key = u"{}x{}".format(a, b)
        return key, key
    t = norm_text(re.sub(u"[ ]*mm", u"", txt, flags=re.IGNORECASE)).replace(" ", "")
    return t, t

# SEP (dividers + height)
def SEP_dividers_ok(elem):
    def _val(pp):
        if not pp: return 0
        try: return int(pp.AsInteger())
        except:
            s = (pp.AsString() or pp.AsValueString() or "").replace(",", ".").strip()
            if not s: return 0
            try: return int(float(s))
            except: return 0
    try:
        p = elem.LookupParameter("MAN_Dividers")
        if _val(p) > 0: return True
    except: pass
    try:
        tp = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
        tid = tp.AsElementId() if tp else None
        if tid and tid.IntegerValue > 0:
            t = doc.GetElement(tid)
            if t:
                p2 = t.LookupParameter("MAN_Dividers")
                if _val(p2) > 0: return True
    except: pass
    return False

def _feet_to_mm(val_ft):
    try:
        if _HAS_UTID:
            return float(UnitUtils.ConvertFromInternalUnits(float(val_ft), UnitTypeId.Millimeters))
        else:
            return float(UnitUtils.ConvertFromInternalUnits(float(val_ft), DisplayUnitType.DUT_MILLIMETERS))
    except:
        try: return float(val_ft) * 304.8
        except: return 0.0

def SEP_height_key(elem, prec=6):
    try:
        p = elem.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
    except: p = None
    if p:
        try:
            ft = p.AsDouble()
            if ft and ft > 0:
                mm = _feet_to_mm(ft)
                return round(mm, prec)
        except: pass
        try:
            s = (p.AsString() or p.AsValueString() or "").strip()
            if s:
                f = to_float_dot(s)
                if f and f > 0: return round(float(f), prec)
        except: pass
    return 0.0

# CONDUITS
def COND_diam_key(elem, prec=6):
    p = None
    try: p = elem.get_Parameter(BuiltInParameter.RBS_CONDUIT_OUTER_DIAM_PARAM)
    except: p = None
    if p:
        try:
            ft = p.AsDouble()
            if ft and ft > 0:
                mm = _feet_to_mm(ft)
                return round(mm, prec)
        except: pass
        try:
            s = (p.AsString() or p.AsValueString() or "").strip()
            if s:
                f = to_float_dot(s)
                if f and f > 0: return round(float(f), prec)
        except: pass
    return 0.0

# PIPE
def pipe_diameter_key_from_elem(elem):
    try:
        p = elem.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
        if not p: return ""
        s = (p.AsString() or p.AsValueString() or "").strip()
        if not s: return ""
        s2 = s.replace(",", "."); m = re.search(r'(\d+(?:\.\d+)?)', s2)
        if not m: return ""
        num = m.group(1)
        if "." in num: num = num.rstrip("0").rstrip(".")
        return num
    except: return ""

# PFIT/DFIT size param (istanza) in feet -> mm key
def fittings_max_mm_key(elem, prec=6):
    try:
        p = elem.LookupParameter("MAN_Fittings_MaxSize")
        if not p: return 0.0
        d_ft = None
        try: d_ft = p.AsDouble()
        except:
            s = (p.AsString() or p.AsValueString() or "").replace(",", ".").strip()
            if not s: return 0.0
            try: d_ft = float(s)
            except: return 0.0
        return round(_feet_to_mm(d_ft), prec)
    except: return 0.0

# DUCT size key (diameter or max of W/H) in mm
def duct_size_mm_key(elem, prec=6):
    d = None; w = None; h = None
    try:
        p = elem.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
        if p: d = p.AsDouble()
    except: pass
    if d and d > 0:
        return round(_feet_to_mm(d), prec)
    try:
        pw = elem.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
        if pw: w = pw.AsDouble()
    except: pass
    try:
        ph = elem.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
        if ph: h = ph.AsDouble()
    except: pass
    mx = 0.0
    try:
        if w: mx = max(mx, w)
        if h: mx = max(mx, h)
    except: pass
    return round(_feet_to_mm(mx), prec)

# EEQ helpers
def EEQ_family_name(elem):
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

def EEQ_type_name(elem):
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

def EEQ_panel_name(elem):
    try:
        p = elem.LookupParameter("Panel Name")
        if p:
            s = p.AsString()
            if s: return norm_text(s)
            vs = p.AsValueString()
            if vs: return norm_text(vs)
    except: pass
    return ""

def EEQ_level_name(elem):
    try:
        p = elem.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
        if p:
            try:
                lid = p.AsElementId()
                if lid and lid.IntegerValue > 0:
                    lv = doc.GetElement(lid)
                    if lv and hasattr(lv, "Name") and lv.Name:
                        return norm_text(lv.Name)
            except: pass
            vs = p.AsValueString()
            if vs: return norm_text(vs)
    except: pass
    return ""

# ==================== Lettura sheet -> dict chiave->(PC,Units) ====================
def build_row_map(sheet, header_row, min_row, key_cols_names, key_builder, numeric_cols=None):
    """
    Ritorna: { key_tuple : (prodcode, units) }
    key_cols_names: elenco intestazioni coinvolte per r1 detection
    key_builder: funzione (headers, row_index) -> key_tuple or None
    numeric_cols: opzionali per migliorare detection (es. Height/Diameter)
    """
    headers = xl_headers_map(sheet, header_row)
    # colonne per detection
    cols_for_detection = []
    for nm in key_cols_names or []:
        c = headers.get(nm)
        if c: cols_for_detection.append(c)
    if numeric_cols:
        for nm in numeric_cols:
            c = headers.get(nm)
            if c: cols_for_detection.append(c)
    if not cols_for_detection:
        return {}, (min_row, min_row-1), headers, "manca colonne chiave"

    r0, r1 = detect_data_region_by_cols(sheet, header_row, min_row, cols_for_detection, empty_run_stop=20)
    if r1 < r0:
        return {}, (r0, r1), headers, "sheet vuoto"

    pc_col = headers.get("MAN_ProductCode")
    uq_col = headers.get("MAN_BoQ_Units")
    if not (pc_col or uq_col):
        return {}, (r0, r1), headers, "mancano MAN_ProductCode/MAN_BoQ_Units"

    result = {}
    for r in range(r0, r1+1):
        key = key_builder(headers, r)
        if not key: continue
        pc = sheet.Cells(r, pc_col).Value2 if pc_col else None
        uq = sheet.Cells(r, uq_col).Value2 if uq_col else None
        if (pc is None or U(pc).strip()==u"") and (uq is None or U(uq).strip()==u""):
            continue  # niente da impostare
        result[key] = (U(pc).strip() if pc is not None else u"", U(uq).strip() if uq is not None else u"")
    return result, (r0, r1), headers, None

# ==================== Blocchi: indice elementi + import ====================
def import_passerelle(sheet):
    # indice elementi per (TypeName_strong, SizeKey)
    elems = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CableTray).WhereElementIsNotElementType().ToElements()
    idx = {}
    for e in elems:
        t = norm_strong(PAS_type_name(e) or "")
        raw = PAS_instance_size_raw(e) or ""
        skey, _ = PAS_size_key_and_display(raw)
        if not (t or skey): continue
        idx.setdefault((t, skey), []).append(e)

    def key_builder(headers, r):
        tcol = headers.get("Type Name"); scol = headers.get("Size")
        if not (tcol and scol): return None
        t = norm_strong(sheet.Cells(r, tcol).Value2 or u"")
        skey, _ = PAS_size_key_and_display(sheet.Cells(r, scol).Value2 or u"")
        if not (t or skey): return None
        return (t, skey)

    rows, region, headers, err = build_row_map(sheet, 3, 5, ["Type Name", "Size"], key_builder)
    stats = {"matched_keys":0, "updated_elems":0, "missing_param":{}, "set_count":{}, "errors":{}}
    if err: print("[PAS] Skip:", err); return
    t = Transaction(doc, "Excel→Revit | Passerelle"); t.Start()
    try:
        for k,(pc,uq) in rows.items():
            lst = idx.get(k, [])
            if not lst: continue
            stats["matched_keys"] += 1
            for e in lst:
                if apply_two_params(e, pc, uq, stats): stats["updated_elems"] += 1
    finally:
        t.Commit()
    print("[PAS] Chiavi corrisposte:", stats["matched_keys"], "| Istanze aggiornate:", stats["updated_elems"])

def import_sep(sheet):
    elems = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CableTray).WhereElementIsNotElementType().ToElements()
    idx = {}
    for e in elems:
        if not SEP_dividers_ok(e): continue
        t = norm_strong(PAS_type_name(e) or "")
        hk = SEP_height_key(e)
        if hk <= 0: continue
        idx.setdefault((t, hk), []).append(e)

    def key_builder(headers, r):
        tcol = headers.get("Type Name"); hcol = headers.get("Height")
        if not (tcol and hcol): return None
        t = norm_strong(sheet.Cells(r, tcol).Value2 or u"")
        h = to_float_dot(sheet.Cells(r, hcol).Value2); hk = round(h, 6) if h is not None else 0.0
        if not (t or hk): return None
        return (t, hk)

    rows, region, headers, err = build_row_map(sheet, 3, 5, ["Type Name","Height"], key_builder, numeric_cols=["Height"])
    stats = {"matched_keys":0, "updated_elems":0, "missing_param":{}, "set_count":{}, "errors":{}}
    if err: print("[SEP] Skip:", err); return
    t = Transaction(doc, "Excel→Revit | Sep Passerelle"); t.Start()
    try:
        for k,(pc,uq) in rows.items():
            lst = idx.get(k, [])
            if not lst: continue
            stats["matched_keys"] += 1
            for e in lst:
                if apply_two_params(e, pc, uq, stats): stats["updated_elems"] += 1
    finally:
        t.Commit()
    print("[SEP] Chiavi corrisposte:", stats["matched_keys"], "| Istanze aggiornate:", stats["updated_elems"])

def import_conduits(sheet):
    elems = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Conduit).WhereElementIsNotElementType().ToElements()
    idx = {}
    for e in elems:
        t = norm_strong(PAS_type_name(e) or "")
        if not t: continue
        tnlc = t.lower()
        if ("thermocable" in tnlc) or ("airsampling" in tnlc): continue
        dk = COND_diam_key(e)
        if dk <= 0: continue
        idx.setdefault((t, dk), []).append(e)

    def key_builder(headers, r):
        tcol = headers.get("Type Name"); dcol = headers.get("Outside Diameter")
        if not (tcol and dcol): return None
        t = norm_strong(sheet.Cells(r, tcol).Value2 or u"")
        d = to_float_dot(sheet.Cells(r, dcol).Value2); dk = round(d, 6) if d is not None else 0.0
        if not (t or dk): return None
        return (t, dk)

    rows, region, headers, err = build_row_map(sheet, 3, 5, ["Type Name","Outside Diameter"], key_builder, numeric_cols=["Outside Diameter"])
    stats = {"matched_keys":0, "updated_elems":0, "missing_param":{}, "set_count":{}, "errors":{}}
    if err: print("[COND] Skip:", err); return
    t = Transaction(doc, "Excel→Revit | Conduits"); t.Start()
    try:
        for k,(pc,uq) in rows.items():
            for e in idx.get(k, []):
                if apply_two_params(e, pc, uq, stats): stats["updated_elems"] += 1
            if k in idx: stats["matched_keys"] += 1
    finally:
        t.Commit()
    print("[COND] Chiavi corrisposte:", stats["matched_keys"], "| Istanze aggiornate:", stats["updated_elems"])

def import_eeq(sheet):
    elems = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()
    idx = {}
    for e in elems:
        fam = norm_strong(EEQ_family_name(e) or "")
        if not fam or not fam.startswith("MAN_EEQ_PNB_SwitchBoard"): continue
        typ = norm_strong(EEQ_type_name(e) or "")
        lvl = norm_strong(EEQ_level_name(e) or "")
        pnl = norm_strong(EEQ_panel_name(e) or "")
        idx.setdefault((fam, typ, lvl, pnl), []).append(e)

    def key_builder(headers, r):
        fcol = headers.get("Family Name"); tcol = headers.get("Type Name")
        lcol = headers.get("Level"); pcol = headers.get("Panel Name")
        if not (fcol and tcol and lcol and pcol): return None
        fam = norm_strong(sheet.Cells(r, fcol).Value2 or u"")
        if not fam.startswith("MAN_EEQ_PNB_SwitchBoard"): return None
        typ = norm_strong(sheet.Cells(r, tcol).Value2 or u"")
        lvl = norm_strong(sheet.Cells(r, lcol).Value2 or u"")
        pnl = norm_strong(sheet.Cells(r, pcol).Value2 or u"")
        if not (fam or typ or lvl or pnl): return None
        return (fam, typ, lvl, pnl)

    rows, region, headers, err = build_row_map(sheet, 3, 5, ["Family Name","Type Name","Level","Panel Name"], key_builder)
    stats = {"matched_keys":0, "updated_elems":0, "missing_param":{}, "set_count":{}, "errors":{}}
    if err: print("[EEQ] Skip:", err); return
    t = Transaction(doc, "Excel→Revit | EEQ"); t.Start()
    try:
        for k,(pc,uq) in rows.items():
            for e in idx.get(k, []):
                if apply_two_params(e, pc, uq, stats): stats["updated_elems"] += 1
            if k in idx: stats["matched_keys"] += 1
    finally:
        t.Commit()
    print("[EEQ] Chiavi corrisposte:", stats["matched_keys"], "| Istanze aggiornate:", stats["updated_elems"])

def import_generale(sheet):
    cats = (BuiltInCategory.OST_CableTrayFitting,
            BuiltInCategory.OST_ConduitFitting,
            BuiltInCategory.OST_ElectricalEquipment,
            BuiltInCategory.OST_ElectricalFixtures,
            BuiltInCategory.OST_LightingDevices,
            BuiltInCategory.OST_LightingFixtures)
    elems = []
    for bic in cats:
        elems.extend(list(FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType().ToElements()))
    idx = {}
    for e in elems:
        if not isinstance(e, FamilyInstance): continue
        try:
            if e.Category and e.Category.Id.IntegerValue == int(BuiltInCategory.OST_ElectricalEquipment):
                fam_raw = EEQ_family_name(e) or ""
                fam_s = (fam_raw or "").strip()
                if fam_s.startswith("MAN_EEQ_PNB_SwitchBoard") or fam_s.startswith("MAN_SEQ"):
                    continue
        except: pass
        try:
            if e.Category and e.Category.Id.IntegerValue == int(BuiltInCategory.OST_ConduitFitting):
                tname_raw = EEQ_type_name(e) or ""
                tnlc = (tname_raw or "").lower()
                if ("thermocable" in tnlc) or ("airsampling" in tnlc):
                    continue
        except: pass

        fam = norm_strong(EEQ_family_name(e) or "")
        typ = norm_strong(EEQ_type_name(e) or "")
        if not (fam or typ): continue
        idx.setdefault((fam, typ), []).append(e)

    def key_builder(headers, r):
        fcol = headers.get("Family Name"); tcol = headers.get("Type Name")
        if not (fcol and tcol): return None
        fam = norm_strong(sheet.Cells(r, fcol).Value2 or u"")
        typ = norm_strong(sheet.Cells(r, tcol).Value2 or u"")
        if not (fam or typ): return None
        # escludi famiglie dei quadri/SEQ come in export
        if fam.startswith("MAN_EEQ_PNB_SwitchBoard") or fam.startswith("MAN_SEQ"): return None
        return (fam, typ)

    rows, region, headers, err = build_row_map(sheet, 3, 5, ["Family Name","Type Name"], key_builder)
    stats = {"matched_keys":0, "updated_elems":0, "missing_param":{}, "set_count":{}, "errors":{}}
    if err: print("[GEN] Skip:", err); return
    t = Transaction(doc, "Excel→Revit | Generale"); t.Start()
    try:
        for k,(pc,uq) in rows.items():
            for e in idx.get(k, []):
                if apply_two_params(e, pc, uq, stats): stats["updated_elems"] += 1
            if k in idx: stats["matched_keys"] += 1
    finally:
        t.Commit()
    print("[GEN] Chiavi corrisposte:", stats["matched_keys"], "| Istanze aggiornate:", stats["updated_elems"])

def import_pipe(sheet):
    elems = FilteredElementCollector(doc).OfClass(Pipe).WhereElementIsNotElementType().ToElements()
    idx = {}
    for p in elems:
        t = norm_strong(EEQ_type_name(p) or "")
        dkey = pipe_diameter_key_from_elem(p)
        if not dkey: continue
        idx.setdefault((t, dkey), []).append(p)

    def key_builder(headers, r):
        tcol = headers.get("Type Name"); dcol = headers.get("Diameter")
        if not (tcol and dcol): return None
        t = norm_strong(sheet.Cells(r, tcol).Value2 or u"")
        raw = U(sheet.Cells(r, dcol).Value2 or u"")
        s = raw.replace(",", "."); m = re.search(r'(\d+(?:\.\d+)?)', s)
        dkey = ""
        if m:
            dkey = m.group(1)
            if "." in dkey: dkey = dkey.rstrip("0").rstrip(".")
        if not (t or dkey): return None
        return (t, dkey)

    rows, region, headers, err = build_row_map(sheet, 3, 5, ["Type Name","Diameter"], key_builder, numeric_cols=["Diameter"])
    stats = {"matched_keys":0, "updated_elems":0, "missing_param":{}, "set_count":{}, "errors":{}}
    if err: print("[PIPE] Skip:", err); return
    t = Transaction(doc, "Excel→Revit | Pipe"); t.Start()
    try:
        for k,(pc,uq) in rows.items():
            for e in idx.get(k, []):
                if apply_two_params(e, pc, uq, stats): stats["updated_elems"] += 1
            if k in idx: stats["matched_keys"] += 1
    finally:
        t.Commit()
    print("[PIPE] Chiavi corrisposte:", stats["matched_keys"], "| Istanze aggiornate:", stats["updated_elems"])

def import_pfit(sheet, sheet_name="Raccordi Tubi"):
    elems = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsNotElementType().ToElements()
    idx = {}
    for e in elems:
        if not isinstance(e, FamilyInstance): continue
        fam = norm_strong(EEQ_family_name(e) or "")
        typ = norm_strong(EEQ_type_name(e) or "")
        msz = fittings_max_mm_key(e)
        if not (fam or typ or msz): continue
        idx.setdefault((fam, typ, msz), []).append(e)

    def key_builder(headers, r):
        fcol = headers.get("Family Name"); tcol = headers.get("Type Name"); mcol = headers.get("MAN_Fittings_MaxSize")
        if not (fcol and tcol and mcol): return None
        fam = norm_strong(sheet.Cells(r, fcol).Value2 or u"")
        typ = norm_strong(sheet.Cells(r, tcol).Value2 or u"")
        mv  = to_float_dot(sheet.Cells(r, mcol).Value2); mk = round(mv, 6) if mv is not None else 0.0
        if not (fam or typ or mk): return None
        return (fam, typ, mk)

    rows, region, headers, err = build_row_map(sheet, 3, 5, ["Family Name","Type Name","MAN_Fittings_MaxSize"], key_builder, numeric_cols=["MAN_Fittings_MaxSize"])
    stats = {"matched_keys":0, "updated_elems":0, "missing_param":{}, "set_count":{}, "errors":{}}
    if err: print("[PFIT] Skip:", err); return
    t = Transaction(doc, "Excel→Revit | Pipe Fittings"); t.Start()
    try:
        for k,(pc,uq) in rows.items():
            for e in idx.get(k, []):
                if apply_two_params(e, pc, uq, stats): stats["updated_elems"] += 1
            if k in idx: stats["matched_keys"] += 1
    finally:
        t.Commit()
    print("[PFIT] Chiavi corrisposte:", stats["matched_keys"], "| Istanze aggiornate:", stats["updated_elems"])

def import_ducts(sheet):
    elems = FilteredElementCollector(doc).OfClass(Duct).WhereElementIsNotElementType().ToElements()
    idx = {}
    for d in elems:
        t = norm_strong(EEQ_type_name(d) or "")
        if not t: continue
        sk = duct_size_mm_key(d)
        if sk <= 0: continue
        idx.setdefault((t, sk), []).append(d)

    def key_builder(headers, r):
        tcol = headers.get("Type Name"); scol = headers.get("Width/Height - Diameter")
        if not (tcol and scol): return None
        t = norm_strong(sheet.Cells(r, tcol).Value2 or u"")
        sv = to_float_dot(sheet.Cells(r, scol).Value2); sk = round(sv, 6) if sv is not None else 0.0
        if not (t or sk): return None
        return (t, sk)

    rows, region, headers, err = build_row_map(sheet, 3, 5, ["Type Name","Width/Height - Diameter"], key_builder, numeric_cols=["Width/Height - Diameter"])
    stats = {"matched_keys":0, "updated_elems":0, "missing_param":{}, "set_count":{}, "errors":{}}
    if err: print("[DUCT] Skip:", err); return
    t = Transaction(doc, "Excel→Revit | Ducts"); t.Start()
    try:
        for k,(pc,uq) in rows.items():
            for e in idx.get(k, []):
                if apply_two_params(e, pc, uq, stats): stats["updated_elems"] += 1
            if k in idx: stats["matched_keys"] += 1
    finally:
        t.Commit()
    print("[DUCT] Chiavi corrisposte:", stats["matched_keys"], "| Istanze aggiornate:", stats["updated_elems"])

def import_dfit(sheet):
    elems = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctFitting).WhereElementIsNotElementType().ToElements()
    idx = {}
    for e in elems:
        if not isinstance(e, FamilyInstance): continue
        fam = norm_strong(EEQ_family_name(e) or "")
        typ = norm_strong(EEQ_type_name(e) or "")
        msz = fittings_max_mm_key(e)
        if not (fam or typ or msz): continue
        idx.setdefault((fam, typ, msz), []).append(e)

    def key_builder(headers, r):
        fcol = headers.get("Family Name"); tcol = headers.get("Type Name"); mcol = headers.get("MAN_Fittings_MaxSize")
        if not (fcol and tcol and mcol): return None
        fam = norm_strong(sheet.Cells(r, fcol).Value2 or u"")
        typ = norm_strong(sheet.Cells(r, tcol).Value2 or u"")
        mv  = to_float_dot(sheet.Cells(r, mcol).Value2); mk = round(mv, 6) if mv is not None else 0.0
        if not (fam or typ or mk): return None
        return (fam, typ, mk)

    rows, region, headers, err = build_row_map(sheet, 3, 5, ["Family Name","Type Name","MAN_Fittings_MaxSize"], key_builder, numeric_cols=["MAN_Fittings_MaxSize"])
    stats = {"matched_keys":0, "updated_elems":0, "missing_param":{}, "set_count":{}, "errors":{}}
    if err: print("[DFIT] Skip:", err); return
    t = Transaction(doc, "Excel→Revit | Duct Fittings"); t.Start()
    try:
        for k,(pc,uq) in rows.items():
            for e in idx.get(k, []):
                if apply_two_params(e, pc, uq, stats): stats["updated_elems"] += 1
            if k in idx: stats["matched_keys"] += 1
    finally:
        t.Commit()
    print("[DFIT] Chiavi corrisposte:", stats["matched_keys"], "| Istanze aggiornate:", stats["updated_elems"])

# ==================== Runner per sheet ====================
SHEETS_DISPATCH = {
    "Passerelle":                import_passerelle,
    "Separatore passerelle":     import_sep,
    "Cavidotti":                 import_conduits,
    "Quadri elettrici":          import_eeq,
    "Generale":                  import_generale,
    "Tubazioni":                 import_pipe,
    "Raccordi Tubi":             import_pfit,
    "Canali Rigidi":             import_ducts,
    "Fitting canali":            import_dfit,
}

def get_sheet(workbook, name):
    try: return workbook.Worksheets.Item[name]
    except: return None

# ==================== MAIN ====================
def main():
    form = RunPickerForm()
    if form.ShowDialog() != DialogResult.OK: return

    run_flags = {
        "Passerelle":            form.chkPAS.Checked,
        "Separatore passerelle": form.chkSEP.Checked,
        "Cavidotti":             form.chkCOND.Checked,
        "Quadri elettrici":      form.chkEEQ.Checked,
        "Generale":              form.chkGEN.Checked,
        "Tubazioni":             form.chkPIPE.Checked,
        "Raccordi Tubi":         form.chkPFIT.Checked,
        "Canali Rigidi":         form.chkDUCT.Checked,
        "Fitting canali":        form.chkDFIT.Checked,
    }
    if not any(run_flags.values()):
        print("Nessuna opzione selezionata. Operazione annullata.")
        return

    excel_path = pick_excel_path_once()
    if not excel_path: return

    excel = None; workbook = None
    try:
        excel = Excel.ApplicationClass()
        excel.Visible = False; excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(excel_path)

        for sheet_name, do_run in run_flags.items():
            if not do_run: continue
            fn = SHEETS_DISPATCH.get(sheet_name)
            if not fn:
                print("[{}] Nessun handler.".format(sheet_name)); continue
            sh = get_sheet(workbook, sheet_name)
            if not sh:
                print("[{}] Sheet non trovato: salto.".format(sheet_name)); continue
            try:
                fn(sh)
            except Exception as ex:
                print("[{}] Errore: {}".format(sheet_name, ex))

        # non salviamo l'Excel (solo lettura)
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
