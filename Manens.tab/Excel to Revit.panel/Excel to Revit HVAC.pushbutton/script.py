# -*- coding: utf-8 -*-
"""
Excel -> Revit | HVAC (Import)
Importa MAN_ProductCode e MAN_BoQ_Units (istanza) dagli sheet dell'export HVAC verso gli elementi Revit.

Sheet supportati e CHIAVI di corrispondenza (coerenti con l'export HVAC):
- Tubazioni (Pipe)                        [Type Name + Diameter]
- Isolante Tubazioni (Pipe Insulations)   [Type Name + Insulation Thickness + Pipe Size]
- Raccordi Tubi (Pipe Fittings)           [Family Name + Type Name + MAN_Fittings_MaxSize mm]
- Apparecchiature Mec (Mech. Equipment)   [Family Name + Type Name + MAN_Type_Code]
- Generale (PA/PF/Sprinklers/DA)          [Family Name + Type Name]
- Canali Rigidi (Ducts)                   [Type Name + Width/Height - Diameter (max mm)]
- Isolamento canali (Duct Insulations)    [Type Name + Insulation Thickness]
- Fitting canali (Duct Fittings)          [Family Name + Type Name + MAN_Fittings_MaxSize mm]
- Canali Flessibili (Flex Ducts)          [Type Name + Diameter]
"""
__title__  = 'Excel to Revit\nHVAC'
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
from Autodesk.Revit.DB.Plumbing import Pipe, PipeInsulation
from Autodesk.Revit.DB.Mechanical import Duct, FlexDuct, DuctInsulation

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


# ==================== Util testo / parsing ====================
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
    except:
        m = re.search(r'(-?\d+(?:\.\d+)?)', ss)
        if m:
            try: return float(m.group(1))
            except: return None
    return None

# =============== Param setter (istanza, solo STRING) ===============
PARAM_NAMES = ("MAN_ProductCode", "MAN_BoQ_Units")

def set_param_generic(param, text_val):
    """Imposta sempre come stringa (trim), con fallback SetValueString."""
    if not param:
        return False, "missing"
    if param.IsReadOnly:
        return False, "read-only"

    s = U(text_val).strip()  # sempre unicode/stringa
    try:
        # caso normale: parametro di tipo STRING
        param.Set(s)
        return True, None
    except Exception as ex:
        # se per qualsiasi motivo non è STRING, prova con SetValueString
        try:
            param.SetValueString(s)
            return True, None
        except Exception as ex2:
            return False, "exc: {}".format(ex2)

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
            out.append(norm_text(val))
        return out
    if isinstance(data, (tuple, list)):
        for row in data:
            try: val = row[0]
            except: val = row
            out.append(norm_text(val))
        return out
    out.append(norm_text(data))
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

def build_row_map(sheet, header_row, min_row, key_cols_names, key_builder, numeric_cols=None):
    """
    Ritorna: { key_tuple : (prodcode, units) }
    key_cols_names: elenco intestazioni coinvolte per r1 detection
    key_builder: funzione (headers, row_index) -> key_tuple or None
    numeric_cols: opzionali per migliorare detection
    """
    headers = xl_headers_map(sheet, header_row)
    cols_for_detection = []
    for nm in key_cols_names or []:
        c = headers.get(nm)
        if c: cols_for_detection.append(c)
    if numeric_cols:
        for nm in numeric_cols:
            c = headers.get(nm)
            if c: cols_for_detection.append(c)
    if not cols_for_detection:
        return {}, (min_row, min_row-1), headers, "mancano colonne chiave"

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
            continue
        result[key] = (U(pc).strip() if pc is not None else u"", U(uq).strip() if uq is not None else u"")
    return result, (r0, r1), headers, None


# ==================== Helpers (chiavi dagli elementi) ====================
def type_name_from_instance(elem):
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

def family_name_from_instance(elem):
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

def _feet_to_mm(val_ft):
    try:
        if _HAS_UTID:
            return float(UnitUtils.ConvertFromInternalUnits(float(val_ft), UnitTypeId.Millimeters))
        else:
            return float(UnitUtils.ConvertFromInternalUnits(float(val_ft), DisplayUnitType.DUT_MILLIMETERS))
    except:
        try: return float(val_ft) * 304.8
        except: return 0.0

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

def flex_diam_mm_key(elem, prec=6):
    try:
        p = elem.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
        if not p: return 0.0
        d = p.AsDouble()
        if d and d > 0:
            return round(_feet_to_mm(d), prec)
        s = (p.AsString() or p.AsValueString() or "").replace(",", ".").strip()
        if not s: return 0.0
        f = to_float_dot(s)
        return round(float(f), prec) if f is not None else 0.0
    except: return 0.0

def eq_instance_param(elem, name):
    try:
        q = elem.LookupParameter(name)
        if not q: return ""
        s = q.AsString()
        if s: return norm_text(s)
        vs = q.AsValueString()
        if vs: return norm_text(vs)
    except: pass
    return ""


# ==================== BLOCCHI: INDICI + IMPORT ====================

# 1) TUBAZIONI (Pipe) : [TypeName + Diameter]
def import_pipe(sheet):
    elems = FilteredElementCollector(doc).OfClass(Pipe).WhereElementIsNotElementType().ToElements()
    idx = {}
    for p in elems:
        t = norm_strong(type_name_from_instance(p) or "")
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
    print("[PIPE] Chiavi:", stats["matched_keys"], "| Istanze aggiornate:", stats["updated_elems"])

# 2) ISOLANTE TUBAZIONI (Pipe Insulations): [Type + Thickness + Pipe Size]
def import_pipe_ins(sheet):
    elems = FilteredElementCollector(doc).OfClass(PipeInsulation).WhereElementIsNotElementType().ToElements()
    def is_host_pipe(ins):
        try:
            hid = ins.HostElementId
            if hid and hid.IntegerValue > 0:
                host = doc.GetElement(hid)
                return isinstance(host, Pipe)
        except: pass
        return False
    idx = {}
    for ins in elems:
        t = norm_strong(type_name_from_instance(ins) or "")
        # thickness (mm)
        thkey = ""
        try:
            pth = ins.get_Parameter(BuiltInParameter.RBS_INSULATION_THICKNESS_FOR_PIPE)
            if pth:
                try:
                    d = pth.AsDouble()
                    if d and d>0: thkey = ("%.6f" % _feet_to_mm(d)).rstrip("0").rstrip(".")
                except:
                    s = (pth.AsString() or pth.AsValueString() or "").replace(",", ".")
                    f = to_float_dot(s); thkey = ("%.6f" % f).rstrip("0").rstrip(".") if f is not None else ""
        except: pass
        # pipe size (mm number inside the string "Φ...")
        szkey = ""
        try:
            psz = ins.get_Parameter(BuiltInParameter.RBS_PIPE_CALCULATED_SIZE)
            s = (psz.AsString() or psz.AsValueString() or "")
            s = re.sub(u"[ΦφØø⌀]", u"", s)
            m = re.search(r'(\d+(?:[.,]\d+)?)', s.replace(",", "."))
            if m:
                z = m.group(1)
                if "." in z: z = z.rstrip("0").rstrip(".")
                szkey = z
        except: pass
        if not (t or thkey or szkey): continue
        idx.setdefault((t, thkey or "0", szkey), []).append(ins)

    def key_builder(headers, r):
        tcol = headers.get("Type Name"); thcol = headers.get("Insulation Thickness"); szcol = headers.get("Pipe Size")
        if not (tcol and thcol and szcol): return None
        t  = norm_strong(sheet.Cells(r, tcol).Value2 or u"")
        th = to_float_dot(sheet.Cells(r, thcol).Value2); thk = ""
        if th is not None: thk = ("%.6f" % th).rstrip("0").rstrip(".")
        sz = U(sheet.Cells(r, szcol).Value2 or u"")
        sz = re.sub(u"[ΦφØø⌀]", u"", sz)
        m = re.search(r'(\d+(?:[.,]\d+)?)', sz.replace(",", "."))
        szk = ""
        if m:
            szk = m.group(1)
            if "." in szk: szk = szk.rstrip("0").rstrip(".")
        if not (t or thk or szk): return None
        return (t, thk or "0", szk)

    rows, region, headers, err = build_row_map(sheet, 3, 5, ["Type Name","Insulation Thickness","Pipe Size"], key_builder, numeric_cols=["Insulation Thickness","Pipe Size"])
    stats = {"matched_keys":0, "updated_elems":0, "missing_param":{}, "set_count":{}, "errors":{}}
    if err: print("[PIPE INS] Skip:", err); return
    t = Transaction(doc, "Excel→Revit | Pipe Insulation"); t.Start()
    try:
        for k,(pc,uq) in rows.items():
            for e in idx.get(k, []):
                if apply_two_params(e, pc, uq, stats): stats["updated_elems"] += 1
            if k in idx: stats["matched_keys"] += 1
    finally:
        t.Commit()
    print("[PIPE INS] Chiavi:", stats["matched_keys"], "| Istanze aggiornate:", stats["updated_elems"])

# 3) RACCORDI TUBI (Pipe Fittings): [Family + Type + MaxSize mm]
def import_pfit(sheet):
    elems = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsNotElementType().ToElements()
    idx = {}
    for e in elems:
        if not isinstance(e, FamilyInstance): continue
        fam = norm_strong(family_name_from_instance(e) or "")
        typ = norm_strong(type_name_from_instance(e) or "")
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
    print("[PFIT] Chiavi:", stats["matched_keys"], "| Istanze aggiornate:", stats["updated_elems"])

# 4) APPARECCHIATURE MEC (Mechanical Equipment): [Family + Type + MAN_Type_Code]
def import_meq(sheet):
    elems = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements()
    idx = {}
    for e in elems:
        if not isinstance(e, FamilyInstance): continue
        fam = norm_strong(family_name_from_instance(e) or "")
        typ = norm_strong(type_name_from_instance(e) or "")
        code = norm_strong(eq_instance_param(e, "MAN_Type_Code") or "")
        idx.setdefault((fam, typ, code), []).append(e)

    def key_builder(headers, r):
        fcol = headers.get("Family Name"); tcol = headers.get("Type Name"); ccol = headers.get("MAN_Type_Code")
        if not (fcol and tcol and ccol): return None
        fam = norm_strong(sheet.Cells(r, fcol).Value2 or u"")
        typ = norm_strong(sheet.Cells(r, tcol).Value2 or u"")
        code = norm_strong(sheet.Cells(r, ccol).Value2 or u"")
        if not (fam or typ or code): return None
        return (fam, typ, code)

    rows, region, headers, err = build_row_map(sheet, 3, 5, ["Family Name","Type Name","MAN_Type_Code"], key_builder)
    stats = {"matched_keys":0, "updated_elems":0, "missing_param":{}, "set_count":{}, "errors":{}}
    if err: print("[MEQ] Skip:", err); return
    t = Transaction(doc, "Excel→Revit | Mechanical Equipment"); t.Start()
    try:
        for k,(pc,uq) in rows.items():
            for e in idx.get(k, []):
                if apply_two_params(e, pc, uq, stats): stats["updated_elems"] += 1
            if k in idx: stats["matched_keys"] += 1
    finally:
        t.Commit()
    print("[MEQ] Chiavi:", stats["matched_keys"], "| Istanze aggiornate:", stats["updated_elems"])

# 5) GENERALE (DuctTerminal / DuctAccessory / PipeAccessory / PlumbingFixtures / Sprinklers): [Family + Type]
def import_generale(sheet):
    cats = (BuiltInCategory.OST_DuctTerminal,
            BuiltInCategory.OST_DuctAccessory,
            BuiltInCategory.OST_PipeAccessory,
            BuiltInCategory.OST_PlumbingFixtures,
            BuiltInCategory.OST_Sprinklers)
    elems = []
    for bic in cats:
        elems.extend(list(FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType().ToElements()))
    idx = {}
    for e in elems:
        if not isinstance(e, FamilyInstance): continue
        fam = norm_strong(family_name_from_instance(e) or "")
        typ = norm_strong(type_name_from_instance(e) or "")
        if not (fam or typ): continue
        idx.setdefault((fam, typ), []).append(e)

    def key_builder(headers, r):
        fcol = headers.get("Family Name"); tcol = headers.get("Type Name")
        if not (fcol and tcol): return None
        fam = norm_strong(sheet.Cells(r, fcol).Value2 or u"")
        typ = norm_strong(sheet.Cells(r, tcol).Value2 or u"")
        if not (fam or typ): return None
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
    print("[GEN] Chiavi:", stats["matched_keys"], "| Istanze aggiornate:", stats["updated_elems"])

# 6) CANALI RIGIDI (Ducts): [Type + MaxDim/Diameter mm]
def import_ducts(sheet):
    elems = FilteredElementCollector(doc).OfClass(Duct).WhereElementIsNotElementType().ToElements()
    idx = {}
    for d in elems:
        t = norm_strong(type_name_from_instance(d) or "")
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
    print("[DUCT] Chiavi:", stats["matched_keys"], "| Istanze aggiornate:", stats["updated_elems"])

# 7) ISOLAMENTO CANALI (Duct Insulation): [Type + Thickness]
def import_duct_ins(sheet):
    elems = FilteredElementCollector(doc).OfClass(DuctInsulation).WhereElementIsNotElementType().ToElements()
    def is_host_duct(ins):
        try:
            hid = ins.HostElementId
            if hid and hid.IntegerValue > 0:
                host = doc.GetElement(hid)
                return isinstance(host, Duct)
        except: pass
        return False
    idx = {}
    for ins in elems:
        t  = norm_strong(type_name_from_instance(ins) or "")
        th = ""
        pth = None
        # diversi BuiltInParameter per versioni/famiglie: prova in cascata
        for bipname in ("RBS_INSULATION_THICKNESS_FOR_DUCT", "RBS_INSULATION_THICKNESS", "RBS_INSULATION_THICKNESS_FOR_PIPE"):
            try:
                pth = ins.get_Parameter(getattr(BuiltInParameter, bipname))
                if pth: break
            except: pass
        if pth:
            try:
                d = pth.AsDouble()
                if d and d>0: th = ("%.6f" % _feet_to_mm(d)).rstrip("0").rstrip(".")
            except:
                s = (pth.AsString() or pth.AsValueString() or "").replace(",", ".")
                f = to_float_dot(s); th = ("%.6f" % f).rstrip("0").rstrip(".") if f is not None else ""
        if not (t or th): continue
        idx.setdefault((t, th), []).append(ins)

    def key_builder(headers, r):
        tcol = headers.get("Type Name"); thcol = headers.get("Insulation Thickness")
        if not (tcol and thcol): return None
        t  = norm_strong(sheet.Cells(r, tcol).Value2 or u"")
        th = to_float_dot(sheet.Cells(r, thcol).Value2); thk = ""
        if th is not None: thk = ("%.6f" % th).rstrip("0").rstrip(".")
        if not (t or thk): return None
        return (t, thk)

    rows, region, headers, err = build_row_map(sheet, 3, 5, ["Type Name","Insulation Thickness"], key_builder, numeric_cols=["Insulation Thickness"])
    stats = {"matched_keys":0, "updated_elems":0, "missing_param":{}, "set_count":{}, "errors":{}}
    if err: print("[DUCT INS] Skip:", err); return
    t = Transaction(doc, "Excel→Revit | Duct Insulation"); t.Start()
    try:
        for k,(pc,uq) in rows.items():
            for e in idx.get(k, []):
                if apply_two_params(e, pc, uq, stats): stats["updated_elems"] += 1
            if k in idx: stats["matched_keys"] += 1
    finally:
        t.Commit()
    print("[DUCT INS] Chiavi:", stats["matched_keys"], "| Istanze aggiornate:", stats["updated_elems"])

# 8) FITTING CANALI (Duct Fittings): [Family + Type + MaxSize mm]
def import_dfit(sheet):
    elems = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctFitting).WhereElementIsNotElementType().ToElements()
    idx = {}
    for e in elems:
        if not isinstance(e, FamilyInstance): continue
        fam = norm_strong(family_name_from_instance(e) or "")
        typ = norm_strong(type_name_from_instance(e) or "")
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
    print("[DFIT] Chiavi:", stats["matched_keys"], "| Istanze aggiornate:", stats["updated_elems"])

# 9) CANALI FLESSIBILI (Flex Ducts): [Type + Diameter]
def import_flex(sheet):
    elems = FilteredElementCollector(doc).OfClass(FlexDuct).WhereElementIsNotElementType().ToElements()
    idx = {}
    for d in elems:
        t = norm_strong(type_name_from_instance(d) or "")
        dk = flex_diam_mm_key(d)
        if dk <= 0: continue
        idx.setdefault((t, dk), []).append(d)

    # accetta sia "Diameter" che eventuali varianti ("Width/Height - Diameter" usato per rigid)
    def key_builder(headers, r):
        tcol = headers.get("Type Name")
        dcol = headers.get("Diameter") or headers.get("Width/Height - Diameter")
        if not (tcol and dcol): return None
        t = norm_strong(sheet.Cells(r, tcol).Value2 or u"")
        d = to_float_dot(sheet.Cells(r, dcol).Value2); dk = round(d, 6) if d is not None else 0.0
        if not (t or dk): return None
        return (t, dk)

    rows, region, headers, err = build_row_map(sheet, 3, 5, ["Type Name", "Diameter"], key_builder, numeric_cols=["Diameter"])
    stats = {"matched_keys":0, "updated_elems":0, "missing_param":{}, "set_count":{}, "errors":{}}
    if err: print("[FLEX] Skip:", err); return
    t = Transaction(doc, "Excel→Revit | Flex Ducts"); t.Start()
    try:
        for k,(pc,uq) in rows.items():
            for e in idx.get(k, []):
                if apply_two_params(e, pc, uq, stats): stats["updated_elems"] += 1
            if k in idx: stats["matched_keys"] += 1
    finally:
        t.Commit()
    print("[FLEX] Chiavi:", stats["matched_keys"], "| Istanze aggiornate:", stats["updated_elems"])


# ==================== UI + MAIN ====================
class RunPickerForm(Form):
    def __init__(self):
        Form.__init__(self)
        self.Text = "Revit → Excel | Seleziona cosa esportare"
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False; self.MinimizeBox = False
        self.ClientSize = Size(430, 360)

        lbl = Label(); lbl.Text = "Scegli le esportazioni da eseguire:"
        lbl.Location = Point(16, 16); lbl.AutoSize = True
        lbl.Font = Font(self.Font, FontStyle.Bold)
        self.Controls.Add(lbl)

        self._y = 48; self._dy = 26
        def addchk(text, checked=True):
            c = CheckBox(); c.Text = text
            c.Location = Point(20, self._y); c.AutoSize = True; c.Checked = checked
            self.Controls.Add(c); self._y += self._dy
            return c

        self.chkPIPE = addchk("Pipe (Type → Diameter)")
        self.chkINS  = addchk("Pipe Insulations (Type → Thickness/Size)")
        self.chkFIT  = addchk("Pipe Fittings (Family/Type → MaxSize mm)")
        self.chkMEQ  = addchk("Apparecchiature Mec (Mechanical Equipment)")
        self.chkGEN  = addchk("Generale (Pipe Accessories / Plumbing Fixtures / Sprinklers)")
        self.chkDUCT = addchk("Canali Rigidi (Ducts – Width/Height o Diameter)")
        self.chkDINS = addchk("Isolamento canali (Duct Insulations)")
        self.chkDFT  = addchk("Duct Fittings (Family/Type → MaxSize mm)")
        self.chkFXD  = addchk("Canali Flessibili (Type → Diameter)")

        self.btnOk = Button(); self.btnOk.Text="OK"; self.btnOk.Size=Size(100,28)
        self.btnOk.Location = Point(self.ClientSize.Width - 220, 318)
        self.btnOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btnOk.DialogResult = DialogResult.OK; self.Controls.Add(self.btnOk)

        self.btnCancel = Button(); self.btnCancel.Text="Annulla"; self.btnCancel.Size=Size(100,28)
        self.btnCancel.Location = Point(self.ClientSize.Width - 110, 318)
        self.btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btnCancel.DialogResult = DialogResult.Cancel; self.Controls.Add(self.btnCancel)

        self.AcceptButton = self.btnOk; self.CancelButton = self.btnCancel

def pick_excel_path_once():
    dlg = OpenFileDialog()
    dlg.Title  = "Seleziona il file Excel (stesse sheet dell'export HVAC)"
    dlg.Filter = "Excel (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls"
    dlg.Multiselect = False
    return dlg.FileName if dlg.ShowDialog() == DialogResult.OK else None

SHEETS_DISPATCH = {
    "Tubazioni":             import_pipe,
    "Isolante Tubazioni":    import_pipe_ins,
    "Raccordi Tubi":         import_pfit,
    "Apparecchiature Mec":   import_meq,
    "Generale":              import_generale,
    "Canali Rigidi":         import_ducts,
    "Isolamento canali":     import_duct_ins,
    "Fitting canali":        import_dfit,
    "Canali Flessibili":     import_flex,
}

def get_sheet(workbook, name):
    try: return workbook.Worksheets.Item[name]
    except: return None

def main():
    form = RunPickerForm()
    if form.ShowDialog() != DialogResult.OK: return

    run_flags = {
        "Tubazioni":             form.chkPIPE.Checked,
        "Isolante Tubazioni":    form.chkINS.Checked,
        "Raccordi Tubi":         form.chkFIT.Checked,
        "Apparecchiature Mec":   form.chkMEQ.Checked,
        "Generale":              form.chkGEN.Checked,
        "Canali Rigidi":         form.chkDUCT.Checked,
        "Isolamento canali":     form.chkDINS.Checked,
        "Fitting canali":        form.chkDFT.Checked,
        "Canali Flessibili":     form.chkFXD.Checked,
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