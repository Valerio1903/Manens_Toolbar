# -*- coding: utf-8 -*-
"""
Excel -> Revit | SPE (istanze)
Importa MAN_ProductCode e MAN_BoQ_Units (istanza) da Excel agli elementi Revit per:
- GENERALE: Communication/Data/Fire Alarm/Security/Nurse Call +
            Electrical Equipment SOLO famiglie 'MAN_SEQ_' +
            Conduit Fitting SOLO type con 'ThermoCable'/'AirSampling'
  Chiavi: Family Name + Type Name
- CAVIDOTTI: Conduits SOLO type con 'ThermoCable'/'AirSampling'
  Chiavi: Type Name + Outside Diameter (mm)

Accetta sinonimi di intestazioni:
- ProductCode:  MAN_ProductCode | ProductCode | Product Code
- BoQ_Units:    MAN_BoQ_Units | BoQ_Units | BoQ Units
- Family Name:  Family Name | Family | FamilyName
- Type Name:    Type Name | Type | TypeName
- Outside Diam: Outside Diameter | OutsideDiameter | Outside Dia | OD | OD mm
"""

__title__  = 'Excel to Revit\nSPE'
__author__ = 'Valerio Mascia'

import clr, System, re
from System import String
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

# ----------------------- util testo / numeri -----------------------
def U(s):
    if s is None: return u""
    try: return s if isinstance(s, unicode) else unicode(s)
    except: return unicode(str(s))

def norm_text(s):   return U(s).strip()
def norm_strong(s): return u" ".join(norm_text(s).split())

def to_float_dot(s):
    if s is None: return None
    ss = U(s).replace(",", ".").strip()
    try: return float(ss)
    except: return None

def _feet_to_mm(val_ft):
    try:
        if _HAS_UTID:
            return float(UnitUtils.ConvertFromInternalUnits(float(val_ft), UnitTypeId.Millimeters))
        else:
            return float(UnitUtils.ConvertFromInternalUnits(float(val_ft), DisplayUnitType.DUT_MILLIMETERS))
    except:
        try: return float(val_ft) * 304.8
        except: return 0.0

# -------------------- param setter ISTANZA -----------------
def set_param_generic(param, text_val):
    if not param: return False, "missing"
    if param.IsReadOnly: return False, "read-only"
    st = param.StorageType
    s  = U(text_val).strip()
    try:
        if st == StorageType.String:
            param.Set(s); return True, None
        elif st == StorageType.Integer:
            if s == u"": param.Set(0); return True, None
            f = to_float_dot(s); iv = int(f) if f is not None else 0
            param.Set(iv); return True, None
        elif st == StorageType.Double:
            if s == u"": param.Set(0.0); return True, None
            f = to_float_dot(s); fv = f if f is not None else 0.0
            param.Set(fv); return True, None
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

# ------------------- Excel helpers (con sinonimi) -------------------
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

def get_header_col_ci(headers, candidates):
    if not headers: return None
    low = {}
    for k,v in headers.items():
        try: low[k.lower()] = v
        except: pass
    for nm in candidates:
        if not nm: continue
        c = low.get(U(nm).lower())
        if c: return c
    return None

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

def detect_data_region_by_cols(sheet, min_row, key_col_ids, empty_run_stop=20):
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

# ------------------- Revit helpers: names / diam -------------------
def elem_family_name(elem):
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

def elem_type_name(elem):
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

def conduit_outside_diam_mm_key(elem, prec=6):
    p = None
    try:
        p = elem.get_Parameter(BuiltInParameter.RBS_CONDUIT_OUTER_DIAM_PARAM)
    except:
        p = None
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

# ------------------- BUILD row map (sinonimi) -------------------
def build_row_map_with_syns(sheet, header_row, min_row, key_names_groups, key_builder, extra_numeric_names=None):
    """
    key_names_groups: lista di liste di candidati per ogni colonna chiave (es. [["Family Name","Family"], ["Type Name","Type"]])
    key_builder(headers_map, col_idxs_dict, row) -> key_tuple
    """
    headers = xl_headers_map(sheet, header_row)

    # trova colonne chiave (accetta sinonimi)
    key_cols = []
    col_idxs = {}
    for group in key_names_groups:
        c = get_header_col_ci(headers, group)
        if c: key_cols.append(c)
        # salva il primo nome canonico come chiave per col_idxs
        if group and c:
            col_idxs[group[0]] = c

    if extra_numeric_names:
        for nm_group in extra_numeric_names:
            c2 = get_header_col_ci(headers, nm_group)
            if c2:
                key_cols.append(c2)
                if nm_group and c2:
                    col_idxs[nm_group[0]] = c2

    if not key_cols:
        return {}, (min_row, min_row-1), headers, "mancano colonne chiave"

    r0, r1 = detect_data_region_by_cols(sheet, min_row, key_cols, empty_run_stop=20)
    if r1 < r0:
        return {}, (r0, r1), headers, "sheet vuoto"

    # colonne per valori
    pc_col = get_header_col_ci(headers, ["MAN_ProductCode","ProductCode","Product Code"])
    uq_col = get_header_col_ci(headers, ["MAN_BoQ_Units","BoQ_Units","BoQ Units"])
    if not (pc_col or uq_col):
        return {}, (r0, r1), headers, "mancano MAN_ProductCode / MAN_BoQ_Units (o sinonimi)"

    result = {}
    for r in range(r0, r1+1):
        key = key_builder(headers, col_idxs, r)
        if not key: continue
        pc = sheet.Cells(r, pc_col).Value2 if pc_col else None
        uq = sheet.Cells(r, uq_col).Value2 if uq_col else None
        if (pc is None or U(pc).strip()==u"") and (uq is None or U(uq).strip()==u""):
            continue
        result[key] = (U(pc).strip() if pc is not None else u"", U(uq).strip() if uq is not None else u"")
    return result, (r0, r1), headers, None

# ------------------- IMPORT: GENERALE -------------------
def import_generale(sheet):
    cats = (
        BuiltInCategory.OST_CommunicationDevices,
        BuiltInCategory.OST_DataDevices,
        BuiltInCategory.OST_FireAlarmDevices,
        BuiltInCategory.OST_NurseCallDevices,
        BuiltInCategory.OST_SecurityDevices,
        BuiltInCategory.OST_ElectricalEquipment,
        BuiltInCategory.OST_ConduitFitting
    )
    elems = []
    for bic in cats:
        try:
            coll = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType().ToElements()
            if bic == BuiltInCategory.OST_ElectricalEquipment:
                for e in coll:
                    fam = elem_family_name(e) or ""
                    if fam.strip().startswith("MAN_SEQ_"):
                        elems.append(e)
                continue
            if bic == BuiltInCategory.OST_ConduitFitting:
                for e in coll:
                    tname = elem_type_name(e) or ""
                    tlow = (tname or "").lower()
                    if ("thermocable" in tlow) or ("airsampling" in tlow):
                        elems.append(e)
                continue
            elems.extend(list(coll))
        except:
            pass

    # indice: (Family, Type)
    idx = {}
    for e in elems:
        if not isinstance(e, FamilyInstance):  # per sicurezza
            continue
        fam = norm_strong(elem_family_name(e) or "")
        typ = norm_strong(elem_type_name(e) or "")
        if not (fam or typ): continue
        idx.setdefault((fam, typ), []).append(e)

    def key_builder(headers, col_idxs, r):
        fcol = col_idxs.get("Family Name")
        tcol = col_idxs.get("Type Name")
        if not (fcol and tcol): return None
        fam = norm_strong(sheet.Cells(r, fcol).Value2 or u"")
        typ = norm_strong(sheet.Cells(r, tcol).Value2 or u"")
        if not (fam or typ): return None
        # coerente all'export: non filtriamo qui; l'indice già contiene solo i target
        return (fam, typ)

    rows, region, headers, err = build_row_map_with_syns(
        sheet,
        header_row=3,
        min_row=5,
        key_names_groups=[["Family Name","Family","FamilyName"], ["Type Name","Type","TypeName"]],
        key_builder=key_builder,
        extra_numeric_names=None
    )
    stats = {"matched_keys":0, "updated_elems":0, "missing_param":{}, "set_count":{}, "errors":{}}
    if err:
        print("[GEN] Skip: {}".format(err)); return

    t = Transaction(doc, "Excel→Revit | SPE Generale"); t.Start()
    try:
        for k,(pc,uq) in rows.items():
            lst = idx.get(k, [])
            if not lst: continue
            stats["matched_keys"] += 1
            for e in lst:
                if apply_two_params(e, pc, uq, stats):
                    stats["updated_elems"] += 1
    finally:
        t.Commit()
    print("[GEN] Chiavi corrisposte: {} | Istanze aggiornate: {}".format(stats["matched_keys"], stats["updated_elems"]))

# ------------------- IMPORT: CAVIDOTTI (Thermo/Air) -------------------
def import_cavidotti(sheet):
    elems = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Conduit).WhereElementIsNotElementType().ToElements()
    idx = {}
    for e in elems:
        t = norm_strong(elem_type_name(e) or "")
        if not t: continue
        tl = t.lower()
        if ("thermocable" not in tl) and ("airsampling" not in tl):
            continue
        dk = conduit_outside_diam_mm_key(e)
        if dk <= 0: continue
        idx.setdefault((t, dk), []).append(e)

    def key_builder(headers, col_idxs, r):
        tcol = col_idxs.get("Type Name")
        dcol = col_idxs.get("Outside Diameter")
        if not (tcol and dcol): return None
        t = norm_strong(sheet.Cells(r, tcol).Value2 or u"")
        dv = to_float_dot(sheet.Cells(r, dcol).Value2); dk = round(dv, 6) if dv is not None else 0.0
        if not (t or dk): return None
        return (t, dk)

    rows, region, headers, err = build_row_map_with_syns(
        sheet,
        header_row=3,
        min_row=5,
        key_names_groups=[["Type Name","Type","TypeName"], ["Outside Diameter","OutsideDiameter","Outside Dia","OD","OD mm"]],
        key_builder=key_builder,
        extra_numeric_names=[["Outside Diameter","OutsideDiameter","Outside Dia","OD","OD mm"]]
    )
    stats = {"matched_keys":0, "updated_elems":0, "missing_param":{}, "set_count":{}, "errors":{}}
    if err:
        print("[CAVIDOTTI] Skip: {}".format(err)); return

    t = Transaction(doc, "Excel→Revit | SPE Cavidotti"); t.Start()
    try:
        for k,(pc,uq) in rows.items():
            lst = idx.get(k, [])
            if not lst: continue
            stats["matched_keys"] += 1
            for e in lst:
                if apply_two_params(e, pc, uq, stats):
                    stats["updated_elems"] += 1
    finally:
        t.Commit()
    print("[CAVIDOTTI] Chiavi corrisposte: {} | Istanze aggiornate: {}".format(stats["matched_keys"], stats["updated_elems"]))

# ------------------- UI -------------------
class RunPickerForm(Form):
    def __init__(self):
        Form.__init__(self)
        self.Text = "Excel → Revit | SPE (istanze)"
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.ClientSize = Size(500, 180)

        lbl = Label()
        lbl.Text = "Importa ProductCode e BoQ_Units su ISTANZE (SPE)"
        lbl.Location = Point(16, 16)
        lbl.AutoSize = True
        lbl.Font = Font(self.Font, FontStyle.Bold)
        self.Controls.Add(lbl)

        self.chkGen = CheckBox()
        self.chkGen.Text = "Generale (COMM/DATA/FIRE/SEC/Nurse + EEQ MAN_SEQ_ + ConduitFitting Thermo/Air)"
        self.chkGen.Location = Point(20, 52)
        self.chkGen.AutoSize = True
        self.chkGen.Checked = True
        self.Controls.Add(self.chkGen)

        self.chkCond = CheckBox()
        self.chkCond.Text = "Cavidotti (Conduits ThermoCable/AirSampling – Outside Diameter)"
        self.chkCond.Location = Point(20, 78)
        self.chkCond.AutoSize = True
        self.chkCond.Checked = True
        self.Controls.Add(self.chkCond)

        self.btnOk = Button(); self.btnOk.Text="OK"; self.btnOk.Size=Size(100,28)
        self.btnOk.Location = Point(self.ClientSize.Width-220, 130)
        self.btnOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btnOk.DialogResult = DialogResult.OK; self.Controls.Add(self.btnOk)

        self.btnCancel = Button(); self.btnCancel.Text="Annulla"; self.btnCancel.Size=Size(100,28)
        self.btnCancel.Location = Point(self.ClientSize.Width-110, 130)
        self.btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btnCancel.DialogResult = DialogResult.Cancel; self.Controls.Add(self.btnCancel)

        self.AcceptButton = self.btnOk
        self.CancelButton = self.btnCancel

def pick_excel_path_once():
    dlg = OpenFileDialog()
    dlg.Title = "Seleziona il file Excel (SPE | Generale / Cavidotti)"
    dlg.Filter = "Excel (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls"
    dlg.Multiselect = False
    return dlg.FileName if dlg.ShowDialog() == DialogResult.OK else None

def get_sheet(workbook, name):
    try: return workbook.Worksheets.Item[name]
    except: return None

# ------------------- MAIN -------------------
def main():
    form = RunPickerForm()
    if form.ShowDialog() != DialogResult.OK:
        return

    run_gen  = form.chkGen.Checked
    run_cond = form.chkCond.Checked
    if not (run_gen or run_cond):
        print("Nessuna opzione selezionata. Operazione annullata.")
        return

    excel_path = pick_excel_path_once()
    if not excel_path:
        return

    excel = None; workbook = None
    try:
        excel = Excel.ApplicationClass()
        excel.Visible = False; excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(excel_path)

        if run_gen:
            sh = get_sheet(workbook, "Generale")
            if not sh: print("[Generale] Sheet non trovato: salto.")
            else:
                try: import_generale(sh)
                except Exception as ex:
                    print("[Generale] Errore: {}".format(ex))

        if run_cond:
            sh = get_sheet(workbook, "Cavidotti")
            if not sh: print("[Cavidotti] Sheet non trovato: salto.")
            else:
                try: import_cavidotti(sh)
                except Exception as ex:
                    print("[Cavidotti] Errore: {}".format(ex))

        # solo lettura: non salviamo Excel
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