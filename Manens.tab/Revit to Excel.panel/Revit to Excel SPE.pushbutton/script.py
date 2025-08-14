# -*- coding: utf-8 -*-
"""
COMBO (unico Excel) con CHECKBOX UI: Cable Trays + Separators + PanelBoards + Electrical -> Excel
Un solo dialog per scegliere il file Excel
Checkbox per scegliere cosa eseguire (Cable Trays / PanelBoards / Conduit / Fixtures)
Cinque blocchi separati con funzioni rinominate per evitare collisioni
"""

__title__ = 'Revit to Excel\nSPE'
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
    FilteredElementCollector, BuiltInParameter, BuiltInCategory, FamilyInstance
)
from Autodesk.Revit.DB.Plumbing import Pipe, PipeInsulation
from Autodesk.Revit.DB.Mechanical import DuctInsulation, Duct, FlexDuct
from Autodesk.Revit.DB.Electrical import CableTray, Conduit
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


# ============================================================
# ======================= UI CHECKBOX ========================
# ============================================================
class RunPickerForm(Form):
    def __init__(self):
        Form.__init__(self)
        self.Text = "Revit → Excel | Seleziona cosa esportare"
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.ClientSize = Size(420, 200)

        self.lbl = Label()
        self.lbl.Text = "Scegli le esportazioni da eseguire:"
        self.lbl.Location = Point(16, 16)
        self.lbl.AutoSize = True
        self.lbl.Font = Font(self.Font, FontStyle.Bold)
        self.Controls.Add(self.lbl)

        self.chkGen = CheckBox()
        self.chkGen.Text = "Generale (Communication / Data / Fire Alarm / Security Devices)"
        self.chkGen.Location = Point(20, 50)
        self.chkGen.AutoSize = True
        self.chkGen.Checked = True
        self.Controls.Add(self.chkGen)

        self.chkCond = CheckBox()
        self.chkCond.Text = "AirSampling ThermoCable (Conduits – Outside Diameter)"
        self.chkCond.Location = Point(20, 78)
        self.chkCond.AutoSize = True
        self.chkCond.Checked = True
        self.Controls.Add(self.chkCond)

        self.btnOk = Button()
        self.btnOk.Text = "OK"
        self.btnOk.Size = Size(100, 28)
        self.btnOk.Location = Point(self.ClientSize.Width - 220, 150)
        self.btnOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btnOk.DialogResult = DialogResult.OK
        self.Controls.Add(self.btnOk)

        self.btnCancel = Button()
        self.btnCancel.Text = "Annulla"
        self.btnCancel.Size = Size(100, 28)
        self.btnCancel.Location = Point(self.ClientSize.Width - 110, 150)
        self.btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btnCancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(self.btnCancel)

        self.AcceptButton = self.btnOk
        self.CancelButton = self.btnCancel


def pick_excel_path_once():
    dlg = OpenFileDialog()
    dlg.Title = "Seleziona il file Excel (unico per COMMUNICATION DEVICES / DATA DEVICES / FIRE ALARM DEVICES / SECURITY DEVICES )"
    dlg.Filter = "Excel (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls"
    dlg.Multiselect = False
    return dlg.FileName if dlg.ShowDialog() == DialogResult.OK else None


# ============================================================
# ========== BLOCCO 1 — GENERALE (CD / SD / Devices) ======
# ============================================================
GEN_SHEET_NAME = "Generale"
GEN_HEADER_ROW = 3
GEN_MIN_START_DATA_ROW = 5
GEN_HEADERS = [
    "Category",
    "Family Name",
    "Type Name",
    "MAN_TypeDescription_IT",
    "MAN_FamilyTypePrefix",
]
GEN_EMPTY_RUN_STOP = 20

# ----------------------- Utils testo ------------------------
def GEN_u(s):
    if s is None: return u""
    try: return s if isinstance(s, unicode) else unicode(s)
    except: return unicode(str(s))

def GEN_norm_text(s):
    v = GEN_u(s).strip()
    if v == u"": 
        return v
    # se il valore è "numeric-like" (es. 101, 101.0, 101,0) rendilo testo canonico
    vv = v.replace(",", ".")
    try:
        import re
        if re.match(r'^\d+(?:\.\d+)?$', vv):
            f = float(vv)
            if f.is_integer():
                return unicode(int(f))  # "101.0" -> "101"
            else:
                # rimuovi zeri finali e il punto se inutili: "101.50" -> "101.5"
                return unicode(vv.rstrip("0").rstrip("."))
    except:
        pass
    return v

def GEN_norm_text_strong(s):
    # compattazione spazi + normalizzazione numerica
    return u" ".join(GEN_norm_text(s).split())

# ----------------- Lettura proprietà Revit ------------------
def GEN_category_name(elem):
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

def GEN_type_name(elem):
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

def GEN_family_name(elem):
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

# cache descrizione di tipo
_GEN_type_desc_cache = {}
def GEN_type_desc(elem):
    try:
        p = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
        if not p: return ""
        tid = p.AsElementId()
        if not tid or tid.IntegerValue <= 0: return ""
        tid_i = tid.IntegerValue
        if tid_i in _GEN_type_desc_cache: return _GEN_type_desc_cache[tid_i]
        t = doc.GetElement(tid)
        val = ""
        if t:
            pp = t.LookupParameter("MAN_TypeDescription_IT")
            if pp: val = (pp.AsString() or "") or (pp.AsValueString() or "")
        _GEN_type_desc_cache[tid_i] = val
        return val
    except:
        return ""

def GEN_type_param_text(elem, param_name):
    try:
        p = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
        if not p: return ""
        tid = p.AsElementId()
        if not tid or tid.IntegerValue <= 0: return ""
        t = doc.GetElement(tid)
        if not t: return ""
        q = t.LookupParameter(param_name)
        if not q: return ""
        s = q.AsString()
        if s: return GEN_norm_text(s)
        vs = q.AsValueString()
        if vs: return GEN_norm_text(vs)
        return ""
    except:
        return ""

# ---------------------- Excel helpers -----------------------
def GEN_get_sheet_or_create(workbook, name):
    try:
        return workbook.Worksheets.Item[name]
    except:
        sh = workbook.Worksheets.Add()
        sh.Name = name
        return sh

def GEN_ensure_headers(sheet):
    last_col = sheet.Cells(GEN_HEADER_ROW, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
    if last_col < 1: last_col = 1
    headers = {}
    for c in range(1, last_col + 1):
        v = sheet.Cells(GEN_HEADER_ROW, c).Value2
        try:
            nm = v.strip() if isinstance(v, String) else None
        except:
            nm = None
        if nm: headers[nm] = c
    next_col = last_col + 1
    for h in GEN_HEADERS:
        if h not in headers:
            sheet.Cells(GEN_HEADER_ROW, next_col).Value2 = h
            headers[h] = next_col
            next_col += 1
    return headers

def GEN_read_column_block(sheet, col, r0, r1):
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
            out.append(GEN_u(val).strip())
        return out
    if isinstance(data, (tuple, list)):
        for row in data:
            try: val = row[0]
            except: val = row
            out.append(GEN_u(val).strip())
        return out
    out.append(GEN_u(data).strip())
    return out

def GEN_detect_data_region(sheet, headers):
    fam_col = headers["Family Name"]
    typ_col = headers["Type Name"]
    r0 = GEN_MIN_START_DATA_ROW
    step = 2000
    curr_start = r0
    last_data_row = r0 - 1
    empty_run = 0
    max_rows = sheet.Rows.Count
    while curr_start <= max_rows:
        r1_try = min(max_rows, curr_start + step - 1)
        f_vals = GEN_read_column_block(sheet, fam_col, curr_start, r1_try)
        t_vals = GEN_read_column_block(sheet, typ_col, curr_start, r1_try)
        block_len = max(len(f_vals), len(t_vals))
        for i in range(block_len):
            f = f_vals[i] if i < len(f_vals) else ""
            t = t_vals[i] if i < len(t_vals) else ""
            if (f or t):
                last_data_row = curr_start + i
                empty_run = 0
            else:
                empty_run += 1
                if empty_run >= GEN_EMPTY_RUN_STOP:
                    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)
        curr_start = r1_try + 1
    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)

def GEN_build_existing_index(sheet, headers):
    (r0, r1) = GEN_detect_data_region(sheet, headers)
    if r1 < r0:
        return {}, (r0, r1)
    fam_col = headers["Family Name"]; typ_col = headers["Type Name"]
    col_f = GEN_read_column_block(sheet, fam_col, r0, r1)
    col_t = GEN_read_column_block(sheet, typ_col, r0, r1)
    index = {}
    n = max(len(col_f), len(col_t))
    for i in range(n):
        f = GEN_norm_text_strong(col_f[i] if i < len(col_f) else "")
        t = GEN_norm_text_strong(col_t[i] if i < len(col_t) else "")
        if f or t:
            index[(f, t)] = r0 + i
    return index, (r0, r1)

def GEN_first_empty_row_after_region(region_tuple):
    r0, r1 = region_tuple
    if r1 < r0: return GEN_MIN_START_DATA_ROW
    return max(GEN_MIN_START_DATA_ROW, r1 + 1)

def GEN_chunk_consecutive_rows(sorted_rows):
    runs = []
    if not sorted_rows: return runs
    start_r = prev_r = sorted_rows[0]
    for r in sorted_rows[1:]:
        if r == prev_r + 1:
            prev_r = r
        else:
            runs.append((start_r, prev_r))
            start_r = prev_r = r
    runs.append((start_r, prev_r))
    return runs

def GEN_delete_rows_batched(sheet, rows_to_delete):
    if not rows_to_delete: return 0
    rows_sorted = sorted(rows_to_delete, reverse=True)
    runs = GEN_chunk_consecutive_rows(rows_sorted)
    count = 0
    for (r_start, r_end) in runs:
        rng = sheet.Range[sheet.Rows[r_start], sheet.Rows[r_end]]
        rng.EntireRow.Delete()
        count += (r_start - r_end + 1)
    return count

def GEN_write_updates_batched(sheet, headers, updates):
    if not updates: return 0
    updates_sorted = sorted(updates, key=lambda x: x[0])
    runs = []
    start = None; buf = []
    for row, vals in updates_sorted:
        if start is None:
            start = row; prev = row; buf = [vals]
        elif row == prev + 1:
            buf.append(vals); prev = row
        else:
            runs.append((start, prev, buf))
            start = row; prev = row; buf = [vals]
    if start is not None:
        runs.append((start, prev, buf))

    col_map = {
        "Category": headers["Category"],
        "Family Name": headers["Family Name"],
        "Type Name": headers["Type Name"],
        "MAN_TypeDescription_IT": headers["MAN_TypeDescription_IT"],
        "MAN_FamilyTypePrefix": headers["MAN_FamilyTypePrefix"],
    }
    keys = ["Category","Family Name","Type Name","MAN_TypeDescription_IT","MAN_FamilyTypePrefix"]
    for (r0, r1, vals_list) in runs:
        n = (r1 - r0 + 1)
        for j, key in enumerate(keys):
            col = col_map[key]
            rng = sheet.Range[sheet.Cells(r0, col), sheet.Cells(r1, col)]
            data = Array.CreateInstance(Object, n, 1)
            for i in range(n):
                data[i,0] = GEN_u(vals_list[i][j])
            rng.Value2 = data
    return len(updates)

def GEN_write_appends(sheet, start_row, headers, rows_data):
    if not rows_data: return 0
    if start_row < GEN_MIN_START_DATA_ROW:
        start_row = GEN_MIN_START_DATA_ROW
    cols = [headers[h] for h in GEN_HEADERS]
    n_rows = len(rows_data); n_cols = len(cols)
    data = Array.CreateInstance(Object, n_rows, n_cols)
    for i in range(n_rows):
        r = rows_data[i]
        for j in range(n_cols):
            data[i, j] = GEN_u(r[j])
    first_cell = sheet.Cells(start_row, min(cols))
    last_cell  = sheet.Cells(start_row + n_rows - 1, max(cols))
    dest = sheet.Range[first_cell, last_cell]
    expected_span = max(cols) - min(cols) + 1
    if expected_span == n_cols:
        dest.Value2 = data
    else:
        for idx, col in enumerate(cols):
            col_rng = sheet.Range(sheet.Cells(start_row, col), sheet.Cells(start_row + n_rows - 1, col))
            col_data = Array.CreateInstance(Object, n_rows, 1)
            for r in range(n_rows):
                col_data[r, 0] = GEN_u(rows_data[r][idx])
            col_rng.Value2 = col_data
    return n_rows

def GEN_sort_data_region(sheet, headers):
    r0 = GEN_MIN_START_DATA_ROW
    first_col = 1
    last_col = sheet.Cells(GEN_HEADER_ROW, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

    fam_col = headers["Family Name"]
    typ_col = headers["Type Name"]

    last_row = 0
    for col in (fam_col, typ_col):
        lr = sheet.Cells(sheet.Rows.Count, col).End(Excel.XlDirection.xlUp).Row
        if lr > last_row: last_row = lr
    if last_row < r0: return

    data_rng = sheet.Range(sheet.Cells(r0, first_col), sheet.Cells(last_row, last_col))

    sort = sheet.Sort
    sort.SortFields.Clear()
    sort.SortFields.Add(Key=sheet.Range(sheet.Cells(r0, fam_col), sheet.Cells(last_row, fam_col)),
                        SortOn=Excel.XlSortOn.xlSortOnValues, Order=Excel.XlSortOrder.xlAscending,
                        DataOption=Excel.XlSortDataOption.xlSortNormal)
    sort.SortFields.Add(Key=sheet.Range(sheet.Cells(r0, typ_col), sheet.Cells(last_row, typ_col)),
                        SortOn=Excel.XlSortOn.xlSortOnValues, Order=Excel.XlSortOrder.xlAscending,
                        DataOption=Excel.XlSortDataOption.xlSortNormal)
    sort.SetRange(data_rng)
    sort.Header = Excel.XlYesNoGuess.xlNo
    sort.MatchCase = False
    sort.Orientation = Excel.XlSortOrientation.xlSortColumns
    sort.Apply()

# -------------------------- RUN -----------------------------
def run_general_into_workbook(workbook):
    elems = []
    cats = (BuiltInCategory.OST_CommunicationDevices,
        BuiltInCategory.OST_ConduitFitting,
        BuiltInCategory.OST_DataDevices,
        BuiltInCategory.OST_ElectricalEquipment,
        BuiltInCategory.OST_FireAlarmDevices,
        BuiltInCategory.OST_NurseCallDevices,
        BuiltInCategory.OST_SecurityDevices)

    for bic in cats:
        coll = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType().ToElements()

        if bic == BuiltInCategory.OST_ElectricalEquipment:
            # SOLO family name che iniziano con "MAN_SEQ_"
            for e in coll:
                fam = GEN_family_name(e) or ""
                if fam.strip().startswith("MAN_SEQ_"):
                    elems.append(e)
            continue

        if bic == BuiltInCategory.OST_ConduitFitting:
            # SOLO type name che CONTENGONO "ThermoCable" o "AirSampling" (case-insensitive)
            for e in coll:
                tname = GEN_type_name(e) or ""
                tlow = tname.lower()
                if ("thermocable" in tlow) or ("airsampling" in tlow):
                    elems.append(e)
            continue

        # tutte le altre categorie: nessun prefiltro
        elems.extend(list(coll))


    groups = {}
    for e in elems:
        if not isinstance(e, FamilyInstance):
            continue
        
        cat  = GEN_category_name(e) or ""
        fam  = GEN_family_name(e) or ""
        typ  = GEN_type_name(e) or ""
        desc = GEN_type_desc(e) or ""
        pref = GEN_type_param_text(e, "MAN_FamilyTypePrefix") or ""

        fam_k  = GEN_norm_text_strong(fam)
        typ_k  = GEN_norm_text_strong(typ)

        key = (fam_k, typ_k)
        if key not in groups:
            groups[key] = [cat, fam_k, typ_k, desc, pref]
        else:
            if not groups[key][3]:
                groups[key][3] = desc
            if not groups[key][4]:
                groups[key][4] = pref

    rows_tmp = []
    for (fam_k, typ_k), vals in groups.items():
        rows_tmp.append((fam_k, typ_k, vals))
    rows_tmp.sort(key=lambda x: (x[0] or "", x[1] or ""))

    ordered = []
    current_keys = set()
    for fam_k, typ_k, vals in rows_tmp:
        ordered.append(vals)
        current_keys.add((fam_k, typ_k))

    sheet = None
    try:
        sheet = GEN_get_sheet_or_create(workbook, GEN_SHEET_NAME)
        headers = GEN_ensure_headers(sheet)
        existing, region = GEN_build_existing_index(sheet, headers)

        updates = []; appends = []
        for cat, fam, typ, desc, pref in ordered:
            key = (GEN_norm_text_strong(fam), GEN_norm_text_strong(typ))
            if key in existing:
                updates.append((existing[key], [cat, fam, typ, desc, pref]))
            else:
                appends.append([cat, fam, typ, desc, pref])

        rows_to_delete = []; removed_keys = []
        for key, row in existing.items():
            if key not in current_keys:
                rows_to_delete.append(row); removed_keys.append(key)

        if updates:
            GEN_write_updates_batched(sheet, headers, updates)

        removed_count = GEN_delete_rows_batched(sheet, rows_to_delete) if rows_to_delete else 0

        added_count = 0
        if appends:
            start_row = GEN_first_empty_row_after_region(region)
            added_count = GEN_write_appends(sheet, start_row, headers, appends)

        GEN_sort_data_region(sheet, headers)

        # --- LOG pulito (compatibile IronPython) ---
        print("[GEN] Aggiunte: {}".format(added_count))
        if added_count > 0 and appends:
            preview = [(r[1], r[2]) for r in appends[:20]]
            if preview:
                print("[GEN] Aggiunte (prime {}): {}".format(len(preview), preview))

        print("[GEN] Eliminate: {}".format(removed_count))
        if removed_count > 0 and removed_keys:
            preview_del = removed_keys[:20]
            if preview_del:
                print("[GEN] Eliminate (prime {}): {}".format(len(preview_del), preview_del))
    finally:
        try:
            if sheet: Marshal.ReleaseComObject(sheet)
        except:
            pass


# ============================================================
# ============ BLOCCO 2 — CAVIDOTTI (Conduits) =================
# ============================================================
COND_SHEET_NAME = "Cavidotti"
COND_HEADER_ROW = 3
COND_MIN_START_DATA_ROW = 5
COND_HEADERS = [
    "Category",
    "Type Name",
    "MAN_TypeDescription_IT",
    "MAN_FamilyTypePrefix",
    "Outside Diameter",   # mm (numerico)
]
COND_EMPTY_RUN_STOP = 20
COND_KEY_MM_PREC = 6

# ----------------------- Utils base -------------------------
def COND_u(s):
    if s is None: return u""
    try:
        return s if isinstance(s, unicode) else unicode(s)
    except:
        return unicode(str(s))

def COND_norm_text(s):
    return COND_u(s).strip()

def COND_norm_text_strong(s):
    return u" ".join(COND_norm_text(s).split())

def COND_category_name(elem):
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

def COND_type_name(elem):
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

_COND_type_desc_cache = {}
def COND_type_desc(elem):
    try:
        p = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
        if not p: return ""
        tid = p.AsElementId()
        if not tid or tid.IntegerValue <= 0: return ""
        tid_i = tid.IntegerValue
        if tid_i in _COND_type_desc_cache: return _COND_type_desc_cache[tid_i]
        t = doc.GetElement(tid)
        val = ""
        if t:
            pp = t.LookupParameter("MAN_TypeDescription_IT")
            if pp: val = (pp.AsString() or "") or (pp.AsValueString() or "")
        _COND_type_desc_cache[tid_i] = val
        return val
    except:
        return ""

def COND_type_param_text(elem, param_name):
    try:
        p = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
        if not p: return ""
        tid = p.AsElementId()
        if not tid or tid.IntegerValue <= 0: return ""
        t = doc.GetElement(tid)
        if not t: return ""
        q = t.LookupParameter(param_name)
        if not q: return ""
        s = q.AsString()
        if s: return COND_norm_text(s)
        vs = q.AsValueString()
        if vs: return COND_norm_text(vs)
        return ""
    except:
        return ""

# --------------------- Unità / diametro ---------------------
def COND_feet_to_mm(val_ft):
    try:
        if _HAS_UTID:
            return float(UnitUtils.ConvertFromInternalUnits(float(val_ft), UnitTypeId.Millimeters))
        else:
            return float(UnitUtils.ConvertFromInternalUnits(float(val_ft), DisplayUnitType.DUT_MILLIMETERS))
    except:
        try: return float(val_ft) * 304.8
        except: return 0.0

def COND_to_float_mm(x):
    if x is None or x == "": return 0.0
    try: return float(x)
    except:
        try: return float(COND_u(x).replace(",", ".").strip())
        except:
            m = re.search(r'(\d+(?:\.\d+)?)', COND_u(x).replace(",", "."))
            return float(m.group(1)) if m else 0.0

def COND_outside_diam_mm_key(elem):
    """Ritorna (key_mm, val_mm) per Outside Diameter del Conduit."""
    p = None
    try:
        p = elem.get_Parameter(BuiltInParameter.RBS_CONDUIT_OUTER_DIAM_PARAM)
    except:
        p = None
    if p:
        # double interno (feet)
        try:
            ft = p.AsDouble()
            if ft and ft > 0:
                mm = COND_feet_to_mm(ft)
                return (round(mm, COND_KEY_MM_PREC), mm)
        except: pass
        # fallback string/value string
        try:
            s = (p.AsString() or p.AsValueString() or "").strip()
            if s:
                mm = COND_to_float_mm(s)
                if mm > 0: return (round(mm, COND_KEY_MM_PREC), mm)
        except: pass
    return (0.0, 0.0)

# ---------------------- Excel helpers -----------------------
def COND_get_sheet_or_create(workbook, name):
    try:
        return workbook.Worksheets.Item[name]
    except:
        sh = workbook.Worksheets.Add()
        sh.Name = name
        return sh

def COND_ensure_headers(sheet):
    last_col = sheet.Cells(COND_HEADER_ROW, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
    if last_col < 1: last_col = 1
    headers = {}
    for c in range(1, last_col + 1):
        v = sheet.Cells(COND_HEADER_ROW, c).Value2
        try:
            nm = v.strip() if isinstance(v, String) else None
        except:
            nm = None
        if nm: headers[nm] = c
    next_col = last_col + 1
    for h in COND_HEADERS:
        if h not in headers:
            sheet.Cells(COND_HEADER_ROW, next_col).Value2 = h
            headers[h] = next_col
            next_col += 1
    return headers

def COND_read_column_block(sheet, col, r0, r1):
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
            out.append(COND_u(val).strip())
        return out
    if isinstance(data, (tuple, list)):
        for row in data:
            try: val = row[0]
            except: val = row
            out.append(COND_u(val).strip())
        return out
    out.append(COND_u(data).strip())
    return out

def COND_detect_data_region(sheet, headers):
    typ_col = headers["Type Name"]
    d_col   = headers["Outside Diameter"]
    r0 = COND_MIN_START_DATA_ROW
    step = 2000
    curr_start = r0
    last_data_row = r0 - 1
    empty_run = 0
    max_rows = sheet.Rows.Count
    while curr_start <= max_rows:
        r1_try = min(max_rows, curr_start + step - 1)
        t_vals = COND_read_column_block(sheet, typ_col, curr_start, r1_try)
        d_vals = COND_read_column_block(sheet, d_col,   curr_start, r1_try)
        block_len = max(len(t_vals), len(d_vals))
        for i in range(block_len):
            t = t_vals[i] if i < len(t_vals) else ""
            d = d_vals[i] if i < len(d_vals) else ""
            if (t or d):
                last_data_row = curr_start + i
                empty_run = 0
            else:
                empty_run += 1
                if empty_run >= COND_EMPTY_RUN_STOP:
                    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)
        curr_start = r1_try + 1
    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)

def COND_build_existing_index(sheet, headers):
    (r0, r1) = COND_detect_data_region(sheet, headers)
    if r1 < r0:
        return {}, (r0, r1)
    typ_col = headers["Type Name"]; d_col = headers["Outside Diameter"]
    col_t = COND_read_column_block(sheet, typ_col, r0, r1)
    col_d = COND_read_column_block(sheet, d_col,   r0, r1)
    index = {}
    n = max(len(col_t), len(col_d))
    for i in range(n):
        t = COND_norm_text_strong(col_t[i] if i < len(col_t) else "")
        d_raw = col_d[i] if i < len(col_d) else ""
        d_val = round(COND_to_float_mm(d_raw), COND_KEY_MM_PREC)
        if t or (d_val != 0.0):
            index[(t, d_val)] = r0 + i
    return index, (r0, r1)

def COND_first_empty_row_after_region(region_tuple):
    r0, r1 = region_tuple
    if r1 < r0: return COND_MIN_START_DATA_ROW
    return max(COND_MIN_START_DATA_ROW, r1 + 1)

def COND_chunk_consecutive_rows(sorted_rows):
    runs = []
    if not sorted_rows: return runs
    start_r = prev_r = sorted_rows[0]
    for r in sorted_rows[1:]:
        if r == prev_r + 1:
            prev_r = r
        else:
            runs.append((start_r, prev_r))
            start_r = prev_r = r
    runs.append((start_r, prev_r))
    return runs

def COND_delete_rows_batched(sheet, rows_to_delete):
    if not rows_to_delete: return 0
    rows_sorted = sorted(rows_to_delete, reverse=True)
    runs = COND_chunk_consecutive_rows(rows_sorted)
    count = 0
    for (r_start, r_end) in runs:
        rng = sheet.Range[sheet.Rows[r_start], sheet.Rows[r_end]]
        rng.EntireRow.Delete()
        count += (r_start - r_end + 1)
    return count

def COND_write_updates_batched(sheet, headers, updates):
    if not updates: return 0
    updates_sorted = sorted(updates, key=lambda x: x[0])
    runs = []
    start = None; buf = []
    for row, vals in updates_sorted:
        if start is None:
            start = row; prev = row; buf = [vals]
        elif row == prev + 1:
            buf.append(vals); prev = row
        else:
            runs.append((start, prev, buf))
            start = row; prev = row; buf = [vals]
    if start is not None:
        runs.append((start, prev, buf))

    col_map = {
        "Category": headers["Category"],
        "Type Name": headers["Type Name"],
        "MAN_TypeDescription_IT": headers["MAN_TypeDescription_IT"],
        "MAN_FamilyTypePrefix": headers["MAN_FamilyTypePrefix"],
        "Outside Diameter": headers["Outside Diameter"],
    }
    keys = ["Category","Type Name","MAN_TypeDescription_IT","MAN_FamilyTypePrefix","Outside Diameter"]
    for (r0, r1, vals_list) in runs:
        n = (r1 - r0 + 1)
        for j, key in enumerate(keys):
            col = col_map[key]
            rng = sheet.Range[sheet.Cells(r0, col), sheet.Cells(r1, col)]
            data = Array.CreateInstance(Object, n, 1)
            for i in range(n):
                v = vals_list[i][j]
                if key == "Outside Diameter":
                    v = float(COND_to_float_mm(v))
                data[i,0] = v
            rng.Value2 = data
    return len(updates)

def COND_write_appends(sheet, start_row, headers, rows_data):
    if not rows_data: return 0
    if start_row < COND_MIN_START_DATA_ROW:
        start_row = COND_MIN_START_DATA_ROW
    cols = [headers[h] for h in COND_HEADERS]
    n_rows = len(rows_data); n_cols = len(cols)
    data = Array.CreateInstance(Object, n_rows, n_cols)
    for i in range(n_rows):
        r = rows_data[i]
        for j in range(n_cols):
            v = r[j]
            if COND_HEADERS[j] == "Outside Diameter":
                v = float(COND_to_float_mm(v))
            data[i, j] = v
    first_cell = sheet.Cells(start_row, min(cols))
    last_cell  = sheet.Cells(start_row + n_rows - 1, max(cols))
    dest = sheet.Range[first_cell, last_cell]
    expected_span = max(cols) - min(cols) + 1
    if expected_span == n_cols:
        dest.Value2 = data
    else:
        for idx, col in enumerate(cols):
            col_rng = sheet.Range[sheet.Cells(start_row, col), sheet.Cells(start_row + n_rows - 1, col)]
            col_data = Array.CreateInstance(Object, n_rows, 1)
            for r in range(n_rows):
                v = rows_data[r][idx]
                if COND_HEADERS[idx] == "Outside Diameter":
                    v = float(COND_to_float_mm(v))
                col_data[r, 0] = v
            col_rng.Value2 = col_data
    return n_rows

def COND_sort_data_region(sheet, headers):
    r0 = COND_MIN_START_DATA_ROW
    first_col = 1
    last_col = sheet.Cells(COND_HEADER_ROW, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

    t_col = headers["Type Name"]
    d_col = headers["Outside Diameter"]

    last_row = 0
    for col in (t_col, d_col):
        lr = sheet.Cells(sheet.Rows.Count, col).End(Excel.XlDirection.xlUp).Row
        if lr > last_row: last_row = lr
    if last_row < r0: return

    data_rng = sheet.Range(sheet.Cells(r0, first_col), sheet.Cells(last_row, last_col))

    sort = sheet.Sort
    sort.SortFields.Clear()
    sort.SortFields.Add(Key=sheet.Range(sheet.Cells(r0, t_col), sheet.Cells(last_row, t_col)),
                        SortOn=Excel.XlSortOn.xlSortOnValues, Order=Excel.XlSortOrder.xlAscending,
                        DataOption=Excel.XlSortDataOption.xlSortNormal)
    sort.SortFields.Add(Key=sheet.Range(sheet.Cells(r0, d_col), sheet.Cells(last_row, d_col)),
                        SortOn=Excel.XlSortOn.xlSortOnValues, Order=Excel.XlSortOrder.xlAscending,
                        DataOption=Excel.XlSortDataOption.xlSortNormal)
    sort.SetRange(data_rng)
    sort.Header = Excel.XlYesNoGuess.xlNo
    sort.MatchCase = False
    sort.Orientation = Excel.XlSortOrientation.xlSortColumns
    sort.Apply()

# -------------------------- RUN -----------------------------
def run_conduits_into_workbook(workbook):
    # Raccogli elementi Conduit (Cavidotti)
    elems = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_Conduit) \
        .WhereElementIsNotElementType() \
        .ToElements()

    groups = {}
    for e in elems:
        # Prefiltraggio: solo Type Name che contengono AirSampling o ThermoCable
        tnm_raw = COND_type_name(e) or ""
        tnm_lc = tnm_raw.lower()
        if ("airsampling" not in tnm_lc) and ("thermocable" not in tnm_lc):
            continue

        cat  = COND_category_name(e) or "Conduits"
        tnm  = tnm_raw
        if not tnm: continue
        desc = COND_type_desc(e) or ""
        pref = COND_type_param_text(e, "MAN_FamilyTypePrefix") or ""

        d_key, d_val = COND_outside_diam_mm_key(e)
        if d_key <= 0:  # ignora senza diametro utile
            continue

        t_key = COND_norm_text_strong(tnm)
        inner = groups.get(t_key)
        if inner is None:
            inner = {}; groups[t_key] = inner
        if d_key not in inner:
            inner[d_key] = [cat, t_key, desc, pref, d_key]
        else:
            if not inner[d_key][2]: inner[d_key][2] = desc
            if not inner[d_key][3]: inner[d_key][3] = pref

    # Ordina per Type Name, poi Outside Diameter (mm)
    rows_tmp = []
    for t_key, by_d in groups.items():
        for d_key, vals in by_d.items():
            rows_tmp.append((t_key, float(d_key), vals))
    rows_tmp.sort(key=lambda x: (x[0] or "", x[1]))

    ordered = []
    current_keys = set()
    for t_key, d_val, vals in rows_tmp:
        ordered.append(vals)
        current_keys.add((COND_norm_text_strong(vals[1]), round(float(vals[4]), COND_KEY_MM_PREC)))

    sheet = None
    try:
        sheet = COND_get_sheet_or_create(workbook, COND_SHEET_NAME)
        headers = COND_ensure_headers(sheet)
        existing, region = COND_build_existing_index(sheet, headers)

        updates = []; appends = []
        for cat, tname, tdesc, pref, d_mm in ordered:
            tkey = COND_norm_text_strong(tname)
            dkey = round(float(d_mm), COND_KEY_MM_PREC)
            key = (tkey, dkey)
            row_vals = [cat, tname, tdesc, pref, dkey]
            if key in existing:
                updates.append((existing[key], row_vals))
            else:
                appends.append(row_vals)

        rows_to_delete = []; removed_keys = []
        for key, row in existing.items():
            if key not in current_keys:
                rows_to_delete.append(row); removed_keys.append(key)

        if updates: COND_write_updates_batched(sheet, headers, updates)
        removed_count = COND_delete_rows_batched(sheet, rows_to_delete) if rows_to_delete else 0

        added_count = 0
        if appends:
            start_row = COND_first_empty_row_after_region(region)
            added_count = COND_write_appends(sheet, start_row, headers, appends)

        COND_sort_data_region(sheet, headers)

        # --- LOG pulito (compatibile IronPython) ---
        print("[CAVIDOTTI] Aggiunte: {}".format(added_count))
        if added_count > 0:
            preview = [(r[1], round(float(r[4]), 3)) for r in appends[:20]]
            if preview:
                print("[CAVIDOTTI] Aggiunte (prime {}): {}".format(len(preview), preview))

        print("[CAVIDOTTI] Eliminate: {}".format(removed_count))
        if removed_count > 0 and removed_keys:
            preview_del = removed_keys[:20]
            print("[CAVIDOTTI] Eliminate (prime {}): {}".format(len(preview_del), preview_del))
    finally:
        try:
            if sheet: Marshal.ReleaseComObject(sheet)
        except: pass


# ============================================================
# ========================= MAIN =============================
# ============================================================
def main():
    # 1) scegli cosa eseguire
    form = RunPickerForm()
    dr = form.ShowDialog()
    if dr != DialogResult.OK:
        return
    run_gen  = form.chkGen.Checked
    run_cond = form.chkCond.Checked

    if not (run_gen or run_cond):
        print("Nessuna opzione selezionata. Operazione annullata.")
        return

    # 2) scegli Excel una sola volta
    excel_path = pick_excel_path_once()
    if not excel_path:
        return

    # 3) apri Excel una volta e lancia i blocchi selezionati
    excel = None; workbook = None
    try:
        excel = Excel.ApplicationClass()
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(excel_path)
        if run_gen:
            run_general_into_workbook(workbook)
        if run_cond:
            run_conduits_into_workbook(workbook)

        # 4) salva & chiudi
        workbook.Save()
        workbook.Close(True)
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
