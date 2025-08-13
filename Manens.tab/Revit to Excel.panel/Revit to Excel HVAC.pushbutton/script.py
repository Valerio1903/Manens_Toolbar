# -*- coding: utf-8 -*-
"""
COMBO (unico Excel) con CHECKBOX UI: Pipe + Pipe Insulations + Pipe Fittings + Mechanical + General -> Excel
Un solo dialog per scegliere il file Excel
Checkbox per scegliere cosa eseguire (Pipe / Insulation / Fittings / Mechanical / General)
Cinque blocchi separati con funzioni rinominate per evitare collisioni
"""

__title__ = 'Revit to Excel\nHVAC'
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
        self.ClientSize = Size(420, 350)

        self.lbl = Label()
        self.lbl.Text = "Scegli le esportazioni da eseguire:"
        self.lbl.Location = Point(16, 16)
        self.lbl.AutoSize = True
        self.lbl.Font = Font(self.Font, FontStyle.Bold)
        self.Controls.Add(self.lbl)

        self.chkPipe = CheckBox()
        self.chkPipe.Text = "Pipe (Type → Diameter)"
        self.chkPipe.Location = Point(20, 50)
        self.chkPipe.AutoSize = True
        self.chkPipe.Checked = True
        self.Controls.Add(self.chkPipe)

        self.chkIns = CheckBox()
        self.chkIns.Text = "Pipe Insulations (Type → Thickness/Size)"
        self.chkIns.Location = Point(20, 78)
        self.chkIns.AutoSize = True
        self.chkIns.Checked = True
        self.Controls.Add(self.chkIns)

        self.chkFit = CheckBox()
        self.chkFit.Text = "Pipe Fittings (Family/Type → MaxSize mm)"
        self.chkFit.Location = Point(20, 106)
        self.chkFit.AutoSize = True
        self.chkFit.Checked = True
        self.Controls.Add(self.chkFit)

        self.chkMeq = CheckBox()
        self.chkMeq.Text = "Apparecchiature Mec (Mechanical Equipment)"
        self.chkMeq.Location = Point(20, 134)
        self.chkMeq.AutoSize = True
        self.chkMeq.Checked = True
        self.Controls.Add(self.chkMeq)

        self.chkGen = CheckBox()
        self.chkGen.Text = "Generale (Pipe Accessories / Plumbing Fixtures / Sprinklers)"
        self.chkGen.Location = Point(20, 162)
        self.chkGen.AutoSize = True
        self.chkGen.Checked = True
        self.Controls.Add(self.chkGen)
        self.chkDuct = CheckBox()
        self.chkDuct.Text = "Canali Rigidi (Ducts – Width/Height o Diameter)"
        self.chkDuct.Location = Point(20, 190)
        self.chkDuct.AutoSize = True
        self.chkDuct.Checked = True
        self.Controls.Add(self.chkDuct)

        self.chkDuctIns = CheckBox()
        self.chkDuctIns.Text = "Isolamento canali (Duct Insulations)"
        self.chkDuctIns.Location = Point(20, 218)
        self.chkDuctIns.AutoSize = True
        self.chkDuctIns.Checked = True
        self.Controls.Add(self.chkDuctIns)

        self.chkDft = CheckBox()
        self.chkDft.Text = "Duct Fittings (Family/Type → MaxSize mm)"
        self.chkDft.Location = Point(20, 246)
        self.chkDft.AutoSize = True
        self.chkDft.Checked = True
        self.Controls.Add(self.chkDft)

        self.chkFxd = CheckBox()
        self.chkFxd.Text = "Canali Flessibili (Type → Diameter)"
        self.chkFxd.Location = Point(20, 274)
        self.chkFxd.AutoSize = True
        self.chkFxd.Checked = True
        self.Controls.Add(self.chkFxd)

        self.btnOk = Button()
        self.btnOk.Text = "OK"
        self.btnOk.Size = Size(100, 28)
        self.btnOk.Location = Point(self.ClientSize.Width - 220, 300)
        self.btnOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btnOk.DialogResult = DialogResult.OK
        self.Controls.Add(self.btnOk)

        self.btnCancel = Button()
        self.btnCancel.Text = "Annulla"
        self.btnCancel.Size = Size(100, 28)
        self.btnCancel.Location = Point(self.ClientSize.Width - 110, 300)
        self.btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btnCancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(self.btnCancel)

        self.AcceptButton = self.btnOk
        self.CancelButton = self.btnCancel


def pick_excel_path_once():
    dlg = OpenFileDialog()
    dlg.Title = "Seleziona il file Excel (unico per PIPE / INSULATION / FITTINGS / MECH / GENERALE / DUCTS / FLEX )"
    dlg.Filter = "Excel (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls"
    dlg.Multiselect = False
    return dlg.FileName if dlg.ShowDialog() == DialogResult.OK else None


# ============================================================
# =============== BLOCCO 1 — PIPE (Type→Diameter) ============
# ============================================================
SHEET_NAME_PIPE = "Tubazioni"
HEADER_ROW_PIPE = 3
MIN_START_DATA_ROW_PIPE = 5
OUR_HEADERS_PIPE = ["Category", "Type Name", "MAN_TypeDescription_IT", "Diameter"]
EMPTY_RUN_STOP_PIPE = 20

def _norm_text_pipe(s):
    if s is None: return ""
    try:
        u = unicode(s) if not isinstance(s, unicode) else s
    except:
        u = str(s)
    return u.strip()

def _norm_diam_key_from_text_pipe(val):
    if val is None: return ""
    try:
        s = unicode(val) if not isinstance(val, unicode) else val
    except:
        s = str(val)
    s = s.strip().replace(",", ".")
    m = re.search(r'(\d+(?:\.\d+)?)', s)
    if not m: return ""
    num = m.group(1)
    if "." in num:
        num = num.rstrip("0").rstrip(".")
    return num

def _to_number_or_text_pipe(s):
    try:
        if s is None: return ""
        ss = unicode(s) if not isinstance(s, unicode) else s
    except:
        ss = str(s)
    ss = ss.strip().replace(",", ".")
    try:
        return float(ss)
    except:
        return ss

def category_name_ui_en_pipe(elem):
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

def type_name_from_instance_pipe(elem):
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

_type_desc_cache_pipe = {}
def man_type_description_it_from_type_pipe(elem):
    try:
        tid_param = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
        if not tid_param: return ""
        tid = tid_param.AsElementId()
        if not tid or tid.IntegerValue <= 0: return ""
        tid_int = tid.IntegerValue
        if tid_int in _type_desc_cache_pipe: return _type_desc_cache_pipe[tid_int]
        t = doc.GetElement(tid)
        val = ""
        if t:
            p = t.LookupParameter("MAN_TypeDescription_IT")
            if p: val = (p.AsString() or "") or (p.AsValueString() or "")
        _type_desc_cache_pipe[tid_int] = val
        return val
    except:
        return ""

def as_string_or_valuestring_pipe(p):
    if not p: return ""
    try:
        s = p.AsString()
        if s: return s
    except: pass
    try:
        vs = p.AsValueString()
        if vs: return vs
    except: pass
    return ""

def diameter_key_and_display_pipe(elem):
    raw = None
    try:
        p = elem.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
        if not p: return None, None, raw
        vs = as_string_or_valuestring_pipe(p)
        raw = vs
        if not vs: return None, None, raw
        key = _norm_diam_key_from_text_pipe(vs)
        if key:
            return key, key, raw
    except: pass
    return None, None, raw

def get_sheet_or_create_pipe(workbook, name):
    try:
        return workbook.Worksheets.Item[name]
    except:
        sh = workbook.Worksheets.Add()
        sh.Name = name
        return sh

def ensure_headers_pipe(sheet):
    last_col = sheet.Cells(HEADER_ROW_PIPE, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
    if last_col < 1: last_col = 1
    headers = {}
    for c in range(1, last_col + 1):
        v = sheet.Cells(HEADER_ROW_PIPE, c).Value2
        if isinstance(v, String):
            nm = v.strip()
            if nm: headers[nm] = c
    next_col = last_col + 1
    for h in OUR_HEADERS_PIPE:
        if h not in headers:
            sheet.Cells(HEADER_ROW_PIPE, next_col).Value2 = h
            headers[h] = next_col
            next_col += 1
    return headers

def _read_column_block_pipe(sheet, col, r0, r1):
    if r1 < r0: return []
    rng = sheet.Range[sheet.Cells(r0, col), sheet.Cells(r1, col)]
    data = rng.Value2
    out = []
    if isinstance(data, System.Array) and data.Rank == 2:
        n0 = data.GetLength(0)
        for i in range(n0):
            try:
                val = data.GetValue(i, 0)
            except:
                try:
                    val = data.GetValue(i+1, 1)
                except:
                    val = None
            out.append(_norm_text_pipe(val))
        return out
    if isinstance(data, (tuple, list)):
        for row in data:
            try: val = row[0]
            except: val = row
            out.append(_norm_text_pipe(val))
        return out
    out.append(_norm_text_pipe(data))
    return out

def detect_data_region_pipe(sheet, headers):
    type_col = headers["Type Name"]
    diam_col = headers["Diameter"]
    r0 = MIN_START_DATA_ROW_PIPE
    step = 2000
    curr_start = r0
    last_data_row = r0 - 1
    empty_run = 0
    max_rows = sheet.Rows.Count
    while curr_start <= max_rows:
        r1_try = min(max_rows, curr_start + step - 1)
        t_vals = _read_column_block_pipe(sheet, type_col, curr_start, r1_try)
        d_vals = _read_column_block_pipe(sheet, diam_col, curr_start, r1_try)
        block_len = max(len(t_vals), len(d_vals))
        for i in range(block_len):
            t = t_vals[i] if i < len(t_vals) else ""
            d = d_vals[i] if i < len(d_vals) else ""
            if (t or d):
                last_data_row = curr_start + i
                empty_run = 0
            else:
                empty_run += 1
                if empty_run >= EMPTY_RUN_STOP_PIPE:
                    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)
        curr_start = r1_try + 1
    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)

def build_existing_index_bulk_pipe(sheet, headers):
    (r0, r1) = detect_data_region_pipe(sheet, headers)
    if r1 < r0:
        return {}, (r0, r1)
    type_col = headers["Type Name"]
    diam_col = headers["Diameter"]
    col_type_vals = _read_column_block_pipe(sheet, type_col, r0, r1)
    col_diam_vals = _read_column_block_pipe(sheet, diam_col, r0, r1)
    index = {}
    n = max(len(col_type_vals), len(col_diam_vals))
    for i in range(n):
        t = col_type_vals[i] if i < len(col_type_vals) else ""
        d_raw = col_diam_vals[i] if i < len(col_diam_vals) else ""
        d = _norm_diam_key_from_text_pipe(d_raw)
        if t or d:
            index[(t, d)] = r0 + i
    return index, (r0, r1)

def first_empty_row_after_region_pipe(region_tuple):
    r0, r1 = region_tuple
    if r1 < r0: return MIN_START_DATA_ROW_PIPE
    return r1 + 1

def _chunk_consecutive_rows_pipe(sorted_rows):
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

def delete_rows_batched_pipe(sheet, rows_to_delete):
    if not rows_to_delete: return 0
    rows_sorted = sorted(rows_to_delete, reverse=True)
    runs = _chunk_consecutive_rows_pipe(rows_sorted)
    count = 0
    for (r_start, r_end) in runs:
        rng = sheet.Range[sheet.Rows[r_start], sheet.Rows[r_end]]
        rng.EntireRow.Delete()
        count += (r_start - r_end + 1)
    return count

def write_updates_batched_pipe(sheet, headers, updates):
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
        "Diameter": headers["Diameter"],
    }
    keys = ["Category","Type Name","MAN_TypeDescription_IT","Diameter"]
    for (r0, r1, vals_list) in runs:
        n = (r1 - r0 + 1)
        for j, key in enumerate(keys):
            col = col_map[key]
            rng = sheet.Range[sheet.Cells(r0, col), sheet.Cells(r1, col)]
            data = Array.CreateInstance(Object, n, 1)
            for i in range(n):
                val = vals_list[i][j]
                if key == "Diameter":
                    val = _to_number_or_text_pipe(val)
                data[i,0] = val
            rng.Value2 = data
    return len(updates)

def write_appends_pipe(sheet, start_row, headers, rows_data):
    if not rows_data: return 0
    cols = [headers[h] for h in OUR_HEADERS_PIPE]
    n_rows = len(rows_data); n_cols = len(cols)
    data = Array.CreateInstance(Object, n_rows, n_cols)
    for i in range(n_rows):
        r = rows_data[i]
        for j in range(n_cols):
            val = r[j]
            if OUR_HEADERS_PIPE[j] == "Diameter":
                val = _to_number_or_text_pipe(val)
            data[i, j] = val
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
                val = rows_data[r][idx]
                if OUR_HEADERS_PIPE[idx] == "Diameter":
                    val = _to_number_or_text_pipe(val)
                col_data[r, 0] = val
            col_rng.Value2 = col_data
    return n_rows

def sort_data_region_pipe(sheet, headers):
    r0 = MIN_START_DATA_ROW_PIPE
    first_col = 1
    last_col = sheet.Cells(HEADER_ROW_PIPE, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

    t_col = headers["Type Name"]
    d_col = headers["Diameter"]

    # ultima riga usata tra le colonne chiave
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

def run_pipe_into_workbook(workbook):
    pipes = FilteredElementCollector(doc).OfClass(Pipe).WhereElementIsNotElementType().ToElements()
    groups = {}
    for p in pipes:
        tname = type_name_from_instance_pipe(p) or ""
        dkey, ddisp, _ = diameter_key_and_display_pipe(p)
        if not dkey: continue
        inner = groups.get(tname)
        if inner is None:
            inner = {}; groups[tname] = inner
        if dkey not in inner:
            inner[dkey] = {
                "category":  category_name_ui_en_pipe(p) or "PipeCurves",
                "type_name": tname,
                "type_desc": man_type_description_it_from_type_pipe(p) or "",
                "diam_disp": ddisp,
            }
        else:
            g = inner[dkey]
            if not g["type_desc"]:
                g["type_desc"] = man_type_description_it_from_type_pipe(p) or g["type_desc"]

    def sort_key_type_pipe(t): return t or ""
    def sort_key_d_pipe(d):
        try: return (0, float(d))
        except: return (1, d)

    ordered = []
    current_keys = set()
    for t in sorted(groups.keys(), key=sort_key_type_pipe):
        inner = groups[t]
        for d in sorted(inner.keys(), key=sort_key_d_pipe):
            g = inner[d]
            ordered.append([g["category"], g["type_name"], g["type_desc"], g["diam_disp"]])
            current_keys.add((_norm_text_pipe(g["type_name"]), _norm_diam_key_from_text_pipe(g["diam_disp"])))

    sheet = None
    try:
        sheet = get_sheet_or_create_pipe(workbook, SHEET_NAME_PIPE)
        headers = ensure_headers_pipe(sheet)
        existing, region = build_existing_index_bulk_pipe(sheet, headers)

        updates = []; appends = []
        for cat, tname, tdesc, diam in ordered:
            tkey = _norm_text_pipe(tname)
            dkey = _norm_diam_key_from_text_pipe(diam)
            if (tkey, dkey) in existing:
                updates.append((existing[(tkey, dkey)], [cat, tname, tdesc, diam]))
            else:
                appends.append([cat, tname, tdesc, diam])

        rows_to_delete = []; removed_keys = []
        for key, row in existing.items():
            if key not in current_keys:
                rows_to_delete.append(row); removed_keys.append(key)

        if updates: write_updates_batched_pipe(sheet, headers, updates)
        removed_count = delete_rows_batched_pipe(sheet, rows_to_delete) if rows_to_delete else 0
        added_count = 0
        if appends:
            start_row = first_empty_row_after_region_pipe(region)
            added_count = write_appends_pipe(sheet, start_row, headers, appends)

        sort_data_region_pipe(sheet, headers)

        print("[PIPE] Aggiunte:", added_count)
        if appends: print("[PIPE] Aggiunte ({}): {}".format(min(20, len(appends)), [(r[1], r[3]) for r in appends[:20]]))
        print("[PIPE] Eliminate:", removed_count)
        if removed_keys: print("[PIPE] Eliminate ({}): {}".format(min(20, len(removed_keys)), removed_keys[:20]))
    finally:
        try:
            if sheet: Marshal.ReleaseComObject(sheet)
        except: pass


# ============================================================
# ========= BLOCCO 2 — PIPE INSULATIONS (solo Pipe host) =====
# ============================================================
SHEET_NAME_INS = "Isolante Tubazioni"
HEADER_ROW_INS = 3
MIN_START_DATA_ROW_INS = 5
OUR_HEADERS_INS = ["Category", "Type Name", "MAN_TypeDescription_IT", "Insulation Thickness", "Pipe Size"]
EMPTY_RUN_STOP_INS = 20
FEET_TO_MM_INS = 304.8

def _norm_text_ins(s):
    if s is None: return ""
    try:
        u = unicode(s) if not isinstance(s, unicode) else s
    except:
        u = str(s)
    return u.strip()

def _fmt_mm_ins(x):
    s = ("%.3f" % float(x)).rstrip("0").rstrip(".")
    return s if s else "0"

def _strip_phi_ins(s):
    if s is None: return ""
    try:
        u = unicode(s) if not isinstance(s, unicode) else s
    except:
        u = str(s)
    u = u.strip()
    u = re.sub(u"[ \t]*(?:[ΦφØø⌀])$", "", u)
    u = re.sub(u"^(?:[ΦφØø⌀])[ \t]*", "", u)
    return u.strip()

def _number_from_text_ins(s):
    if not s: return ""
    ss = _strip_phi_ins(s).replace(",", ".")
    m = re.search(r'(\d+(?:\.\d+)?)', ss)
    if not m: return ""
    num = m.group(1)
    if "." in num: num = num.rstrip("0").rstrip(".")
    return num

def _to_number_or_text_for_thickness_ins(s):
    if s is None: return ""
    try:
        ss = unicode(s) if not isinstance(s, unicode) else s
    except:
        ss = str(s)
    ss = ss.replace(",", ".").strip()
    try:
        return float(ss)
    except:
        return ss

def category_name_ui_en_ins(elem):
    try:
        if elem.Category and elem.Category.Name:
            return elem.Category.Name
    except: pass
    return ""

def type_name_from_instance_ins(elem):
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

_type_desc_cache_ins = {}
def man_type_description_it_from_type_ins(elem):
    try:
        tid_param = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
        if not tid_param: return ""
        tid = tid_param.AsElementId()
        if not tid or tid.IntegerValue <= 0: return ""
        tid_int = tid.IntegerValue
        if tid_int in _type_desc_cache_ins: return _type_desc_cache_ins[tid_int]
        t = doc.GetElement(tid)
        val = ""
        if t:
            p = t.LookupParameter("MAN_TypeDescription_IT")
            if p: val = (p.AsString() or "") or (p.AsValueString() or "")
        _type_desc_cache_ins[tid_int] = val
        return val
    except:
        return ""

def as_string_or_valuestring_ins(p):
    if not p: return ""
    try:
        s = p.AsString()
        if s: return s
    except: pass
    try:
        vs = p.AsValueString()
        if vs: return vs
    except: pass
    return ""

def is_pipe_hosted_ins(ins_elem):
    try:
        hid = ins_elem.HostElementId
        if hid and hid.IntegerValue > 0:
            host = doc.GetElement(hid)
            return isinstance(host, Pipe)
    except: pass
    return False

def thickness_and_size_ins(elem):
    thick_key = ""; thick_disp = ""
    size_key = ""; size_disp = ""
    try:
        pth = elem.get_Parameter(BuiltInParameter.RBS_INSULATION_THICKNESS_FOR_PIPE)
        if pth:
            try:
                d_ft = pth.AsDouble()
                if d_ft is not None:
                    d_mm = d_ft * FEET_TO_MM_INS
                    thick_disp = _fmt_mm_ins(d_mm)
                    thick_key  = thick_disp
            except:
                disp = as_string_or_valuestring_ins(pth)
                n = _number_from_text_ins(disp)
                thick_disp = n
                thick_key  = n
    except: pass
    try:
        psz = elem.get_Parameter(BuiltInParameter.RBS_PIPE_CALCULATED_SIZE)
        if psz:
            raw = as_string_or_valuestring_ins(psz)
            size_disp = _strip_phi_ins(raw)
            size_key  = _number_from_text_ins(size_disp)
    except: pass
    return thick_key, thick_disp, size_key, size_disp

def get_sheet_or_create_ins(workbook, name):
    try:
        return workbook.Worksheets.Item[name]
    except:
        sh = workbook.Worksheets.Add()
        sh.Name = name
        return sh

def ensure_headers_ins(sheet):
    last_col = sheet.Cells(HEADER_ROW_INS, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
    if last_col < 1: last_col = 1
    headers = {}
    for c in range(1, last_col + 1):
        v = sheet.Cells(HEADER_ROW_INS, c).Value2
        if isinstance(v, String):
            nm = v.strip()
            if nm: headers[nm] = c
    next_col = last_col + 1
    for h in OUR_HEADERS_INS:
        if h not in headers:
            sheet.Cells(HEADER_ROW_INS, next_col).Value2 = h
            headers[h] = next_col
            next_col += 1
    return headers

def _read_column_block_ins(sheet, col, r0, r1):
    if r1 < r0: return []
    rng = sheet.Range[sheet.Cells(r0, col), sheet.Cells(r1, col)]
    data = rng.Value2
    out = []
    if isinstance(data, System.Array) and data.Rank == 2:
        n0 = data.GetLength(0)
        for i in range(n0):
            try:
                val = data.GetValue(i, 0)
            except:
                try:
                    val = data.GetValue(i+1, 1)
                except:
                    val = None
            out.append(_norm_text_ins(val))
        return out
    if isinstance(data, (tuple, list)):
        for row in data:
            try: val = row[0]
            except: val = row
            out.append(_norm_text_ins(val))
        return out
    out.append(_norm_text_ins(data))
    return out

def detect_data_region_ins(sheet, headers):
    tn_col = headers["Type Name"]
    th_col = headers["Insulation Thickness"]
    sz_col = headers["Pipe Size"]
    r0 = MIN_START_DATA_ROW_INS
    step = 2000
    curr_start = r0
    last_data_row = r0 - 1
    empty_run = 0
    max_rows = sheet.Rows.Count
    while curr_start <= max_rows:
        r1_try = min(max_rows, curr_start + step - 1)
        t_vals  = _read_column_block_ins(sheet, tn_col, curr_start, r1_try)
        th_vals = _read_column_block_ins(sheet, th_col, curr_start, r1_try)
        sz_vals = _read_column_block_ins(sheet, sz_col, curr_start, r1_try)
        block_len = max(len(t_vals), len(th_vals), len(sz_vals))
        for i in range(block_len):
            t  = t_vals[i]  if i < len(t_vals)  else ""
            th = th_vals[i] if i < len(th_vals) else ""
            sz = sz_vals[i] if i < len(sz_vals) else ""
            if (t or th or sz):
                last_data_row = curr_start + i
                empty_run = 0
            else:
                empty_run += 1
                if empty_run >= EMPTY_RUN_STOP_INS:
                    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)
        curr_start = r1_try + 1
    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)

def build_existing_index_bulk_ins(sheet, headers):
    (r0, r1) = detect_data_region_ins(sheet, headers)
    if r1 < r0:
        return {}, (r0, r1)
    tn_col = headers["Type Name"]; th_col = headers["Insulation Thickness"]; sz_col = headers["Pipe Size"]
    col_tn = _read_column_block_ins(sheet, tn_col, r0, r1)
    col_th = _read_column_block_ins(sheet, th_col, r0, r1)
    col_sz = _read_column_block_ins(sheet, sz_col, r0, r1)
    index = {}
    n = max(len(col_tn), len(col_th), len(col_sz))
    for i in range(n):
        t  = col_tn[i] if i < len(col_tn) else ""
        th = _number_from_text_ins(col_th[i] if i < len(col_th) else "")
        sz = _number_from_text_ins(_strip_phi_ins(col_sz[i] if i < len(col_sz) else ""))
        if t or th or sz:
            index[(t, th, sz)] = r0 + i
    return index, (r0, r1)

def first_empty_row_after_region_ins(region_tuple):
    r0, r1 = region_tuple
    if r1 < r0: return MIN_START_DATA_ROW_INS
    return r1 + 1

def _chunk_consecutive_rows_ins(sorted_rows):
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

def delete_rows_batched_ins(sheet, rows_to_delete):
    if not rows_to_delete: return 0
    rows_sorted = sorted(rows_to_delete, reverse=True)
    runs = _chunk_consecutive_rows_ins(rows_sorted)
    count = 0
    for (r_start, r_end) in runs:
        rng = sheet.Range[sheet.Rows[r_start], sheet.Rows[r_end]]
        rng.EntireRow.Delete()
        count += (r_start - r_end + 1)
    return count

def write_updates_batched_ins(sheet, headers, updates):
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
        "Insulation Thickness": headers["Insulation Thickness"],
        "Pipe Size": headers["Pipe Size"],
    }
    keys = ["Category","Type Name","MAN_TypeDescription_IT","Insulation Thickness","Pipe Size"]
    for (r0, r1, vals_list) in runs:
        n = (r1 - r0 + 1)
        for j, key in enumerate(keys):
            col = col_map[key]
            rng = sheet.Range[sheet.Cells(r0, col), sheet.Cells(r1, col)]
            data = Array.CreateInstance(Object, n, 1)
            for i in range(n):
                val = vals_list[i][j]
                if key == "Insulation Thickness":
                    val = _to_number_or_text_for_thickness_ins(val)
                elif key == "Pipe Size":
                    val = _strip_phi_ins(val)
                data[i,0] = val
            rng.Value2 = data
    return len(updates)

def write_appends_ins(sheet, start_row, headers, rows_data):
    if not rows_data: return 0
    cols = [headers[h] for h in OUR_HEADERS_INS]
    n_rows = len(rows_data); n_cols = len(cols)
    data = Array.CreateInstance(Object, n_rows, n_cols)
    for i in range(n_rows):
        r = rows_data[i]
        for j in range(n_cols):
            val = r[j]; head = OUR_HEADERS_INS[j]
            if head == "Insulation Thickness":
                val = _to_number_or_text_for_thickness_ins(val)
            elif head == "Pipe Size":
                val = _strip_phi_ins(val)
            data[i, j] = val
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
                val = rows_data[r][idx]; head = OUR_HEADERS_INS[idx]
                if head == "Insulation Thickness":
                    val = _to_number_or_text_for_thickness_ins(val)
                elif head == "Pipe Size":
                    val = _strip_phi_ins(val)
                col_data[r, 0] = val
            col_rng.Value2 = col_data
    return n_rows

def sort_data_region_ins(sheet, headers):
    r0 = MIN_START_DATA_ROW_INS
    first_col = 1
    last_col = sheet.Cells(HEADER_ROW_INS, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

    tn_col = headers["Type Name"]
    th_col = headers["Insulation Thickness"]
    sz_col = headers["Pipe Size"]

    last_row = 0
    for col in (tn_col, th_col, sz_col):
        lr = sheet.Cells(sheet.Rows.Count, col).End(Excel.XlDirection.xlUp).Row
        if lr > last_row: last_row = lr
    if last_row < r0: return

    data_rng = sheet.Range(sheet.Cells(r0, first_col), sheet.Cells(last_row, last_col))

    sort = sheet.Sort
    sort.SortFields.Clear()
    sort.SortFields.Add(Key=sheet.Range(sheet.Cells(r0, tn_col), sheet.Cells(last_row, tn_col)),
                        SortOn=Excel.XlSortOn.xlSortOnValues, Order=Excel.XlSortOrder.xlAscending,
                        DataOption=Excel.XlSortDataOption.xlSortNormal)
    sort.SortFields.Add(Key=sheet.Range(sheet.Cells(r0, th_col), sheet.Cells(last_row, th_col)),
                        SortOn=Excel.XlSortOn.xlSortOnValues, Order=Excel.XlSortOrder.xlAscending,
                        DataOption=Excel.XlSortDataOption.xlSortNormal)
    sort.SortFields.Add(Key=sheet.Range(sheet.Cells(r0, sz_col), sheet.Cells(last_row, sz_col)),
                        SortOn=Excel.XlSortOn.xlSortOnValues, Order=Excel.XlSortOrder.xlAscending,
                        DataOption=Excel.XlSortDataOption.xlSortNormal)
    sort.SetRange(data_rng)
    sort.Header = Excel.XlYesNoGuess.xlNo
    sort.MatchCase = False
    sort.Orientation = Excel.XlSortOrientation.xlSortColumns
    sort.Apply()

def run_ins_into_workbook(workbook):
    insulations = FilteredElementCollector(doc).OfClass(PipeInsulation).WhereElementIsNotElementType().ToElements()
    groups = {}
    for ins in insulations:
        if not is_pipe_hosted_ins(ins): continue
        tname = type_name_from_instance_ins(ins) or ""
        th_key, th_disp, sz_key, sz_disp = thickness_and_size_ins(ins)
        if not th_key and not sz_key: continue
        inner = groups.get(tname)
        if inner is None:
            inner = {}; groups[tname] = inner
        pair = (th_key or "0", sz_key or "")
        if pair not in inner:
            inner[pair] = {
                "category":  category_name_ui_en_ins(ins) or "Pipe Insulations",
                "type_name": tname,
                "type_desc": man_type_description_it_from_type_ins(ins) or "",
                "thick_disp": th_disp or "",
                "size_disp":  sz_disp or "",
            }
        else:
            g = inner[pair]
            if not g["type_desc"]:
                g["type_desc"] = man_type_description_it_from_type_ins(ins) or g["type_desc"]

    rows_tmp = []
    for tname, inner in groups.items():
        for (th_key, sz_key), g in inner.items():
            try: th_num = float(th_key) if th_key not in ("", None) else 0.0
            except: th_num = 0.0
            try: sz_num = float(sz_key) if sz_key not in ("", None) else 0.0
            except: sz_num = 0.0
            rows_tmp.append((_norm_text_ins(g["type_name"]), th_num, sz_num,
                             [g["category"], g["type_name"], g["type_desc"], g["thick_disp"], g["size_disp"]]))
    rows_tmp.sort(key=lambda x: (x[0], x[1], x[2]))

    ordered = []
    current_keys = set()
    for _, _, _, rowvals in rows_tmp:
        ordered.append(rowvals)
        current_keys.add((_norm_text_ins(rowvals[1]),
                          _number_from_text_ins(rowvals[3]) or "0",
                          _number_from_text_ins(rowvals[4])))

    sheet = None
    try:
        sheet = get_sheet_or_create_ins(workbook, SHEET_NAME_INS)
        headers = ensure_headers_ins(sheet)
        existing, region = build_existing_index_bulk_ins(sheet, headers)

        updates = []; appends = []
        for cat, tname, tdesc, thick, size in ordered:
            tkey  = _norm_text_ins(tname)
            thkey = _number_from_text_ins(thick) or "0"
            szkey = _number_from_text_ins(size)
            key = (tkey, thkey, szkey)
            if key in existing:
                updates.append((existing[key], [cat, tname, tdesc, thick, _strip_phi_ins(size)]))
            else:
                appends.append([cat, tname, tdesc, thick, _strip_phi_ins(size)])

        rows_to_delete = []; removed_keys = []
        for key, row in existing.items():
            if key not in current_keys:
                rows_to_delete.append(row); removed_keys.append(key)

        if updates: write_updates_batched_ins(sheet, headers, updates)
        removed_count = delete_rows_batched_ins(sheet, rows_to_delete) if rows_to_delete else 0
        added_count = 0
        if appends:
            start_row = first_empty_row_after_region_ins(region)
            added_count = write_appends_ins(sheet, start_row, headers, appends)

        sort_data_region_ins(sheet, headers)

        print("[INS] Aggiunte:", added_count)
        if appends: print("[INS] Aggiunte ({}): {}".format(min(20, len(appends)), [(r[1], r[3], r[4]) for r in appends[:20]]))
        print("[INS] Eliminate:", removed_count)
        if removed_keys: print("[INS] Eliminate ({}): {}".format(min(20, len(removed_keys)), removed_keys[:20]))
    finally:
        try:
            if sheet: Marshal.ReleaseComObject(sheet)
        except: pass


# ============================================================
# ====== BLOCCO 3 — PIPE FITTINGS (Family/Type → MaxSize) ====
# ============================================================
SHEET_NAME_FIT = "Raccordi Tubi"
HEADER_ROW_FIT = 3
MIN_START_DATA_ROW_FIT = 5
OUR_HEADERS_FIT = ["Category", "Family Name", "Type Name", "MAN_TypeDescription_IT", "MAN_Fittings_MaxSize"]
EMPTY_RUN_STOP_FIT = 20
KEY_MM_PREC_FIT = 6

def _u_fit(s):
    if s is None: return u""
    try: return s if isinstance(s, unicode) else unicode(s)
    except: return unicode(str(s))

def _norm_text_fit(s):
    return _u_fit(s).strip()

def _norm_text_strong_fit(s):
    return u" ".join(_norm_text_fit(s).split())

def _to_float_fit(x):
    if x is None or x == "": return 0.0
    try: return float(x)
    except:
        try: return float(_u_fit(x).replace(",", ".").strip())
        except: return 0.0

def _norm_mm_key_fit(v):
    try: return round(float(v), KEY_MM_PREC_FIT)
    except: return 0.0

def _category_name_fit(elem):
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

def _type_name_fit(elem):
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

def _family_name_fit(elem):
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

_type_desc_cache_fit = {}
def _man_type_description_it_fit(elem):
    try:
        p = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
        if not p: return ""
        tid = p.AsElementId()
        if not tid or tid.IntegerValue <= 0: return ""
        tid_i = tid.IntegerValue
        if tid_i in _type_desc_cache_fit: return _type_desc_cache_fit[tid_i]
        t = doc.GetElement(tid)
        val = ""
        if t:
            pp = t.LookupParameter("MAN_TypeDescription_IT")
            if pp: val = (pp.AsString() or "") or (pp.AsValueString() or "")
        _type_desc_cache_fit[tid_i] = val
        return val
    except:
        return ""

def _feet_to_mm_fit(val_ft):
    try:
        if _HAS_UTID:
            return float(UnitUtils.ConvertFromInternalUnits(float(val_ft), UnitTypeId.Millimeters))
        else:
            return float(UnitUtils.ConvertFromInternalUnits(float(val_ft), DisplayUnitType.DUT_MILLIMETERS))
    except:
        try: return float(val_ft) * 304.8
        except: return 0.0

def _maxsize_mm_fit(elem):
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
        return _feet_to_mm_fit(d_ft)
    except:
        return 0.0

def get_sheet_or_create_fit(workbook, name):
    try:
        return workbook.Worksheets.Item[name]
    except:
        sh = workbook.Worksheets.Add()
        sh.Name = name
        return sh

def ensure_headers_fit(sheet):
    last_col = sheet.Cells(HEADER_ROW_FIT, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
    if last_col < 1: last_col = 1
    headers = {}
    for c in range(1, last_col + 1):
        v = sheet.Cells(HEADER_ROW_FIT, c).Value2
        try:
            nm = v.strip() if isinstance(v, String) else None
        except:
            nm = None
        if nm: headers[nm] = c
    next_col = last_col + 1
    for h in OUR_HEADERS_FIT:
        if h not in headers:
            sheet.Cells(HEADER_ROW_FIT, next_col).Value2 = h
            headers[h] = next_col
            next_col += 1
    return headers

def _read_column_block_fit(sheet, col, r0, r1):
    if r1 < r0: return []
    rng = sheet.Range[sheet.Cells(r0, col), sheet.Cells(r1, col)]
    data = rng.Value2
    out = []
    if isinstance(data, System.Array) and data.Rank == 2:
        n0 = data.GetLength(0)
        for i in range(n0):
            try:
                val = data.GetValue(i, 0)
            except:
                try:
                    val = data.GetValue(i+1, 1)
                except:
                    val = None
            out.append(_u_fit(val).strip())
        return out
    if isinstance(data, (tuple, list)):
        for row in data:
            try: val = row[0]
            except: val = row
            out.append(_u_fit(val).strip())
        return out
    out.append(_u_fit(data).strip())
    return out

def detect_data_region_fit(sheet, headers):
    fam_col = headers["Family Name"]
    typ_col = headers["Type Name"]
    max_col = headers["MAN_Fittings_MaxSize"]
    r0 = MIN_START_DATA_ROW_FIT
    step = 2000
    curr_start = r0
    last_data_row = r0 - 1
    empty_run = 0
    max_rows = sheet.Rows.Count
    while curr_start <= max_rows:
        r1_try = min(max_rows, curr_start + step - 1)
        f_vals = _read_column_block_fit(sheet, fam_col, curr_start, r1_try)
        t_vals = _read_column_block_fit(sheet, typ_col, curr_start, r1_try)
        m_vals = _read_column_block_fit(sheet, max_col, curr_start, r1_try)
        block_len = max(len(f_vals), len(t_vals), len(m_vals))
        for i in range(block_len):
            f = f_vals[i] if i < len(f_vals) else ""
            t = t_vals[i] if i < len(t_vals) else ""
            m = m_vals[i] if i < len(m_vals) else ""
            if (f or t or m):
                last_data_row = curr_start + i
                empty_run = 0
            else:
                empty_run += 1
                if empty_run >= EMPTY_RUN_STOP_FIT:
                    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)
        curr_start = r1_try + 1
    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)

def build_existing_index_bulk_fit(sheet, headers):
    (r0, r1) = detect_data_region_fit(sheet, headers)
    if r1 < r0:
        return {}, (r0, r1)
    fam_col = headers["Family Name"]; typ_col = headers["Type Name"]; max_col = headers["MAN_Fittings_MaxSize"]
    col_f = _read_column_block_fit(sheet, fam_col, r0, r1)
    col_t = _read_column_block_fit(sheet, typ_col, r0, r1)
    col_m = _read_column_block_fit(sheet, max_col, r0, r1)
    index = {}
    n = max(len(col_f), len(col_t), len(col_m))
    for i in range(n):
        f = _norm_text_strong_fit(col_f[i] if i < len(col_f) else "")
        t = _norm_text_strong_fit(col_t[i] if i < len(col_t) else "")
        m_raw = col_m[i] if i < len(col_m) else ""
        try:
            m_val = float(m_raw) if m_raw != "" else 0.0
        except:
            m_val = _to_float_fit(m_raw)
        m_key = _norm_mm_key_fit(m_val)
        if f or t or (m_key != 0.0):
            index[(f, t, m_key)] = r0 + i
    return index, (r0, r1)

def first_empty_row_after_region_fit(region_tuple):
    r0, r1 = region_tuple
    if r1 < r0: return MIN_START_DATA_ROW_FIT
    return r1 + 1

def _chunk_consecutive_rows_fit(sorted_rows):
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

def delete_rows_batched_fit(sheet, rows_to_delete):
    if not rows_to_delete: return 0
    rows_sorted = sorted(rows_to_delete, reverse=True)
    runs = _chunk_consecutive_rows_fit(rows_sorted)
    count = 0
    for (r_start, r_end) in runs:
        rng = sheet.Range[sheet.Rows[r_start], sheet.Rows[r_end]]
        rng.EntireRow.Delete()
        count += (r_start - r_end + 1)
    return count

def write_updates_batched_fit(sheet, headers, updates):
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
        "MAN_Fittings_MaxSize": headers["MAN_Fittings_MaxSize"],
    }
    keys = ["Category","Family Name","Type Name","MAN_TypeDescription_IT","MAN_Fittings_MaxSize"]
    for (r0, r1, vals_list) in runs:
        n = (r1 - r0 + 1)
        for j, key in enumerate(keys):
            col = col_map[key]
            rng = sheet.Range[sheet.Cells(r0, col), sheet.Cells(r1, col)]
            data = Array.CreateInstance(Object, n, 1)
            for i in range(n):
                val = vals_list[i][j]
                if key == "MAN_Fittings_MaxSize":
                    val = float(_to_float_fit(val))  # numerico in mm
                data[i,0] = val
            rng.Value2 = data
    return len(updates)

def write_appends_batched_fit(sheet, start_row, headers, rows_data):
    if not rows_data: return 0
    if start_row < MIN_START_DATA_ROW_FIT:
        start_row = MIN_START_DATA_ROW_FIT
    cols = [headers[h] for h in OUR_HEADERS_FIT]
    n_rows = len(rows_data); n_cols = len(cols)
    data = Array.CreateInstance(Object, n_rows, n_cols)
    for i in range(n_rows):
        r = rows_data[i]
        for j in range(n_cols):
            val = r[j]
            if OUR_HEADERS_FIT[j] == "MAN_Fittings_MaxSize":
                val = float(_to_float_fit(val))  # numerico in mm
            data[i, j] = val
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
                val = rows_data[r][idx]
                if OUR_HEADERS_FIT[idx] == "MAN_Fittings_MaxSize":
                    val = float(_to_float_fit(val))
                col_data[r, 0] = val
            col_rng.Value2 = col_data
    return n_rows

def sort_data_region_fit(sheet, headers):
    r0 = MIN_START_DATA_ROW_FIT
    first_col = 1
    last_col = sheet.Cells(HEADER_ROW_FIT, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

    fam_col = headers["Family Name"]
    typ_col = headers["Type Name"]
    max_col = headers["MAN_Fittings_MaxSize"]

    last_row = 0
    for col in (fam_col, typ_col, max_col):
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
    sort.SortFields.Add(Key=sheet.Range(sheet.Cells(r0, max_col), sheet.Cells(last_row, max_col)),
                        SortOn=Excel.XlSortOn.xlSortOnValues, Order=Excel.XlSortOrder.xlAscending,
                        DataOption=Excel.XlSortDataOption.xlSortNormal)
    sort.SetRange(data_rng)
    sort.Header = Excel.XlYesNoGuess.xlNo
    sort.MatchCase = False
    sort.Orientation = Excel.XlSortOrientation.xlSortColumns
    sort.Apply()

def run_fittings_into_workbook(workbook):
    elems = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_PipeFitting)\
        .WhereElementIsNotElementType()\
        .ToElements()

    groups = {}
    for e in elems:
        if not isinstance(e, FamilyInstance):
            continue
        cat  = _category_name_fit(e) or "Pipe Fittings"
        fam  = _family_name_fit(e) or ""
        typ  = _type_name_fit(e) or ""
        desc = _man_type_description_it_fit(e) or ""
        msz_mm = _maxsize_mm_fit(e)

        fam_k = _norm_text_strong_fit(fam)
        typ_k = _norm_text_strong_fit(typ)
        msz_k = _norm_mm_key_fit(msz_mm)

        key = (fam_k, typ_k, msz_k)
        if key not in groups:
            groups[key] = [cat, fam_k, typ_k, desc, msz_k]
        else:
            if not groups[key][3]:
                groups[key][3] = desc

    rows_tmp = []
    for (fam_k, typ_k, msz_k), vals in groups.items():
        rows_tmp.append((fam_k, typ_k, msz_k, vals))
    rows_tmp.sort(key=lambda x: (x[0] or "", x[1] or "", x[2]))

    ordered = []
    current_keys = set()
    for fam_k, typ_k, msz_k, vals in rows_tmp:
        ordered.append(vals)
        current_keys.add((fam_k, typ_k, msz_k))

    sheet = None
    try:
        sheet = get_sheet_or_create_fit(workbook, SHEET_NAME_FIT)
        headers = ensure_headers_fit(sheet)
        existing, region = build_existing_index_bulk_fit(sheet, headers)

        updates = []; appends = []
        for cat, fam, typ, desc, msz_mm in ordered:
            fam_k = _norm_text_strong_fit(fam)
            typ_k = _norm_text_strong_fit(typ)
            msz_k = _norm_mm_key_fit(msz_mm)
            key   = (fam_k, typ_k, msz_k)
            if key in existing:
                updates.append((existing[key], [cat, fam_k, typ_k, desc, msz_k]))
            else:
                appends.append([cat, fam_k, typ_k, desc, msz_k])

        rows_to_delete = []; removed_keys = []
        for key, row in existing.items():
            if key not in current_keys:
                rows_to_delete.append(row); removed_keys.append(key)

        if updates: write_updates_batched_fit(sheet, headers, updates)
        removed_count = delete_rows_batched_fit(sheet, rows_to_delete) if rows_to_delete else 0
        added_count = 0
        if appends:
            start_row = first_empty_row_after_region_fit(region)
            added_count = write_appends_batched_fit(sheet, start_row, headers, appends)

        sort_data_region_fit(sheet, headers)

        print("[FITTINGS] Aggiunte:", added_count)
        if appends:
            preview = [(r[1], r[2], round(float(r[4]), 3)) for r in appends[:20]]
            print("[FITTINGS] Aggiunte: {}".format(preview))
        print("[FITTINGS] Eliminate:", removed_count)
        if removed_keys:
            preview_del = [(k[0], k[1], round(float(k[2]), 3)) for k in removed_keys[:20]]
            print("[FITTINGS] Eliminate: {}".format(preview_del))
    finally:
        try:
            if sheet: Marshal.ReleaseComObject(sheet)
        except: pass


# ============================================================
# === BLOCCO 4 — MECHANICAL EQUIPMENT (Apparecchiature Mec) ==
# ============================================================
MEQ_SHEET_NAME = "Apparecchiature Mec"
MEQ_HEADER_ROW = 3
MEQ_MIN_START_DATA_ROW = 5
MEQ_HEADERS = [
    "Category",
    "Family Name",
    "Type Name",
    "MAN_TypeDescription_IT",
    "MAN_FamilyTypePrefix",
    "MAN_Type_Code",
]
MEQ_EMPTY_RUN_STOP = 20

def MEQ_u(s):
    if s is None: return u""
    try: return s if isinstance(s, unicode) else unicode(s)
    except: return unicode(str(s))

def MEQ_norm_text(s):
    return MEQ_u(s).strip()

def MEQ_norm_text_strong(s):
    return u" ".join(MEQ_norm_text(s).split())

def MEQ_category_name(elem):
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

def MEQ_type_name(elem):
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

def MEQ_family_name(elem):
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

_MEQ_type_desc_cache = {}
def MEQ_type_desc(elem):
    try:
        p = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
        if not p: return ""
        tid = p.AsElementId()
        if not tid or tid.IntegerValue <= 0: return ""
        tid_i = tid.IntegerValue
        if tid_i in _MEQ_type_desc_cache: return _MEQ_type_desc_cache[tid_i]
        t = doc.GetElement(tid)
        val = ""
        if t:
            pp = t.LookupParameter("MAN_TypeDescription_IT")
            if pp: val = (pp.AsString() or "") or (pp.AsValueString() or "")
        _MEQ_type_desc_cache[tid_i] = val
        return val
    except:
        return ""

def MEQ_type_param_text(elem, param_name):
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
        if s: return MEQ_norm_text(s)
        vs = q.AsValueString()
        if vs: return MEQ_norm_text(vs)
        return ""
    except:
        return ""
    
def MEQ_instance_param_text(elem, param_name):
    try:
        q = elem.LookupParameter(param_name)
        if not q: return ""
        s = q.AsString()
        if s: return MEQ_norm_text(s)
        vs = q.AsValueString()
        if vs: return MEQ_norm_text(vs)
        return ""
    except:
        return ""


def MEQ_get_sheet_or_create(workbook, name):
    try:
        return workbook.Worksheets.Item[name]
    except:
        sh = workbook.Worksheets.Add()
        sh.Name = name
        return sh

def MEQ_ensure_headers(sheet):
    last_col = sheet.Cells(MEQ_HEADER_ROW, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
    if last_col < 1: last_col = 1
    headers = {}
    for c in range(1, last_col + 1):
        v = sheet.Cells(MEQ_HEADER_ROW, c).Value2
        try:
            nm = v.strip() if isinstance(v, String) else None
        except:
            nm = None
        if nm: headers[nm] = c
    next_col = last_col + 1
    for h in MEQ_HEADERS:
        if h not in headers:
            sheet.Cells(MEQ_HEADER_ROW, next_col).Value2 = h
            headers[h] = next_col
            next_col += 1
    return headers

def MEQ_read_column_block(sheet, col, r0, r1):
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
            out.append(MEQ_u(val).strip())
        return out
    if isinstance(data, (tuple, list)):
        for row in data:
            try: val = row[0]
            except: val = row
            out.append(MEQ_u(val).strip())
        return out
    out.append(MEQ_u(data).strip())
    return out

def MEQ_detect_data_region(sheet, headers):
    fam_col = headers["Family Name"]
    typ_col = headers["Type Name"]
    code_col = headers["MAN_Type_Code"]
    r0 = MEQ_MIN_START_DATA_ROW
    step = 2000
    curr_start = r0
    last_data_row = r0 - 1
    empty_run = 0
    max_rows = sheet.Rows.Count
    while curr_start <= max_rows:
        r1_try = min(max_rows, curr_start + step - 1)
        f_vals = MEQ_read_column_block(sheet, fam_col, curr_start, r1_try)
        t_vals = MEQ_read_column_block(sheet, typ_col, curr_start, r1_try)
        c_vals = MEQ_read_column_block(sheet, code_col, curr_start, r1_try)
        block_len = max(len(f_vals), len(t_vals), len(c_vals))
        for i in range(block_len):
            f = f_vals[i] if i < len(f_vals) else ""
            t = t_vals[i] if i < len(t_vals) else ""
            c = c_vals[i] if i < len(c_vals) else ""
            if (f or t or c):
                last_data_row = curr_start + i
                empty_run = 0
            else:
                empty_run += 1
                if empty_run >= MEQ_EMPTY_RUN_STOP:
                    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)
        curr_start = r1_try + 1
    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)

def MEQ_build_existing_index(sheet, headers):
    (r0, r1) = MEQ_detect_data_region(sheet, headers)
    if r1 < r0:
        return {}, (r0, r1)
    fam_col = headers["Family Name"]; typ_col = headers["Type Name"]; code_col = headers["MAN_Type_Code"]
    col_f = MEQ_read_column_block(sheet, fam_col, r0, r1)
    col_t = MEQ_read_column_block(sheet, typ_col, r0, r1)
    col_c = MEQ_read_column_block(sheet, code_col, r0, r1)
    index = {}
    n = max(len(col_f), len(col_t), len(col_c))
    for i in range(n):
        f = MEQ_norm_text_strong(col_f[i] if i < len(col_f) else "")
        t = MEQ_norm_text_strong(col_t[i] if i < len(col_t) else "")
        c = MEQ_norm_text_strong(col_c[i] if i < len(col_c) else "")
        if f or t or c:
            index[(f, t, c)] = r0 + i
    return index, (r0, r1)

def MEQ_first_empty_row_after_region(region_tuple):
    r0, r1 = region_tuple
    if r1 < r0: return MEQ_MIN_START_DATA_ROW
    return max(MEQ_MIN_START_DATA_ROW, r1 + 1)

def MEQ_chunk_consecutive_rows(sorted_rows):
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

def MEQ_delete_rows_batched(sheet, rows_to_delete):
    if not rows_to_delete: return 0
    rows_sorted = sorted(rows_to_delete, reverse=True)
    runs = MEQ_chunk_consecutive_rows(rows_sorted)
    count = 0
    for (r_start, r_end) in runs:
        rng = sheet.Range[sheet.Rows[r_start], sheet.Rows[r_end]]
        rng.EntireRow.Delete()
        count += (r_start - r_end + 1)
    return count

def MEQ_write_updates_batched(sheet, headers, updates):
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
        "MAN_Type_Code": headers["MAN_Type_Code"],
    }
    keys = ["Category","Family Name","Type Name","MAN_TypeDescription_IT","MAN_FamilyTypePrefix","MAN_Type_Code"]
    for (r0, r1, vals_list) in runs:
        n = (r1 - r0 + 1)
        for j, key in enumerate(keys):
            col = col_map[key]
            rng = sheet.Range[sheet.Cells(r0, col), sheet.Cells(r1, col)]
            data = Array.CreateInstance(Object, n, 1)
            for i in range(n):
                data[i,0] = MEQ_u(vals_list[i][j])
            rng.Value2 = data
    return len(updates)

def MEQ_write_appends(sheet, start_row, headers, rows_data):
    if not rows_data: return 0
    if start_row < MEQ_MIN_START_DATA_ROW:
        start_row = MEQ_MIN_START_DATA_ROW
    cols = [headers[h] for h in MEQ_HEADERS]
    n_rows = len(rows_data); n_cols = len(cols)
    data = Array.CreateInstance(Object, n_rows, n_cols)
    for i in range(n_rows):
        r = rows_data[i]
        for j in range(n_cols):
            data[i, j] = MEQ_u(r[j])
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
                col_data[r, 0] = MEQ_u(rows_data[r][idx])
            col_rng.Value2 = col_data
    return n_rows

def MEQ_sort_data_region(sheet, headers):
    r0 = MEQ_MIN_START_DATA_ROW
    first_col = 1
    last_col = sheet.Cells(MEQ_HEADER_ROW, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

    fam_col = headers["Family Name"]
    typ_col = headers["Type Name"]
    code_col = headers["MAN_Type_Code"]

    last_row = 0
    for col in (fam_col, typ_col, code_col):
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
    sort.SortFields.Add(Key=sheet.Range(sheet.Cells(r0, code_col), sheet.Cells(last_row, code_col)),
                        SortOn=Excel.XlSortOn.xlSortOnValues, Order=Excel.XlSortOrder.xlAscending,
                        DataOption=Excel.XlSortDataOption.xlSortNormal)
    sort.SetRange(data_rng)
    sort.Header = Excel.XlYesNoGuess.xlNo
    sort.MatchCase = False
    sort.Orientation = Excel.XlSortOrientation.xlSortColumns
    sort.Apply()

def run_mechanical_equipment_into_workbook(workbook):
    elems = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_MechanicalEquipment)\
        .WhereElementIsNotElementType()\
        .ToElements()

    groups = {}
    for e in elems:
        if not isinstance(e, FamilyInstance):
            continue
        cat  = MEQ_category_name(e) or "Mechanical Equipment"
        fam  = MEQ_family_name(e) or ""
        typ  = MEQ_type_name(e) or ""
        desc = MEQ_type_desc(e) or ""
        pref = MEQ_type_param_text(e, "MAN_FamilyTypePrefix") or ""
        code  = MEQ_instance_param_text(e, "MAN_Type_Code") or ""


        fam_k  = MEQ_norm_text_strong(fam)
        typ_k  = MEQ_norm_text_strong(typ)
        code_k = MEQ_norm_text_strong(code)

        key = (fam_k, typ_k, code_k)
        if key not in groups:
            groups[key] = [cat, fam_k, typ_k, desc, pref, code_k]
        else:
            if not groups[key][3]:
                groups[key][3] = desc
            if not groups[key][4]:
                groups[key][4] = pref

    rows_tmp = []
    for (fam_k, typ_k, code_k), vals in groups.items():
        rows_tmp.append((fam_k, typ_k, code_k, vals))
    rows_tmp.sort(key=lambda x: (x[0] or "", x[1] or "", x[2] or ""))

    ordered = []
    current_keys = set()
    for fam_k, typ_k, code_k, vals in rows_tmp:
        ordered.append(vals)
        current_keys.add((fam_k, typ_k, code_k))

    sheet = None
    try:
        sheet = MEQ_get_sheet_or_create(workbook, MEQ_SHEET_NAME)
        headers = MEQ_ensure_headers(sheet)
        existing, region = MEQ_build_existing_index(sheet, headers)

        updates = []; appends = []
        for cat, fam, typ, desc, pref, code in ordered:
            key = (MEQ_norm_text_strong(fam), MEQ_norm_text_strong(typ), MEQ_norm_text_strong(code))
            if key in existing:
                updates.append((existing[key], [cat, fam, typ, desc, pref, code]))
            else:
                appends.append([cat, fam, typ, desc, pref, code])

        rows_to_delete = []; removed_keys = []
        for key, row in existing.items():
            if key not in current_keys:
                rows_to_delete.append(row); removed_keys.append(key)

        if updates: MEQ_write_updates_batched(sheet, headers, updates)
        removed_count = MEQ_delete_rows_batched(sheet, rows_to_delete) if rows_to_delete else 0
        added_count = 0
        if appends:
            start_row = MEQ_first_empty_row_after_region(region)
            added_count = MEQ_write_appends(sheet, start_row, headers, appends)

        MEQ_sort_data_region(sheet, headers)

        print("[MECH EQ] Aggiunte:", added_count)
        if appends:
            preview = [(r[1], r[2], r[5]) for r in appends[:20]]
            print("[MECH EQ] Aggiunte (prime 20): {}".format(preview))
        print("[MECH EQ] Eliminate:", removed_count)
        if removed_keys:
            print("[MECH EQ] Eliminate (prime 20): {}".format(removed_keys[:20]))
    finally:
        try:
            if sheet: Marshal.ReleaseComObject(sheet)
        except: pass


# ============================================================
# ========== BLOCCO 4 — GENERALE (PA / PF / Sprinklers) ======
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
    return GEN_u(s).strip()

def GEN_norm_text_strong(s):
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
    for bic in (BuiltInCategory.OST_DuctTerminal,
                BuiltInCategory.OST_DuctAccessory,
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_PlumbingFixtures,
                BuiltInCategory.OST_Sprinklers):
        elems.extend(
            list(FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType().ToElements())
        )

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

        if updates: GEN_write_updates_batched(sheet, headers, updates)
        removed_count = GEN_delete_rows_batched(sheet, rows_to_delete) if rows_to_delete else 0

        added_count = 0
        if appends:
            start_row = GEN_first_empty_row_after_region(region)
            added_count = GEN_write_appends(sheet, start_row, headers, appends)

        GEN_sort_data_region(sheet, headers)

        print("[GEN] Aggiunte:", added_count)
        if appends:
            preview = [(r[1], r[2]) for r in appends[:20]]
            print("[GEN] Aggiunte (prime 20): {}".format(preview))
        print("[GEN] Eliminate:", removed_count)
        if removed_keys:
            print("[GEN] Eliminate (prime 20): {}".format(removed_keys[:20]))
    finally:
        try:
            if sheet: Marshal.ReleaseComObject(sheet)
        except: pass

# ============================================================
# ======== BLOCCO 5 — CANALI RIGIDI (Type → Max Dim) =========
# ============================================================
# Sheet: "Canali Rigidi"
# Colonne: Category | Type Name | MAN_TypeDescription_IT | Width/Height - Diameter
SHEET_NAME_DUCT = "Canali Rigidi"
HEADER_ROW_DUCT = 3
MIN_START_DATA_ROW_DUCT = 5
OUR_HEADERS_DUCT = ["Category", "Type Name", "MAN_TypeDescription_IT", "Width/Height - Diameter"]
EMPTY_RUN_STOP_DUCT = 20
KEY_MM_PREC_DUCT = 6

# -------- utils base --------
def _u_duct(s):
    if s is None: return u""
    try: return s if isinstance(s, unicode) else unicode(s)
    except: return unicode(str(s))

def _norm_text_duct(s):
    return _u_duct(s).strip()

def _category_name_duct(elem):
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

def _type_name_duct(elem):
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

_duct_type_desc_cache = {}
def _man_type_description_it_duct(elem):
    try:
        p = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
        if not p: return ""
        tid = p.AsElementId()
        if not tid or tid.IntegerValue <= 0: return ""
        tid_i = tid.IntegerValue
        if tid_i in _duct_type_desc_cache: return _duct_type_desc_cache[tid_i]
        t = doc.GetElement(tid)
        val = ""
        if t:
            pp = t.LookupParameter("MAN_TypeDescription_IT")
            if pp: val = (pp.AsString() or "") or (pp.AsValueString() or "")
        _duct_type_desc_cache[tid_i] = val
        return val
    except:
        return ""

def _feet_to_mm_duct(val_ft):
    try:
        if _HAS_UTID:
            return float(UnitUtils.ConvertFromInternalUnits(float(val_ft), UnitTypeId.Millimeters))
        else:
            return float(UnitUtils.ConvertFromInternalUnits(float(val_ft), DisplayUnitType.DUT_MILLIMETERS))
    except:
        try: return float(val_ft) * 304.8
        except: return 0.0

def _size_mm_key_duct(elem):
    """Restituisce (size_key_mm, size_val_mm) come float in mm.
       Se circolare usa Diameter; altrimenti max(Width, Height)."""
    d = None; w = None; h = None
    try:
        p = elem.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
        if p: d = p.AsDouble()
    except: pass
    if d and d > 0:
        mm = _feet_to_mm_duct(d)
        return (round(mm, KEY_MM_PREC_DUCT), mm)

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
    mm = _feet_to_mm_duct(mx)
    return (round(mm, KEY_MM_PREC_DUCT), mm)

# -------- Excel helpers --------
def get_sheet_or_create_duct(workbook, name):
    try:
        return workbook.Worksheets.Item[name]
    except:
        sh = workbook.Worksheets.Add()
        sh.Name = name
        return sh

def ensure_headers_duct(sheet):
    last_col = sheet.Cells(HEADER_ROW_DUCT, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
    if last_col < 1: last_col = 1
    headers = {}
    for c in range(1, last_col + 1):
        v = sheet.Cells(HEADER_ROW_DUCT, c).Value2
        try:
            nm = v.strip() if isinstance(v, String) else None
        except:
            nm = None
        if nm: headers[nm] = c
    next_col = last_col + 1
    for h in OUR_HEADERS_DUCT:
        if h not in headers:
            sheet.Cells(HEADER_ROW_DUCT, next_col).Value2 = h
            headers[h] = next_col
            next_col += 1
    return headers

def _read_column_block_duct(sheet, col, r0, r1):
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
            out.append(_u_duct(val).strip())
        return out
    if isinstance(data, (tuple, list)):
        for row in data:
            try: val = row[0]
            except: val = row
            out.append(_u_duct(val).strip())
        return out
    out.append(_u_duct(data).strip())
    return out

def detect_data_region_duct(sheet, headers):
    tn_col = headers["Type Name"]
    sz_col = headers["Width/Height - Diameter"]
    r0 = MIN_START_DATA_ROW_DUCT
    step = 2000
    curr_start = r0
    last_data_row = r0 - 1
    empty_run = 0
    max_rows = sheet.Rows.Count
    while curr_start <= max_rows:
        r1_try = min(max_rows, curr_start + step - 1)
        t_vals = _read_column_block_duct(sheet, tn_col, curr_start, r1_try)
        s_vals = _read_column_block_duct(sheet, sz_col, curr_start, r1_try)
        block_len = max(len(t_vals), len(s_vals))
        for i in range(block_len):
            t = t_vals[i] if i < len(t_vals) else ""
            s = s_vals[i] if i < len(s_vals) else ""
            if (t or s):
                last_data_row = curr_start + i
                empty_run = 0
            else:
                empty_run += 1
                if empty_run >= EMPTY_RUN_STOP_DUCT:
                    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)
        curr_start = r1_try + 1
    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)

def build_existing_index_bulk_duct(sheet, headers):
    (r0, r1) = detect_data_region_duct(sheet, headers)
    if r1 < r0:
        return {}, (r0, r1)
    tn_col = headers["Type Name"]; sz_col = headers["Width/Height - Diameter"]
    col_tn = _read_column_block_duct(sheet, tn_col, r0, r1)
    col_sz = _read_column_block_duct(sheet, sz_col, r0, r1)
    index = {}
    n = max(len(col_tn), len(col_sz))
    for i in range(n):
        t = _norm_text_duct(col_tn[i] if i < len(col_tn) else "")
        s = col_sz[i] if i < len(col_sz) else ""
        try:
            s_val = float(s) if s != "" else 0.0
        except:
            try: s_val = float(_u_duct(s).replace(",", "."))
            except: s_val = 0.0
        s_key = round(s_val, KEY_MM_PREC_DUCT)
        if t or (s_key != 0.0):
            index[(t, s_key)] = r0 + i
    return index, (r0, r1)

def first_empty_row_after_region_duct(region_tuple):
    r0, r1 = region_tuple
    if r1 < r0: return MIN_START_DATA_ROW_DUCT
    return r1 + 1

def _chunk_consecutive_rows_duct(sorted_rows):
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

def delete_rows_batched_duct(sheet, rows_to_delete):
    if not rows_to_delete: return 0
    rows_sorted = sorted(rows_to_delete, reverse=True)
    runs = _chunk_consecutive_rows_duct(rows_sorted)
    count = 0
    for (r_start, r_end) in runs:
        rng = sheet.Range[sheet.Rows[r_start], sheet.Rows[r_end]]
        rng.EntireRow.Delete()
        count += (r_start - r_end + 1)
    return count

def write_updates_batched_duct(sheet, headers, updates):
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
        "Width/Height - Diameter": headers["Width/Height - Diameter"],
    }
    keys = ["Category","Type Name","MAN_TypeDescription_IT","Width/Height - Diameter"]
    for (r0, r1, vals_list) in runs:
        n = (r1 - r0 + 1)
        for j, key in enumerate(keys):
            col = col_map[key]
            rng = sheet.Range[sheet.Cells(r0, col), sheet.Cells(r1, col)]
            data = Array.CreateInstance(Object, n, 1)
            for i in range(n):
                v = vals_list[i][j]
                if key == "Width/Height - Diameter":
                    try: v = float(v)
                    except:
                        try: v = float(_u_duct(v).replace(",", "."))  # fallback
                        except: v = 0.0
                data[i,0] = v
            rng.Value2 = data
    return len(updates)

def write_appends_batched_duct(sheet, start_row, headers, rows_data):
    if not rows_data: return 0
    if start_row < MIN_START_DATA_ROW_DUCT:
        start_row = MIN_START_DATA_ROW_DUCT
    cols = [headers[h] for h in OUR_HEADERS_DUCT]
    n_rows = len(rows_data); n_cols = len(cols)
    data = Array.CreateInstance(Object, n_rows, n_cols)
    for i in range(n_rows):
        r = rows_data[i]
        for j in range(n_cols):
            v = r[j]
            if OUR_HEADERS_DUCT[j] == "Width/Height - Diameter":
                try: v = float(v)
                except:
                    try: v = float(_u_duct(v).replace(",", "."))  # fallback
                    except: v = 0.0
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
                if OUR_HEADERS_DUCT[idx] == "Width/Height - Diameter":
                    try: v = float(v)
                    except:
                        try: v = float(_u_duct(v).replace(",", "."))  # fallback
                        except: v = 0.0
                col_data[r, 0] = v
            col_rng.Value2 = col_data
    return n_rows

def sort_data_region_duct(sheet, headers):
    r0 = MIN_START_DATA_ROW_DUCT
    last_col = sheet.Cells(HEADER_ROW_DUCT, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

    # Trova l'ultima riga piena considerando le colonne chiave
    tn_col = headers["Type Name"]
    sz_col = headers["Width/Height - Diameter"]

    last_used_rows = []
    for col in (tn_col, sz_col):
        try:
            r = sheet.Cells(sheet.Rows.Count, col).End(Excel.XlDirection.xlUp).Row
            last_used_rows.append(r)
        except:
            pass

    r1 = max([r for r in last_used_rows if r >= r0] or [r0 - 1])
    if r1 < r0:
        return  # niente da ordinare

    data_rng = sheet.Range(sheet.Cells(r0, 1), sheet.Cells(r1, last_col))

    sort = sheet.Sort
    sort.SortFields.Clear()

    # Ordine: Type Name -> Width/Height - Diameter
    sort.SortFields.Add(
        Key=sheet.Range(sheet.Cells(r0, tn_col), sheet.Cells(r1, tn_col)),
        SortOn=Excel.XlSortOn.xlSortOnValues,
        Order=Excel.XlSortOrder.xlAscending,
        DataOption=Excel.XlSortDataOption.xlSortNormal
    )
    sort.SortFields.Add(
        Key=sheet.Range(sheet.Cells(r0, sz_col), sheet.Cells(r1, sz_col)),
        SortOn=Excel.XlSortOn.xlSortOnValues,
        Order=Excel.XlSortOrder.xlAscending,
        DataOption=Excel.XlSortDataOption.xlSortNormal
    )

    sort.SetRange(data_rng)
    sort.Header = Excel.XlYesNoGuess.xlNo
    sort.MatchCase = False
    sort.Orientation = Excel.XlSortOrientation.xlSortColumns
    sort.Apply()

def run_ducts_into_workbook(workbook):
    # Raccoglie i Duct (rigidi) e crea coppie (Type Name, MaxDim_mm)
    try:
        from Autodesk.Revit.DB.Mechanical import Duct
    except:
        # se l'import fallisce, niente da fare
        return

    ducts = FilteredElementCollector(doc).OfClass(Duct).WhereElementIsNotElementType().ToElements()

    groups = {}
    for d in ducts:
        tname = _type_name_duct(d) or ""
        if not tname: continue
        size_key, size_val = _size_mm_key_duct(d)
        if size_key <= 0:  # ignora elementi senza dimensioni utili
            continue

        if tname not in groups:
            groups[tname] = {}
        if size_key not in groups[tname]:
            groups[tname][size_key] = {
                "category": _category_name_duct(d) or "Ducts",
                "type_name": tname,
                "type_desc": _man_type_description_it_duct(d) or "",
                "size_mm": size_key  # già arrotondato
            }
        else:
            g = groups[tname][size_key]
            if not g["type_desc"]:
                g["type_desc"] = _man_type_description_it_duct(d) or g["type_desc"]

    # ordina per Type Name, poi per Size
    def sort_key_type(t): return t or ""
    def sort_key_size(s): return float(s)

    ordered = []
    current_keys = set()
    for t in sorted(groups.keys(), key=sort_key_type):
        inner = groups[t]
        for s in sorted(inner.keys(), key=sort_key_size):
            g = inner[s]
            ordered.append([g["category"], g["type_name"], g["type_desc"], g["size_mm"]])
            current_keys.add((_norm_text_duct(g["type_name"]), float(g["size_mm"])))

    # scrittura Excel
    sheet = None
    try:
        sheet = get_sheet_or_create_duct(workbook, SHEET_NAME_DUCT)
        headers = ensure_headers_duct(sheet)
        existing, region = build_existing_index_bulk_duct(sheet, headers)

        updates = []; appends = []
        for cat, tname, tdesc, size in ordered:
            tkey = _norm_text_duct(tname)
            skey = round(float(size), KEY_MM_PREC_DUCT)
            key = (tkey, skey)
            if key in existing:
                updates.append((existing[key], [cat, tname, tdesc, skey]))
            else:
                appends.append([cat, tname, tdesc, skey])

        rows_to_delete = []; removed_keys = []
        for key, row in existing.items():
            if key not in current_keys:
                rows_to_delete.append(row); removed_keys.append(key)

        if updates: write_updates_batched_duct(sheet, headers, updates)
        removed_count = delete_rows_batched_duct(sheet, rows_to_delete) if rows_to_delete else 0
        added_count = 0
        if appends:
            start_row = first_empty_row_after_region_duct(region)
            added_count = write_appends_batched_duct(sheet, start_row, headers, appends)

        sort_data_region_duct(sheet, headers)

        print("[DUCTS] Aggiunte:", added_count)
        if appends:
            prev = [(r[1], r[3]) for r in appends[:20]]
            print("[DUCTS] Aggiunte (prime 20): {}".format(prev))
        print("[DUCTS] Eliminate:", removed_count)
        if removed_keys:
            print("[DUCTS] Eliminate (prime 20): {}".format(removed_keys[:20]))
    finally:
        try:
            if sheet: Marshal.ReleaseComObject(sheet)
        except: pass

# ============================================================
# ===== BLOCCO 6 — DUCT INSULATIONS (solo Duct come host) ======
# ============================================================
DIN_SHEET_NAME = "Isolamento canali"
DIN_HEADER_ROW = 3
DIN_MIN_START_DATA_ROW = 5
DIN_HEADERS = ["Category", "Type Name", "MAN_TypeDescription_IT", "Insulation Thickness"]
DIN_EMPTY_RUN_STOP = 20
FEET_TO_MM_DIN = 304.8

def _norm_text_din(s):
    if s is None: return ""
    try:
        u = unicode(s) if not isinstance(s, unicode) else s
    except:
        u = str(s)
    return u.strip()

def _fmt_mm_din(x):
    s = ("%.3f" % float(x)).rstrip("0").rstrip(".")
    return s if s else "0"

def _to_number_or_text_for_thickness_din(s):
    if s is None: return ""
    try:
        ss = unicode(s) if not isinstance(s, unicode) else s
    except:
        ss = str(s)
    ss = ss.replace(",", ".").strip()
    try:
        return float(ss)
    except:
        return ss

def category_name_ui_en_din(elem):
    try:
        if elem.Category and elem.Category.Name:
            return elem.Category.Name
    except: pass
    return ""

def type_name_from_instance_din(elem):
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

_type_desc_cache_din = {}
def man_type_description_it_from_type_din(elem):
    try:
        tid_param = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
        if not tid_param: return ""
        tid = tid_param.AsElementId()
        if not tid or tid.IntegerValue <= 0: return ""
        tid_int = tid.IntegerValue
        if tid_int in _type_desc_cache_din: return _type_desc_cache_din[tid_int]
        t = doc.GetElement(tid)
        val = ""
        if t:
            p = t.LookupParameter("MAN_TypeDescription_IT")
            if p: val = (p.AsString() or "") or (p.AsValueString() or "")
        _type_desc_cache_din[tid_int] = val
        return val
    except:
        return ""

def as_string_or_valuestring_din(p):
    if not p: return ""
    try:
        s = p.AsString()
        if s: return s
    except: pass
    try:
        vs = p.AsValueString()
        if vs: return vs
    except: pass
    return ""

def is_duct_hosted_din(ins_elem):
    try:
        hid = ins_elem.HostElementId
        if hid and hid.IntegerValue > 0:
            host = doc.GetElement(hid)
            return isinstance(host, Duct)
    except: pass
    return False

def thickness_only_din(elem):
    """Ritorna (thick_key, thick_disp) in mm come stringa normalizzata."""
    thick_key = ""; thick_disp = ""
    pth = None
    # prendiamo il parametro spessore con fallback
    try:
        pth = elem.get_Parameter(getattr(BuiltInParameter, "RBS_INSULATION_THICKNESS_FOR_DUCT"))
    except: pass
    if not pth:
        try:
            pth = elem.get_Parameter(getattr(BuiltInParameter, "RBS_INSULATION_THICKNESS"))
        except: pass
    if not pth:
        try:
            pth = elem.get_Parameter(getattr(BuiltInParameter, "RBS_INSULATION_THICKNESS_FOR_PIPE"))
        except: pass

    if pth:
        # prova come double (internal feet)
        try:
            d_ft = pth.AsDouble()
            if d_ft is not None:
                d_mm = d_ft * FEET_TO_MM_DIN
                thick_disp = _fmt_mm_din(d_mm)
                thick_key  = thick_disp
                return thick_key, thick_disp
        except:
            pass
        # fallback a ValueString
        disp = as_string_or_valuestring_din(pth)
        if disp:
            # estrai numero base
            s = disp.replace(",", ".")
            m = re.search(r'(\d+(?:\.\d+)?)', s)
            if m:
                num = m.group(1)
                if "." in num: num = num.rstrip("0").rstrip(".")
                thick_disp = num
                thick_key  = num
                return thick_key, thick_disp
    return "", ""

def get_sheet_or_create_din(workbook, name):
    try:
        return workbook.Worksheets.Item[name]
    except:
        sh = workbook.Worksheets.Add()
        sh.Name = name
        return sh

def ensure_headers_din(sheet):
    last_col = sheet.Cells(DIN_HEADER_ROW, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
    if last_col < 1: last_col = 1
    headers = {}
    for c in range(1, last_col + 1):
        v = sheet.Cells(DIN_HEADER_ROW, c).Value2
        if isinstance(v, String):
            nm = v.strip()
            if nm: headers[nm] = c
    next_col = last_col + 1
    for h in DIN_HEADERS:
        if h not in headers:
            sheet.Cells(DIN_HEADER_ROW, next_col).Value2 = h
            headers[h] = next_col
            next_col += 1
    return headers

def _read_column_block_din(sheet, col, r0, r1):
    if r1 < r0: return []
    rng = sheet.Range[sheet.Cells(r0, col), sheet.Cells(r1, col)]
    data = rng.Value2
    out = []
    if isinstance(data, System.Array) and data.Rank == 2:
        n0 = data.GetLength(0)
        for i in range(n0):
            try:
                val = data.GetValue(i, 0)
            except:
                try:
                    val = data.GetValue(i+1, 1)
                except:
                    val = None
            out.append(_norm_text_din(val))
        return out
    if isinstance(data, (tuple, list)):
        for row in data:
            try: val = row[0]
            except: val = row
            out.append(_norm_text_din(val))
        return out
    out.append(_norm_text_din(data))
    return out

def detect_data_region_din(sheet, headers):
    tn_col = headers["Type Name"]
    th_col = headers["Insulation Thickness"]
    r0 = DIN_MIN_START_DATA_ROW
    step = 2000
    curr_start = r0
    last_data_row = r0 - 1
    empty_run = 0
    max_rows = sheet.Rows.Count
    while curr_start <= max_rows:
        r1_try = min(max_rows, curr_start + step - 1)
        t_vals  = _read_column_block_din(sheet, tn_col, curr_start, r1_try)
        th_vals = _read_column_block_din(sheet, th_col, curr_start, r1_try)
        block_len = max(len(t_vals), len(th_vals))
        for i in range(block_len):
            t  = t_vals[i]  if i < len(t_vals)  else ""
            th = th_vals[i] if i < len(th_vals) else ""
            if (t or th):
                last_data_row = curr_start + i
                empty_run = 0
            else:
                empty_run += 1
                if empty_run >= DIN_EMPTY_RUN_STOP:
                    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)
        curr_start = r1_try + 1
    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)

def build_existing_index_bulk_din(sheet, headers):
    (r0, r1) = detect_data_region_din(sheet, headers)
    if r1 < r0:
        return {}, (r0, r1)
    tn_col = headers["Type Name"]; th_col = headers["Insulation Thickness"]
    col_tn = _read_column_block_din(sheet, tn_col, r0, r1)
    col_th = _read_column_block_din(sheet, th_col, r0, r1)
    index = {}
    n = max(len(col_tn), len(col_th))
    for i in range(n):
        t  = col_tn[i] if i < len(col_tn) else ""
        th = col_th[i] if i < len(col_th) else ""
        # normalizza chiave spessore come stringa numerica
        try:
            th_key = ("%.6f" % float(str(th).replace(",", "."))).rstrip("0").rstrip(".")
        except:
            th_key = _norm_text_din(th)
        if t or th_key:
            index[(t, th_key)] = r0 + i
    return index, (r0, r1)

def first_empty_row_after_region_din(region_tuple):
    r0, r1 = region_tuple
    if r1 < r0: return DIN_MIN_START_DATA_ROW
    return r1 + 1

def _chunk_consecutive_rows_din(sorted_rows):
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

def delete_rows_batched_din(sheet, rows_to_delete):
    if not rows_to_delete: return 0
    rows_sorted = sorted(rows_to_delete, reverse=True)
    runs = _chunk_consecutive_rows_din(rows_sorted)
    count = 0
    for (r_start, r_end) in runs:
        rng = sheet.Range[sheet.Rows[r_start], sheet.Rows[r_end]]
        rng.EntireRow.Delete()
        count += (r_start - r_end + 1)
    return count

def write_updates_batched_din(sheet, headers, updates):
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
        "Insulation Thickness": headers["Insulation Thickness"],
    }
    keys = ["Category","Type Name","MAN_TypeDescription_IT","Insulation Thickness"]
    for (r0, r1, vals_list) in runs:
        n = (r1 - r0 + 1)
        for j, key in enumerate(keys):
            col = col_map[key]
            rng = sheet.Range[sheet.Cells(r0, col), sheet.Cells(r1, col)]
            data = Array.CreateInstance(Object, n, 1)
            for i in range(n):
                val = vals_list[i][j]
                if key == "Insulation Thickness":
                    val = _to_number_or_text_for_thickness_din(val)
                data[i,0] = val
            rng.Value2 = data
    return len(updates)

def write_appends_din(sheet, start_row, headers, rows_data):
    if not rows_data: return 0
    cols = [headers[h] for h in DIN_HEADERS]
    n_rows = len(rows_data); n_cols = len(cols)
    data = Array.CreateInstance(Object, n_rows, n_cols)
    for i in range(n_rows):
        r = rows_data[i]
        for j in range(n_cols):
            val = r[j]; head = DIN_HEADERS[j]
            if head == "Insulation Thickness":
                val = _to_number_or_text_for_thickness_din(val)
            data[i, j] = val
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
                val = rows_data[r][idx]; head = DIN_HEADERS[idx]
                if head == "Insulation Thickness":
                    val = _to_number_or_text_for_thickness_din(val)
                col_data[r, 0] = val
            col_rng.Value2 = col_data
    return n_rows

def sort_data_region_din(sheet, headers):
    r0 = DIN_MIN_START_DATA_ROW
    first_col = 1
    last_col = sheet.Cells(DIN_HEADER_ROW, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

    tn_col = headers["Type Name"]
    th_col = headers["Insulation Thickness"]

    last_row = 0
    for col in (tn_col, th_col):
        lr = sheet.Cells(sheet.Rows.Count, col).End(Excel.XlDirection.xlUp).Row
        if lr > last_row: last_row = lr
    if last_row < r0: return

    data_rng = sheet.Range(sheet.Cells(r0, first_col), sheet.Cells(last_row, last_col))

    sort = sheet.Sort
    sort.SortFields.Clear()
    sort.SortFields.Add(Key=sheet.Range(sheet.Cells(r0, tn_col), sheet.Cells(last_row, tn_col)),
                        SortOn=Excel.XlSortOn.xlSortOnValues, Order=Excel.XlSortOrder.xlAscending,
                        DataOption=Excel.XlSortDataOption.xlSortNormal)
    sort.SortFields.Add(Key=sheet.Range(sheet.Cells(r0, th_col), sheet.Cells(last_row, th_col)),
                        SortOn=Excel.XlSortOn.xlSortOnValues, Order=Excel.XlSortOrder.xlAscending,
                        DataOption=Excel.XlSortDataOption.xlSortNormal)
    sort.SetRange(data_rng)
    sort.Header = Excel.XlYesNoGuess.xlNo
    sort.MatchCase = False
    sort.Orientation = Excel.XlSortOrientation.xlSortColumns
    sort.Apply()

def run_duct_ins_into_workbook(workbook):
    insulations = FilteredElementCollector(doc).OfClass(DuctInsulation).WhereElementIsNotElementType().ToElements()
    groups = {}
    for ins in insulations:
        if not is_duct_hosted_din(ins): continue
        tname = type_name_from_instance_din(ins) or ""
        th_key, th_disp = thickness_only_din(ins)
        if not th_key: continue
        inner = groups.get(tname)
        if inner is None:
            inner = {}; groups[tname] = inner
        if th_key not in inner:
            inner[th_key] = {
                "category":  category_name_ui_en_din(ins) or "Duct Insulations",
                "type_name": tname,
                "type_desc": man_type_description_it_from_type_din(ins) or "",
                "thick_disp": th_disp or "",
            }
        else:
            g = inner[th_key]
            if not g["type_desc"]:
                g["type_desc"] = man_type_description_it_from_type_din(ins) or g["type_desc"]

    rows_tmp = []
    for tname, inner in groups.items():
        for th_key, g in inner.items():
            try: th_num = float(th_key) if th_key not in ("", None) else 0.0
            except: th_num = 0.0
            rows_tmp.append((_norm_text_din(g["type_name"]), th_num,
                             [g["category"], g["type_name"], g["type_desc"], g["thick_disp"]]))
    rows_tmp.sort(key=lambda x: (x[0], x[1]))

    ordered = []
    current_keys = set()
    for _, _, rowvals in rows_tmp:
        ordered.append(rowvals)
        current_keys.add((_norm_text_din(rowvals[1]),  # Type Name
                          _norm_text_din(rowvals[3]))) # Thickness

    sheet = None
    try:
        sheet = get_sheet_or_create_din(workbook, DIN_SHEET_NAME)
        headers = ensure_headers_din(sheet)
        existing, region = build_existing_index_bulk_din(sheet, headers)

        updates = []; appends = []
        for cat, tname, tdesc, thick in ordered:
            tkey  = _norm_text_din(tname)
            try:
                thkey = ("%.6f" % float(str(thick).replace(",", "."))).rstrip("0").rstrip(".")
            except:
                thkey = _norm_text_din(thick)
            key = (tkey, thkey)
            if key in existing:
                updates.append((existing[key], [cat, tname, tdesc, thick]))
            else:
                appends.append([cat, tname, tdesc, thick])

        rows_to_delete = []; removed_keys = []
        for key, row in existing.items():
            if key not in current_keys:
                rows_to_delete.append(row); removed_keys.append(key)

        if updates: write_updates_batched_din(sheet, headers, updates)
        removed_count = delete_rows_batched_din(sheet, rows_to_delete) if rows_to_delete else 0
        added_count = 0
        if appends:
            start_row = first_empty_row_after_region_din(region)
            added_count = write_appends_din(sheet, start_row, headers, appends)

        sort_data_region_din(sheet, headers)

        print("[DUCT INS] Aggiunte:", added_count)
        if appends: print("[DUCT INS] Aggiunte ({}): {}".format(min(20, len(appends)), [(r[1], r[3]) for r in appends[:20]]))
        print("[DUCT INS] Eliminate:", removed_count)
        if removed_keys: print("[DUCT INS] Eliminate ({}): {}".format(min(20, len(removed_keys)), removed_keys[:20]))
    finally:
        try:
            if sheet: Marshal.ReleaseComObject(sheet)
        except: pass

# ============================================================
# ===== BLOCCO 7 — DUCT FITTINGS (Family/Type → MaxSize) =====
# ============================================================
DFT_SHEET_NAME = "Fitting canali"
DFT_HEADER_ROW = 3
DFT_MIN_START_DATA_ROW = 5
DFT_HEADERS = ["Category", "Family Name", "Type Name", "MAN_TypeDescription_IT", "MAN_Fittings_MaxSize"]
DFT_EMPTY_RUN_STOP = 20
DFT_KEY_MM_PREC = 6

def DFT_u(s):
    if s is None: return u""
    try: return s if isinstance(s, unicode) else unicode(s)
    except: return unicode(str(s))

def DFT_norm_text(s):
    return DFT_u(s).strip()

def DFT_norm_text_strong(s):
    return u" ".join(DFT_norm_text(s).split())

def DFT_to_float(x):
    if x is None or x == "": return 0.0
    try: return float(x)
    except:
        try: return float(DFT_u(x).replace(",", ".").strip())
        except: return 0.0

def DFT_norm_mm_key(v):
    try: return round(float(v), DFT_KEY_MM_PREC)
    except: return 0.0

def DFT_category_name(elem):
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

def DFT_type_name(elem):
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

def DFT_family_name(elem):
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

_DFT_type_desc_cache = {}
def DFT_type_desc(elem):
    try:
        p = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
        if not p: return ""
        tid = p.AsElementId()
        if not tid or tid.IntegerValue <= 0: return ""
        tid_i = tid.IntegerValue
        if tid_i in _DFT_type_desc_cache: return _DFT_type_desc_cache[tid_i]
        t = doc.GetElement(tid)
        val = ""
        if t:
            pp = t.LookupParameter("MAN_TypeDescription_IT")
            if pp: val = (pp.AsString() or "") or (pp.AsValueString() or "")
        _DFT_type_desc_cache[tid_i] = val
        return val
    except:
        return ""

def DFT_feet_to_mm(val_ft):
    try:
        if _HAS_UTID:
            return float(UnitUtils.ConvertFromInternalUnits(float(val_ft), UnitTypeId.Millimeters))
        else:
            return float(UnitUtils.ConvertFromInternalUnits(float(val_ft), DisplayUnitType.DUT_MILLIMETERS))
    except:
        try: return float(val_ft) * 304.8
        except: return 0.0

def DFT_maxsize_mm(elem):
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
        return DFT_feet_to_mm(d_ft)
    except:
        return 0.0

def DFT_get_sheet_or_create(workbook, name):
    try:
        return workbook.Worksheets.Item[name]
    except:
        sh = workbook.Worksheets.Add()
        sh.Name = name
        return sh

def DFT_ensure_headers(sheet):
    last_col = sheet.Cells(DFT_HEADER_ROW, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
    if last_col < 1: last_col = 1
    headers = {}
    for c in range(1, last_col + 1):
        v = sheet.Cells(DFT_HEADER_ROW, c).Value2
        try:
            nm = v.strip() if isinstance(v, String) else None
        except:
            nm = None
        if nm: headers[nm] = c
    next_col = last_col + 1
    for h in DFT_HEADERS:
        if h not in headers:
            sheet.Cells(DFT_HEADER_ROW, next_col).Value2 = h
            headers[h] = next_col
            next_col += 1
    return headers

def DFT_read_column_block(sheet, col, r0, r1):
    if r1 < r0: return []
    rng = sheet.Range[sheet.Cells(r0, col), sheet.Cells(r1, col)]
    data = rng.Value2
    out = []
    if isinstance(data, System.Array) and data.Rank == 2:
        n0 = data.GetLength(0)
        for i in range(n0):
            try:
                val = data.GetValue(i, 0)
            except:
                try:
                    val = data.GetValue(i+1, 1)
                except:
                    val = None
            out.append(DFT_u(val).strip())
        return out
    if isinstance(data, (tuple, list)):
        for row in data:
            try: val = row[0]
            except: val = row
            out.append(DFT_u(val).strip())
        return out
    out.append(DFT_u(data).strip())
    return out

def DFT_detect_data_region(sheet, headers):
    fam_col = headers["Family Name"]
    typ_col = headers["Type Name"]
    max_col = headers["MAN_Fittings_MaxSize"]
    r0 = DFT_MIN_START_DATA_ROW
    step = 2000
    curr_start = r0
    last_data_row = r0 - 1
    empty_run = 0
    max_rows = sheet.Rows.Count
    while curr_start <= max_rows:
        r1_try = min(max_rows, curr_start + step - 1)
        f_vals = DFT_read_column_block(sheet, fam_col, curr_start, r1_try)
        t_vals = DFT_read_column_block(sheet, typ_col, curr_start, r1_try)
        m_vals = DFT_read_column_block(sheet, max_col, curr_start, r1_try)
        block_len = max(len(f_vals), len(t_vals), len(m_vals))
        for i in range(block_len):
            f = f_vals[i] if i < len(f_vals) else ""
            t = t_vals[i] if i < len(t_vals) else ""
            m = m_vals[i] if i < len(m_vals) else ""
            if (f or t or m):
                last_data_row = curr_start + i
                empty_run = 0
            else:
                empty_run += 1
                if empty_run >= DFT_EMPTY_RUN_STOP:
                    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)
        curr_start = r1_try + 1
    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)

def DFT_build_existing_index_bulk(sheet, headers):
    (r0, r1) = DFT_detect_data_region(sheet, headers)
    if r1 < r0:
        return {}, (r0, r1)
    fam_col = headers["Family Name"]; typ_col = headers["Type Name"]; max_col = headers["MAN_Fittings_MaxSize"]
    col_f = DFT_read_column_block(sheet, fam_col, r0, r1)
    col_t = DFT_read_column_block(sheet, typ_col, r0, r1)
    col_m = DFT_read_column_block(sheet, max_col, r0, r1)
    index = {}
    n = max(len(col_f), len(col_t), len(col_m))
    for i in range(n):
        f = DFT_norm_text_strong(col_f[i] if i < len(col_f) else "")
        t = DFT_norm_text_strong(col_t[i] if i < len(col_t) else "")
        m_raw = col_m[i] if i < len(col_m) else ""
        try:
            m_val = float(m_raw) if m_raw != "" else 0.0
        except:
            m_val = DFT_to_float(m_raw)
        m_key = DFT_norm_mm_key(m_val)
        if f or t or (m_key != 0.0):
            index[(f, t, m_key)] = r0 + i
    return index, (r0, r1)

def DFT_first_empty_row_after_region(region_tuple):
    r0, r1 = region_tuple
    if r1 < r0: return DFT_MIN_START_DATA_ROW
    return r1 + 1

def DFT_chunk_consecutive_rows(sorted_rows):
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

def DFT_delete_rows_batched(sheet, rows_to_delete):
    if not rows_to_delete: return 0
    rows_sorted = sorted(rows_to_delete, reverse=True)
    runs = DFT_chunk_consecutive_rows(rows_sorted)
    count = 0
    for (r_start, r_end) in runs:
        rng = sheet.Range[sheet.Rows[r_start], sheet.Rows[r_end]]
        rng.EntireRow.Delete()
        count += (r_start - r_end + 1)
    return count

def DFT_write_updates_batched(sheet, headers, updates):
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
        "MAN_Fittings_MaxSize": headers["MAN_Fittings_MaxSize"],
    }
    keys = ["Category","Family Name","Type Name","MAN_TypeDescription_IT","MAN_Fittings_MaxSize"]
    for (r0, r1, vals_list) in runs:
        n = (r1 - r0 + 1)
        for j, key in enumerate(keys):
            col = col_map[key]
            rng = sheet.Range[sheet.Cells(r0, col), sheet.Cells(r1, col)]
            data = Array.CreateInstance(Object, n, 1)
            for i in range(n):
                val = vals_list[i][j]
                if key == "MAN_Fittings_MaxSize":
                    val = float(DFT_to_float(val))
                data[i,0] = val
            rng.Value2 = data
    return len(updates)

def DFT_write_appends_batched(sheet, start_row, headers, rows_data):
    if not rows_data: return 0
    if start_row < DFT_MIN_START_DATA_ROW:
        start_row = DFT_MIN_START_DATA_ROW
    cols = [headers[h] for h in DFT_HEADERS]
    n_rows = len(rows_data); n_cols = len(cols)
    data = Array.CreateInstance(Object, n_rows, n_cols)
    for i in range(n_rows):
        r = rows_data[i]
        for j in range(n_cols):
            val = r[j]
            if DFT_HEADERS[j] == "MAN_Fittings_MaxSize":
                val = float(DFT_to_float(val))  # numerico in mm
            data[i, j] = val
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
                val = rows_data[r][idx]
                if DFT_HEADERS[idx] == "MAN_Fittings_MaxSize":
                    val = float(DFT_to_float(val))
                col_data[r, 0] = val
            col_rng.Value2 = col_data
    return n_rows

def DFT_sort_data_region(sheet, headers):
    r0 = DFT_MIN_START_DATA_ROW
    first_col = 1
    last_col = sheet.Cells(DFT_HEADER_ROW, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

    fam_col = headers["Family Name"]
    typ_col = headers["Type Name"]
    max_col = headers["MAN_Fittings_MaxSize"]

    last_row = 0
    for col in (fam_col, typ_col, max_col):
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
    sort.SortFields.Add(Key=sheet.Range(sheet.Cells(r0, max_col), sheet.Cells(last_row, max_col)),
                        SortOn=Excel.XlSortOn.xlSortOnValues, Order=Excel.XlSortOrder.xlAscending,
                        DataOption=Excel.XlSortDataOption.xlSortNormal)
    sort.SetRange(data_rng)
    sort.Header = Excel.XlYesNoGuess.xlNo
    sort.MatchCase = False
    sort.Orientation = Excel.XlSortOrientation.xlSortColumns
    sort.Apply()

def run_duct_fittings_into_workbook(workbook):
    elems = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_DuctFitting) \
        .WhereElementIsNotElementType() \
        .ToElements()

    groups = {}
    for e in elems:
        if not isinstance(e, FamilyInstance):
            continue
        cat  = DFT_category_name(e) or "Duct Fittings"
        fam  = DFT_family_name(e) or ""
        typ  = DFT_type_name(e) or ""
        desc = DFT_type_desc(e) or ""
        msz_mm = DFT_maxsize_mm(e)

        fam_k = DFT_norm_text_strong(fam)
        typ_k = DFT_norm_text_strong(typ)
        msz_k = DFT_norm_mm_key(msz_mm)

        key = (fam_k, typ_k, msz_k)
        if key not in groups:
            groups[key] = [cat, fam_k, typ_k, desc, msz_k]
        else:
            if not groups[key][3]:
                groups[key][3] = desc

    rows_tmp = []
    for (fam_k, typ_k, msz_k), vals in groups.items():
        rows_tmp.append((fam_k, typ_k, msz_k, vals))
    rows_tmp.sort(key=lambda x: (x[0] or "", x[1] or "", x[2]))

    ordered = []
    current_keys = set()
    for fam_k, typ_k, msz_k, vals in rows_tmp:
        ordered.append(vals)
        current_keys.add((fam_k, typ_k, msz_k))

    sheet = None
    try:
        sheet = DFT_get_sheet_or_create(workbook, DFT_SHEET_NAME)
        headers = DFT_ensure_headers(sheet)
        existing, region = DFT_build_existing_index_bulk(sheet, headers)

        updates = []; appends = []
        for cat, fam, typ, desc, msz_mm in ordered:
            fam_k = DFT_norm_text_strong(fam)
            typ_k = DFT_norm_text_strong(typ)
            msz_k = DFT_norm_mm_key(msz_mm)
            key   = (fam_k, typ_k, msz_k)
            if key in existing:
                updates.append((existing[key], [cat, fam_k, typ_k, desc, msz_k]))
            else:
                appends.append([cat, fam_k, typ_k, desc, msz_k])

        rows_to_delete = []; removed_keys = []
        for key, row in existing.items():
            if key not in current_keys:
                rows_to_delete.append(row); removed_keys.append(key)

        if updates: DFT_write_updates_batched(sheet, headers, updates)
        removed_count = DFT_delete_rows_batched(sheet, rows_to_delete) if rows_to_delete else 0
        added_count = 0
        if appends:
            start_row = DFT_first_empty_row_after_region(region)
            added_count = DFT_write_appends_batched(sheet, start_row, headers, appends)

        DFT_sort_data_region(sheet, headers)

        print("[DUCT FIT] Aggiunte:", added_count)
        if appends:
            preview = [(r[1], r[2], round(float(r[4]), 3)) for r in appends[:20]]
            print("[DUCT FIT] Aggiunte:", preview)
        print("[DUCT FIT] Eliminate:", removed_count)
        if removed_keys:
            preview_del = [(k[0], k[1], round(float(k[2]), 3)) for k in removed_keys[:20]]
            print("[DUCT FIT] Eliminate:", preview_del)
    finally:
        try:
            if sheet: Marshal.ReleaseComObject(sheet)
        except: pass

# ============================================================
# ========== BLOCCO 8 — FLEX DUCT (Type → Diameter) =============
# ============================================================
FXD_SHEET_NAME = "Canali Flessibili"
FXD_HEADER_ROW = 3
FXD_MIN_START_DATA_ROW = 5
FXD_HEADERS = ["Category", "Type Name", "MAN_TypeDescription_IT", "Diameter"]
FXD_EMPTY_RUN_STOP = 20

def _fxd_u(s):
    if s is None: return u""
    try: return s if isinstance(s, unicode) else unicode(s)
    except: return unicode(str(s))

def _fxd_norm_text(s):
    return _fxd_u(s).strip()

def _fxd_norm_diam_key(val):
    if val is None: return ""
    try:
        s = _fxd_u(val)
    except:
        s = str(val)
    s = s.strip().replace(",", ".")
    m = re.search(r'(\d+(?:\.\d+)?)', s)
    if not m: return ""
    num = m.group(1)
    if "." in num:
        num = num.rstrip("0").rstrip(".")
    return num

def _fxd_to_number_or_text(s):
    if s is None: return ""
    ss = _fxd_u(s).replace(",", ".").strip()
    try:
        return float(ss)
    except:
        return ss

def fxd_category(elem):
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

def fxd_type_name(elem):
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

_fxd_type_desc_cache = {}
def fxd_type_desc(elem):
    try:
        p = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
        if not p: return ""
        tid = p.AsElementId()
        if not tid or tid.IntegerValue <= 0: return ""
        tid_i = tid.IntegerValue
        if tid_i in _fxd_type_desc_cache: return _fxd_type_desc_cache[tid_i]
        t = doc.GetElement(tid)
        val = ""
        if t:
            pp = t.LookupParameter("MAN_TypeDescription_IT")
            if pp: val = (pp.AsString() or "") or (pp.AsValueString() or "")
        _fxd_type_desc_cache[tid_i] = val
        return val
    except:
        return ""

def _fxd_as_str_or_vs(p):
    if not p: return ""
    try:
        s = p.AsString()
        if s: return s
    except: pass
    try:
        vs = p.AsValueString()
        if vs: return vs
    except: pass
    return ""

def fxd_diameter_key_and_disp(elem):
    raw = None
    try:
        p = elem.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
        if not p: return None, None, raw
        vs = _fxd_as_str_or_vs(p)
        raw = vs
        if vs:
            key = _fxd_norm_diam_key(vs)
            if key:
                return key, key, raw
    except:
        pass
    return None, None, raw

def fxd_get_sheet_or_create(workbook, name):
    try:
        return workbook.Worksheets.Item[name]
    except:
        sh = workbook.Worksheets.Add()
        sh.Name = name
        return sh

def fxd_ensure_headers(sheet):
    last_col = sheet.Cells(FXD_HEADER_ROW, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
    if last_col < 1: last_col = 1
    headers = {}
    for c in range(1, last_col + 1):
        v = sheet.Cells(FXD_HEADER_ROW, c).Value2
        try:
            nm = v.strip() if isinstance(v, String) else None
        except:
            nm = None
        if nm: headers[nm] = c
    next_col = last_col + 1
    for h in FXD_HEADERS:
        if h not in headers:
            sheet.Cells(FXD_HEADER_ROW, next_col).Value2 = h
            headers[h] = next_col
            next_col += 1
    return headers

def _fxd_read_column(sheet, col, r0, r1):
    if r1 < r0: return []
    rng = sheet.Range(sheet.Cells(r0, col), sheet.Cells(r1, col))
    data = rng.Value2
    out = []
    if isinstance(data, System.Array) and data.Rank == 2:
        n0 = data.GetLength(0)
        for i in range(n0):
            try: val = data.GetValue(i, 0)
            except:
                try: val = data.GetValue(i+1, 1)
                except: val = None
            out.append(_fxd_norm_text(val))
        return out
    if isinstance(data, (tuple, list)):
        for row in data:
            try: val = row[0]
            except: val = row
            out.append(_fxd_norm_text(val))
        return out
    out.append(_fxd_norm_text(data))
    return out

def fxd_detect_data_region(sheet, headers):
    t_col = headers["Type Name"]
    d_col = headers["Diameter"]
    r0 = FXD_MIN_START_DATA_ROW
    step = 2000
    curr_start = r0
    last_data_row = r0 - 1
    empty_run = 0
    max_rows = sheet.Rows.Count
    while curr_start <= max_rows:
        r1_try = min(max_rows, curr_start + step - 1)
        t_vals = _fxd_read_column(sheet, t_col, curr_start, r1_try)
        d_vals = _fxd_read_column(sheet, d_col, curr_start, r1_try)
        block_len = max(len(t_vals), len(d_vals))
        for i in range(block_len):
            t = t_vals[i] if i < len(t_vals) else ""
            d = d_vals[i] if i < len(d_vals) else ""
            if (t or d):
                last_data_row = curr_start + i
                empty_run = 0
            else:
                empty_run += 1
                if empty_run >= FXD_EMPTY_RUN_STOP:
                    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)
        curr_start = r1_try + 1
    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)

def fxd_build_existing_index(sheet, headers):
    (r0, r1) = fxd_detect_data_region(sheet, headers)
    if r1 < r0:
        return {}, (r0, r1)
    t_col = headers["Type Name"]; d_col = headers["Diameter"]
    col_t = _fxd_read_column(sheet, t_col, r0, r1)
    col_d = _fxd_read_column(sheet, d_col, r0, r1)
    index = {}
    n = max(len(col_t), len(col_d))
    for i in range(n):
        t = col_t[i] if i < len(col_t) else ""
        d_raw = col_d[i] if i < len(col_d) else ""
        d = _fxd_norm_diam_key(d_raw)
        if t or d:
            index[(t, d)] = r0 + i
    return index, (r0, r1)

def fxd_first_empty_after(region_tuple):
    r0, r1 = region_tuple
    if r1 < r0: return FXD_MIN_START_DATA_ROW
    return r1 + 1

def fxd_delete_rows_batched(sheet, rows_to_delete):
    if not rows_to_delete: return 0
    rows_sorted = sorted(rows_to_delete, reverse=True)
    # chunk
    runs = []
    if rows_sorted:
        s = p = rows_sorted[0]
        for r in rows_sorted[1:]:
            if r == p - 1:
                p = r
            else:
                runs.append((p, s))
                s = p = r
        runs.append((p, s))
    count = 0
    for (r_start, r_end) in runs:
        rng = sheet.Range(sheet.Rows[r_end], sheet.Rows[r_start])
        rng.EntireRow.Delete()
        count += (r_start - r_end + 1)
    return count

def fxd_write_updates(sheet, headers, updates):
    if not updates: return 0
    updates_sorted = sorted(updates, key=lambda x: x[0])
    # chunk
    runs = []
    st = None; buf = []
    for row, vals in updates_sorted:
        if st is None:
            st = pr = row; buf = [vals]
        elif row == pr + 1:
            buf.append(vals); pr = row
        else:
            runs.append((st, pr, buf))
            st = pr = row; buf = [vals]
    if st is not None:
        runs.append((st, pr, buf))

    col_map = {
        "Category": headers["Category"],
        "Type Name": headers["Type Name"],
        "MAN_TypeDescription_IT": headers["MAN_TypeDescription_IT"],
        "Diameter": headers["Diameter"],
    }
    keys = ["Category","Type Name","MAN_TypeDescription_IT","Diameter"]
    for (r0, r1, vals_list) in runs:
        n = (r1 - r0 + 1)
        for j, key in enumerate(keys):
            col = col_map[key]
            rng = sheet.Range(sheet.Cells(r0, col), sheet.Cells(r1, col))
            data = Array.CreateInstance(Object, n, 1)
            for i in range(n):
                val = vals_list[i][j]
                if key == "Diameter":
                    val = _fxd_to_number_or_text(val)
                data[i,0] = val
            rng.Value2 = data
    return len(updates)

def fxd_write_appends(sheet, start_row, headers, rows_data):
    if not rows_data: return 0
    cols = [headers[h] for h in FXD_HEADERS]
    n_rows = len(rows_data); n_cols = len(cols)
    data = Array.CreateInstance(Object, n_rows, n_cols)
    for i in range(n_rows):
        r = rows_data[i]
        for j in range(n_cols):
            val = r[j]
            if FXD_HEADERS[j] == "Diameter":
                val = _fxd_to_number_or_text(val)
            data[i, j] = val
    first_cell = sheet.Cells(start_row, min(cols))
    last_cell  = sheet.Cells(start_row + n_rows - 1, max(cols))
    dest = sheet.Range(first_cell, last_cell)
    expected_span = max(cols) - min(cols) + 1
    if expected_span == n_cols:
        dest.Value2 = data
    else:
        for idx, col in enumerate(cols):
            col_rng = sheet.Range(sheet.Cells(start_row, col), sheet.Cells(start_row + n_rows - 1, col))
            col_data = Array.CreateInstance(Object, n_rows, 1)
            for r in range(n_rows):
                val = rows_data[r][idx]
                if FXD_HEADERS[idx] == "Diameter":
                    val = _fxd_to_number_or_text(val)
                col_data[r, 0] = val
            col_rng.Value2 = col_data
    return n_rows

def fxd_sort_data_region(sheet, headers):
    r0 = FXD_MIN_START_DATA_ROW
    first_col = 1
    last_col = sheet.Cells(FXD_HEADER_ROW, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

    tn_col = headers["Type Name"]
    d_col  = headers["Diameter"]

    last_row = 0
    for col in (tn_col, d_col):
        lr = sheet.Cells(sheet.Rows.Count, col).End(Excel.XlDirection.xlUp).Row
        if lr > last_row: last_row = lr
    if last_row < r0: return

    data_rng = sheet.Range(sheet.Cells(r0, first_col), sheet.Cells(last_row, last_col))

    sort = sheet.Sort
    sort.SortFields.Clear()
    sort.SortFields.Add(Key=sheet.Range(sheet.Cells(r0, tn_col), sheet.Cells(last_row, tn_col)),
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

def run_flexduct_into_workbook(workbook):
    elems = FilteredElementCollector(doc).OfClass(FlexDuct).WhereElementIsNotElementType().ToElements()

    groups = {}
    for e in elems:
        tname = fxd_type_name(e) or ""
        dkey, ddisp, _ = fxd_diameter_key_and_disp(e)
        if not dkey: continue
        inner = groups.get(tname)
        if inner is None:
            inner = {}; groups[tname] = inner
        if dkey not in inner:
            inner[dkey] = {
                "category":  fxd_category(e) or "Flex Ducts",
                "type_name": tname,
                "type_desc": fxd_type_desc(e) or "",
                "diam_disp": ddisp,
            }
        else:
            g = inner[dkey]
            if not g["type_desc"]:
                g["type_desc"] = fxd_type_desc(e) or g["type_desc"]

    # ordina Type, poi Diameter
    def _k_type(t): return t or ""
    def _k_d(d):
        try: return (0, float(d))
        except: return (1, d)

    ordered = []
    current_keys = set()
    for t in sorted(groups.keys(), key=_k_type):
        inner = groups[t]
        for d in sorted(inner.keys(), key=_k_d):
            g = inner[d]
            ordered.append([g["category"], g["type_name"], g["type_desc"], g["diam_disp"]])
            current_keys.add((_fxd_norm_text(g["type_name"]), _fxd_norm_diam_key(g["diam_disp"])))

    sheet = None
    try:
        sheet = fxd_get_sheet_or_create(workbook, FXD_SHEET_NAME)
        headers = fxd_ensure_headers(sheet)
        existing, region = fxd_build_existing_index(sheet, headers)

        updates = []; appends = []
        for cat, tname, tdesc, diam in ordered:
            tkey = _fxd_norm_text(tname)
            dkey = _fxd_norm_diam_key(diam)
            if (tkey, dkey) in existing:
                updates.append((existing[(tkey, dkey)], [cat, tname, tdesc, diam]))
            else:
                appends.append([cat, tname, tdesc, diam])

        rows_to_delete = []; removed_keys = []
        for key, row in existing.items():
            if key not in current_keys:
                rows_to_delete.append(row); removed_keys.append(key)

        if updates: fxd_write_updates(sheet, headers, updates)
        removed_count = fxd_delete_rows_batched(sheet, rows_to_delete) if rows_to_delete else 0
        added_count = 0
        if appends:
            start_row = fxd_first_empty_after(region)
            added_count = fxd_write_appends(sheet, start_row, headers, appends)

        fxd_sort_data_region(sheet, headers)

        print("[FLEX DUCT] Aggiunte:", added_count)
        if appends:
            print("[FLEX DUCT] Aggiunte ({}): {}".format(min(20, len(appends)), [(r[1], r[3]) for r in appends[:20]]))
        print("[FLEX DUCT] Eliminate:", removed_count)
        if removed_keys:
            print("[FLEX DUCT] Eliminate ({}): {}".format(min(20, len(removed_keys)), removed_keys[:20]))
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
    run_pipe = form.chkPipe.Checked
    run_ins  = form.chkIns.Checked
    run_fit  = form.chkFit.Checked
    run_meq  = form.chkMeq.Checked
    run_gen  = form.chkGen.Checked
    run_duct = form.chkDuct.Checked
    run_dins = form.chkDuctIns.Checked
    run_dft = form.chkDft.Checked
    run_fxd = form.chkFxd.Checked




    if not (run_pipe or run_ins or run_fit or run_meq or run_gen or run_duct or run_dins or run_dft or run_fxd):
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

        if run_pipe:
            run_pipe_into_workbook(workbook)
        if run_ins:
            run_ins_into_workbook(workbook)
        if run_fit:
            run_fittings_into_workbook(workbook)
        if run_meq:
            run_mechanical_equipment_into_workbook(workbook)
        if run_gen:
            run_general_into_workbook(workbook)
        if run_duct:
            run_ducts_into_workbook(workbook)
        if run_dins:
            run_duct_ins_into_workbook(workbook)
        if run_dft:
            run_duct_fittings_into_workbook(workbook)
        if run_fxd:
            run_flexduct_into_workbook(workbook)




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
