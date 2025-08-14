# -*- coding: utf-8 -*-
"""
COMBO (unico Excel) con CHECKBOX UI: Cable Trays + Separators + PanelBoards + Electrical -> Excel
Un solo dialog per scegliere il file Excel
Checkbox per scegliere cosa eseguire (Cable Trays / PanelBoards / Conduit / Fixtures)
Cinque blocchi separati con funzioni rinominate per evitare collisioni
"""

__title__ = 'Revit to Excel\nELE'
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
        self.ClientSize = Size(420, 350)

        self.lbl = Label()
        self.lbl.Text = "Scegli le esportazioni da eseguire:"
        self.lbl.Location = Point(16, 16)
        self.lbl.AutoSize = True
        self.lbl.Font = Font(self.Font, FontStyle.Bold)
        self.Controls.Add(self.lbl)

        self.chkTray = CheckBox()
        self.chkTray.Text = "Passerelle (Cable Trays)"
        self.chkTray.Location = Point(20, 50)
        self.chkTray.AutoSize = True
        self.chkTray.Checked = True
        self.Controls.Add(self.chkTray)

        self.chkTraySep = CheckBox()
        self.chkTraySep.Text = "Separatore passerelle (solo canalette con MAN_Dividers > 0)"
        self.chkTraySep.Location = Point(20, 78)
        self.chkTraySep.AutoSize = True
        self.chkTraySep.Checked = True
        self.Controls.Add(self.chkTraySep)

        self.chkCond = CheckBox()
        self.chkCond.Text = "Cavidotti (Conduits – Outside Diameter)"
        self.chkCond.Location = Point(20, 106)
        self.chkCond.AutoSize = True
        self.chkCond.Checked = True
        self.Controls.Add(self.chkCond)

        self.chkEEQ = CheckBox()
        self.chkEEQ.Text = "Quadri elettrici (Electrical Equipment)"
        self.chkEEQ.Location = Point(20, 134)  # sposta se vuoi allineare meglio
        self.chkEEQ.AutoSize = True
        self.chkEEQ.Checked = True
        self.Controls.Add(self.chkEEQ)

        self.chkGen = CheckBox()
        self.chkGen.Text = "Generale (Ligting Devices / Ligting Fixtures / Electrical Fixtures)"
        self.chkGen.Location = Point(20, 162)
        self.chkGen.AutoSize = True
        self.chkGen.Checked = True
        self.Controls.Add(self.chkGen)

        self.chkPipe = CheckBox()
        self.chkPipe.Text = "Pipe (Type → Diameter)"
        self.chkPipe.Location = Point(20, 190)
        self.chkPipe.AutoSize = True
        self.chkPipe.Checked = True
        self.Controls.Add(self.chkPipe)

        self.chkFit = CheckBox()
        self.chkFit.Text = "Pipe Fittings (Family/Type → MaxSize mm)"
        self.chkFit.Location = Point(20, 218)
        self.chkFit.AutoSize = True
        self.chkFit.Checked = True
        self.Controls.Add(self.chkFit)
        
        self.chkDuct = CheckBox()
        self.chkDuct.Text = "Canali Rigidi (Ducts – Width/Height o Diameter)"
        self.chkDuct.Location = Point(20, 246)
        self.chkDuct.AutoSize = True
        self.chkDuct.Checked = True
        self.Controls.Add(self.chkDuct)

        self.chkDft = CheckBox()
        self.chkDft.Text = "Duct Fittings (Family/Type → MaxSize mm)"
        self.chkDft.Location = Point(20, 274)
        self.chkDft.AutoSize = True
        self.chkDft.Checked = True
        self.Controls.Add(self.chkDft)

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
    dlg.Title = "Seleziona il file Excel (unico per CABLE TRAY / FITTINGS / PANELBOARDS / GENERALE / CONDUIT )"
    dlg.Filter = "Excel (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls"
    dlg.Multiselect = False
    return dlg.FileName if dlg.ShowDialog() == DialogResult.OK else None

# ============================================================
# ========== BLOCCO 1 — PASSERELLE (Cable Trays → Size) =========
# ============================================================
PAS_SHEET_NAME = "Passerelle"
PAS_HEADER_ROW = 3
PAS_MIN_START_DATA_ROW = 5
PAS_HEADERS = [
    "Category",
    "Type Name",
    "MAN_TypeDescription_IT",
    "MAN_FamilyTypePrefix",
    "Size",
]
PAS_EMPTY_RUN_STOP = 20

def PAS_u(s):
    if s is None: return u""
    try: return s if isinstance(s, unicode) else unicode(s)
    except: return unicode(str(s))

def PAS_norm_text(s):
    return PAS_u(s).strip()

def PAS_norm_text_strong(s):
    return u" ".join(PAS_norm_text(s).split())

def PAS_category_name(elem):
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

_PAS_type_desc_cache = {}
def PAS_type_desc(elem):
    try:
        p = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
        if not p: return ""
        tid = p.AsElementId()
        if not tid or tid.IntegerValue <= 0: return ""
        tid_i = tid.IntegerValue
        if tid_i in _PAS_type_desc_cache: return _PAS_type_desc_cache[tid_i]
        t = doc.GetElement(tid)
        val = ""
        if t:
            pp = t.LookupParameter("MAN_TypeDescription_IT")
            if pp: val = (pp.AsString() or "") or (pp.AsValueString() or "")
        _PAS_type_desc_cache[tid_i] = val
        return val
    except:
        return ""

def PAS_type_param_text(elem, param_name):
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
        if s: return PAS_norm_text(s)
        vs = q.AsValueString()
        if vs: return PAS_norm_text(vs)
        return ""
    except:
        return ""

def PAS_instance_size_raw(elem):
    try:
        p = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE)
        if not p: return ""
        s = p.AsString()
        if s: return PAS_norm_text(s)
        vs = p.AsValueString()
        if vs: return PAS_norm_text(vs)
        return ""
    except:
        return ""

# Esempi accettati: "300 mmx104 mmϕ", "300x104", "300mm x 104mm", "300 Ø x 104"
# Output chiave: "300x104"
def PAS_size_key_and_display(raw):
    if not raw: return "", ""
    txt = PAS_u(raw)
    # rimuove simboli diametro e spazi attaccati
    txt = re.sub(u"[ΦφØø⌀ϕ]", u"", txt)
    # sostituisce separatori come ' x ' o 'X' con 'x' semplice per parsing robusto
    txt = re.sub(u"[×X]", u"x", txt)
    # prendi i primi due numeri (con eventuali decimali)
    nums = re.findall(r"(\d+(?:[.,]\d+)?)", txt)
    if len(nums) >= 2:
        a = nums[0].replace(",", ".")
        b = nums[1].replace(",", ".")
        # normalizza rimuovendo zeri inutili
        def _trim(n):
            return n.rstrip("0").rstrip(".") if "." in n else n
        a = _trim(a); b = _trim(b)
        key = u"{}x{}".format(a, b)
        disp = key
        return key, disp
    # fallback: rimuovi unità "mm", spazi
    t = PAS_norm_text(re.sub(u"[ ]*mm", u"", txt, flags=re.IGNORECASE))
    t = t.replace(" ", "")
    return t, t

def PAS_size_sort_tuple(size_key):
    # converte "300x104" -> (300.0, 104.0) per ordinamento stabile
    try:
        parts = PAS_u(size_key).split("x")
        a = float(parts[0]) if parts and parts[0] != "" else 0.0
        b = float(parts[1]) if len(parts) > 1 and parts[1] != "" else 0.0
        return (a, b)
    except:
        return (9999999.0, 9999999.0)

# ---------------------- Excel helpers -----------------------
def PAS_get_sheet_or_create(workbook, name):
    try:
        return workbook.Worksheets.Item[name]
    except:
        sh = workbook.Worksheets.Add()
        sh.Name = name
        return sh

def PAS_ensure_headers(sheet):
    last_col = sheet.Cells(PAS_HEADER_ROW, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
    if last_col < 1: last_col = 1
    headers = {}
    for c in range(1, last_col + 1):
        v = sheet.Cells(PAS_HEADER_ROW, c).Value2
        try:
            nm = v.strip() if isinstance(v, String) else None
        except:
            nm = None
        if nm: headers[nm] = c
    next_col = last_col + 1
    for h in PAS_HEADERS:
        if h not in headers:
            sheet.Cells(PAS_HEADER_ROW, next_col).Value2 = h
            headers[h] = next_col
            next_col += 1
    return headers

def PAS_read_column_block(sheet, col, r0, r1):
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
            out.append(PAS_u(val).strip())
        return out
    if isinstance(data, (tuple, list)):
        for row in data:
            try: val = row[0]
            except: val = row
            out.append(PAS_u(val).strip())
        return out
    out.append(PAS_u(data).strip())
    return out

def PAS_detect_data_region(sheet, headers):
    typ_col = headers["Type Name"]
    sz_col  = headers["Size"]
    r0 = PAS_MIN_START_DATA_ROW
    step = 2000
    curr_start = r0
    last_data_row = r0 - 1
    empty_run = 0
    max_rows = sheet.Rows.Count
    while curr_start <= max_rows:
        r1_try = min(max_rows, curr_start + step - 1)
        t_vals = PAS_read_column_block(sheet, typ_col, curr_start, r1_try)
        s_vals = PAS_read_column_block(sheet, sz_col,  curr_start, r1_try)
        block_len = max(len(t_vals), len(s_vals))
        for i in range(block_len):
            t = t_vals[i] if i < len(t_vals) else ""
            s = s_vals[i] if i < len(s_vals) else ""
            if (t or s):
                last_data_row = curr_start + i
                empty_run = 0
            else:
                empty_run += 1
                if empty_run >= PAS_EMPTY_RUN_STOP:
                    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)
        curr_start = r1_try + 1
    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)

def PAS_build_existing_index(sheet, headers):
    (r0, r1) = PAS_detect_data_region(sheet, headers)
    if r1 < r0:
        return {}, (r0, r1)
    typ_col = headers["Type Name"]; sz_col = headers["Size"]
    col_t = PAS_read_column_block(sheet, typ_col, r0, r1)
    col_s = PAS_read_column_block(sheet, sz_col,  r0, r1)
    index = {}
    n = max(len(col_t), len(col_s))
    for i in range(n):
        t = PAS_norm_text_strong(col_t[i] if i < len(col_t) else "")
        s_key, _ = PAS_size_key_and_display(col_s[i] if i < len(col_s) else "")
        if t or s_key:
            index[(t, s_key)] = r0 + i
    return index, (r0, r1)

def PAS_first_empty_row_after_region(region_tuple):
    r0, r1 = region_tuple
    if r1 < r0: return PAS_MIN_START_DATA_ROW
    return max(PAS_MIN_START_DATA_ROW, r1 + 1)

def PAS_chunk_consecutive_rows(sorted_rows):
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

def PAS_delete_rows_batched(sheet, rows_to_delete):
    if not rows_to_delete: return 0
    rows_sorted = sorted(rows_to_delete, reverse=True)
    runs = PAS_chunk_consecutive_rows(rows_sorted)
    count = 0
    for (r_start, r_end) in runs:
        rng = sheet.Range[sheet.Rows[r_start], sheet.Rows[r_end]]
        rng.EntireRow.Delete()
        count += (r_start - r_end + 1)
    return count

def PAS_write_updates_batched(sheet, headers, updates):
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
        "Size": headers["Size"],
    }
    keys = ["Category","Type Name","MAN_TypeDescription_IT","MAN_FamilyTypePrefix","Size"]
    for (r0, r1, vals_list) in runs:
        n = (r1 - r0 + 1)
        for j, key in enumerate(keys):
            col = col_map[key]
            rng = sheet.Range[sheet.Cells(r0, col), sheet.Cells(r1, col)]
            data = Array.CreateInstance(Object, n, 1)
            for i in range(n):
                v = vals_list[i][j]
                if key == "Size":
                    # sempre la versione normalizzata (es. 300x104)
                    skey, sdisp = PAS_size_key_and_display(PAS_u(v))
                    v = sdisp
                data[i,0] = PAS_u(v)
            rng.Value2 = data
    return len(updates)

def PAS_write_appends(sheet, start_row, headers, rows_data):
    if not rows_data: return 0
    if start_row < PAS_MIN_START_DATA_ROW:
        start_row = PAS_MIN_START_DATA_ROW
    cols = [headers[h] for h in PAS_HEADERS]
    n_rows = len(rows_data); n_cols = len(cols)
    data = Array.CreateInstance(Object, n_rows, n_cols)
    for i in range(n_rows):
        r = rows_data[i]
        for j in range(n_cols):
            v = r[j]
            if PAS_HEADERS[j] == "Size":
                skey, sdisp = PAS_size_key_and_display(PAS_u(v))
                v = sdisp
            data[i, j] = PAS_u(v)
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
                if PAS_HEADERS[idx] == "Size":
                    skey, sdisp = PAS_size_key_and_display(PAS_u(v))
                    v = sdisp
                col_data[r, 0] = PAS_u(v)
            col_rng.Value2 = col_data
    return n_rows

def PAS_sort_data_region(sheet, headers):
    r0 = PAS_MIN_START_DATA_ROW
    first_col = 1
    last_col = sheet.Cells(PAS_HEADER_ROW, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

    typ_col = headers["Type Name"]
    sz_col  = headers["Size"]

    last_row = 0
    for col in (typ_col, sz_col):
        lr = sheet.Cells(sheet.Rows.Count, col).End(Excel.XlDirection.xlUp).Row
        if lr > last_row: last_row = lr
    if last_row < r0: return

    data_rng = sheet.Range(sheet.Cells(r0, first_col), sheet.Cells(last_row, last_col))

    # Per ordinare "Size" numericamente su entrambe le dimensioni, scriviamo una colonna di appoggio invisibile?
    # Manteniamo sorting alfanumerico ma coerente: prima Type Name, poi Size
    sort = sheet.Sort
    sort.SortFields.Clear()
    sort.SortFields.Add(Key=sheet.Range(sheet.Cells(r0, typ_col), sheet.Cells(last_row, typ_col)),
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

def run_cable_trays_into_workbook(workbook):
    # Raccogli elementi Passerelle: usiamo la categoria "Cable Trays"
    elems = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_CableTray)\
        .WhereElementIsNotElementType()\
        .ToElements()

    groups = {}
    for e in elems:
        cat  = PAS_category_name(e) or "Cable Trays"
        tnm  = PAS_type_name(e) or ""
        desc = PAS_type_desc(e) or ""
        pref = PAS_type_param_text(e, "MAN_FamilyTypePrefix") or ""
        raw_size = PAS_instance_size_raw(e) or ""
        s_key, s_disp = PAS_size_key_and_display(raw_size)

        t_key = PAS_norm_text_strong(tnm)
        if not t_key and not s_key:
            continue

        inner = groups.get(t_key)
        if inner is None:
            inner = {}
            groups[t_key] = inner
        if s_key not in inner:
            inner[s_key] = [cat, t_key, desc, pref, s_disp]
        else:
            # completa eventuali campi mancanti
            if not inner[s_key][2]: inner[s_key][2] = desc
            if not inner[s_key][3]: inner[s_key][3] = pref

    # Ordina per Type Name e poi Size
    rows_tmp = []
    for t_key, by_size in groups.items():
        for s_key, vals in by_size.items():
            rows_tmp.append((t_key, PAS_size_sort_tuple(s_key), vals))
    rows_tmp.sort(key=lambda x: (x[0] or "", x[1]))

    ordered = []
    current_keys = set()
    for t_key, _, vals in rows_tmp:
        ordered.append(vals)
        s_key, _ = PAS_size_key_and_display(vals[4])
        current_keys.add((PAS_norm_text_strong(vals[1]), s_key))

    sheet = None
    try:
        sheet = PAS_get_sheet_or_create(workbook, PAS_SHEET_NAME)
        headers = PAS_ensure_headers(sheet)
        existing, region = PAS_build_existing_index(sheet, headers)

        updates = []; appends = []
        for cat, tname, tdesc, pref, size_disp in ordered:
            tkey = PAS_norm_text_strong(tname)
            skey, _ = PAS_size_key_and_display(size_disp)
            key = (tkey, skey)
            if key in existing:
                updates.append((existing[key], [cat, tname, tdesc, pref, size_disp]))
            else:
                appends.append([cat, tname, tdesc, pref, size_disp])

        rows_to_delete = []; removed_keys = []
        for key, row in existing.items():
            if key not in current_keys:
                rows_to_delete.append(row); removed_keys.append(key)

        if updates: PAS_write_updates_batched(sheet, headers, updates)
        removed_count = PAS_delete_rows_batched(sheet, rows_to_delete) if rows_to_delete else 0

        added_count = 0
        if appends:
            start_row = PAS_first_empty_row_after_region(region)
            added_count = PAS_write_appends(sheet, start_row, headers, appends)

        PAS_sort_data_region(sheet, headers)

        print("[PASSERELLE] Aggiunte:", added_count)
        if appends:
            preview = [(r[1], PAS_size_key_and_display(r[4])[0]) for r in appends[:20]]
            print("[PASSERELLE] Aggiunte (prime 20): {}".format(preview))
        print("[PASSERELLE] Eliminate:", removed_count)
        if removed_keys:
            print("[PASSERELLE] Eliminate (prime 20): {}".format(removed_keys[:20]))
    finally:
        try:
            if sheet: Marshal.ReleaseComObject(sheet)
        except: pass
# ============================================================
# ===== BLOCCO 2 — SEPARATORE PASSERELLE (Tray con Dividers) ===
# ============================================================
SEP_SHEET_NAME = "Separatore passerelle"
SEP_HEADER_ROW = 3
SEP_MIN_START_DATA_ROW = 5
SEP_HEADERS = [
    "Category",
    "Type Name",
    "MAN_TypeDescription_IT",
    "MAN_FamilyTypePrefix",
    "Height",  # mm
]
SEP_EMPTY_RUN_STOP = 20
SEP_KEY_MM_PREC = 6

def SEP_feet_to_mm(val_ft):
    try:
        if _HAS_UTID:
            return float(UnitUtils.ConvertFromInternalUnits(float(val_ft), UnitTypeId.Millimeters))
        else:
            return float(UnitUtils.ConvertFromInternalUnits(float(val_ft), DisplayUnitType.DUT_MILLIMETERS))
    except:
        try: return float(val_ft) * 304.8
        except: return 0.0

def SEP_to_float_mm(x):
    if x is None or x == "": return 0.0
    try: return float(x)
    except:
        try:
            return float(PAS_u(x).replace(",", ".").strip())
        except:
            m = re.search(r'(\d+(?:\.\d+)?)', PAS_u(x).replace(",", "."))
            return float(m.group(1)) if m else 0.0

def SEP_height_mm_key(elem):
    """Ritorna (key_mm, val_mm) per Height del cable tray."""
    try:
        p = elem.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
    except:
        p = None
    # 1) prova come double interno (feet)
    if p:
        try:
            ft = p.AsDouble()
            if ft and ft > 0:
                mm = SEP_feet_to_mm(ft)
                return (round(mm, SEP_KEY_MM_PREC), mm)
        except: pass
        # 2) fallback string/value string in mm
        try:
            s = (p.AsString() or p.AsValueString() or "").strip()
            if s:
                mm = SEP_to_float_mm(s)
                if mm > 0: return (round(mm, SEP_KEY_MM_PREC), mm)
        except: pass
    return (0.0, 0.0)

def SEP_dividers_ok(elem):
    """True se MAN_Dividers (istanza o tipo) è > 0 / non vuoto."""
    def _val_from_param(pp):
        if not pp: return 0
        try:
            return int(pp.AsInteger())
        except:
            s = (pp.AsString() or pp.AsValueString() or "").strip()
            if not s: return 0
            try: return int(float(s.replace(",", ".")))
            except: return 0

    # istanza
    try:
        p = elem.LookupParameter("MAN_Dividers")
        v = _val_from_param(p)
        if v > 0: return True
    except: pass

    # tipo
    try:
        tp = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
        tid = tp.AsElementId() if tp else None
        if tid and tid.IntegerValue > 0:
            t = doc.GetElement(tid)
            if t:
                p2 = t.LookupParameter("MAN_Dividers")
                v2 = _val_from_param(p2)
                if v2 > 0: return True
    except: pass

    return False

# ---------------------- Excel helpers -----------------------
def SEP_get_sheet_or_create(workbook, name):
    try:
        return workbook.Worksheets.Item[name]
    except:
        sh = workbook.Worksheets.Add()
        sh.Name = name
        return sh

def SEP_ensure_headers(sheet):
    last_col = sheet.Cells(SEP_HEADER_ROW, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
    if last_col < 1: last_col = 1
    headers = {}
    for c in range(1, last_col + 1):
        v = sheet.Cells(SEP_HEADER_ROW, c).Value2
        try:
            nm = v.strip() if isinstance(v, String) else None
        except:
            nm = None
        if nm: headers[nm] = c
    next_col = last_col + 1
    for h in SEP_HEADERS:
        if h not in headers:
            sheet.Cells(SEP_HEADER_ROW, next_col).Value2 = h
            headers[h] = next_col
            next_col += 1
    return headers

def SEP_read_column_block(sheet, col, r0, r1):
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
            out.append(PAS_u(val).strip())
        return out
    if isinstance(data, (tuple, list)):
        for row in data:
            try: val = row[0]
            except: val = row
            out.append(PAS_u(val).strip())
        return out
    out.append(PAS_u(data).strip())
    return out

def SEP_detect_data_region(sheet, headers):
    typ_col = headers["Type Name"]
    h_col   = headers["Height"]
    r0 = SEP_MIN_START_DATA_ROW
    step = 2000
    curr_start = r0
    last_data_row = r0 - 1
    empty_run = 0
    max_rows = sheet.Rows.Count
    while curr_start <= max_rows:
        r1_try = min(max_rows, curr_start + step - 1)
        t_vals = SEP_read_column_block(sheet, typ_col, curr_start, r1_try)
        h_vals = SEP_read_column_block(sheet, h_col,   curr_start, r1_try)
        block_len = max(len(t_vals), len(h_vals))
        for i in range(block_len):
            t = t_vals[i] if i < len(t_vals) else ""
            h = h_vals[i] if i < len(h_vals) else ""
            if (t or h):
                last_data_row = curr_start + i
                empty_run = 0
            else:
                empty_run += 1
                if empty_run >= SEP_EMPTY_RUN_STOP:
                    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)
        curr_start = r1_try + 1
    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)

def SEP_build_existing_index(sheet, headers):
    (r0, r1) = SEP_detect_data_region(sheet, headers)
    if r1 < r0:
        return {}, (r0, r1)
    typ_col = headers["Type Name"]; h_col = headers["Height"]
    col_t = SEP_read_column_block(sheet, typ_col, r0, r1)
    col_h = SEP_read_column_block(sheet, h_col,   r0, r1)
    index = {}
    n = max(len(col_t), len(col_h))
    for i in range(n):
        t = PAS_norm_text_strong(col_t[i] if i < len(col_t) else "")
        h_raw = col_h[i] if i < len(col_h) else ""
        h_val = round(SEP_to_float_mm(h_raw), SEP_KEY_MM_PREC)
        if t or (h_val != 0.0):
            index[(t, h_val)] = r0 + i
    return index, (r0, r1)

def SEP_first_empty_row_after_region(region_tuple):
    r0, r1 = region_tuple
    if r1 < r0: return SEP_MIN_START_DATA_ROW
    return max(SEP_MIN_START_DATA_ROW, r1 + 1)

def SEP_chunk_consecutive_rows(sorted_rows):
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

def SEP_delete_rows_batched(sheet, rows_to_delete):
    if not rows_to_delete: return 0
    rows_sorted = sorted(rows_to_delete, reverse=True)
    runs = SEP_chunk_consecutive_rows(rows_sorted)
    count = 0
    for (r_start, r_end) in runs:
        rng = sheet.Range[sheet.Rows[r_start], sheet.Rows[r_end]]
        rng.EntireRow.Delete()
        count += (r_start - r_end + 1)
    return count

def SEP_write_updates_batched(sheet, headers, updates):
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
        "Height": headers["Height"],
    }
    keys = ["Category","Type Name","MAN_TypeDescription_IT","MAN_FamilyTypePrefix","Height"]
    for (r0, r1, vals_list) in runs:
        n = (r1 - r0 + 1)
        for j, key in enumerate(keys):
            col = col_map[key]
            rng = sheet.Range[sheet.Cells(r0, col), sheet.Cells(r1, col)]
            data = Array.CreateInstance(Object, n, 1)
            for i in range(n):
                v = vals_list[i][j]
                if key == "Height":
                    v = float(SEP_to_float_mm(v))
                data[i,0] = v
            rng.Value2 = data
    return len(updates)

def SEP_write_appends(sheet, start_row, headers, rows_data):
    if not rows_data: return 0
    if start_row < SEP_MIN_START_DATA_ROW:
        start_row = SEP_MIN_START_DATA_ROW
    cols = [headers[h] for h in SEP_HEADERS]
    n_rows = len(rows_data); n_cols = len(cols)
    data = Array.CreateInstance(Object, n_rows, n_cols)
    for i in range(n_rows):
        r = rows_data[i]
        for j in range(n_cols):
            v = r[j]
            if SEP_HEADERS[j] == "Height":
                v = float(SEP_to_float_mm(v))
            data[i, j] = v
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
                v = rows_data[r][idx]
                if SEP_HEADERS[idx] == "Height":
                    v = float(SEP_to_float_mm(v))
                col_data[r, 0] = v
            col_rng.Value2 = col_data
    return n_rows

def SEP_sort_data_region(sheet, headers):
    r0 = SEP_MIN_START_DATA_ROW
    first_col = 1
    last_col = sheet.Cells(SEP_HEADER_ROW, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

    t_col = headers["Type Name"]
    h_col = headers["Height"]

    last_row = 0
    for col in (t_col, h_col):
        lr = sheet.Cells(sheet.Rows.Count, col).End(Excel.XlDirection.xlUp).Row
        if lr > last_row: last_row = lr
    if last_row < r0: return

    data_rng = sheet.Range(sheet.Cells(r0, first_col), sheet.Cells(last_row, last_col))

    sort = sheet.Sort
    sort.SortFields.Clear()
    sort.SortFields.Add(Key=sheet.Range(sheet.Cells(r0, t_col), sheet.Cells(last_row, t_col)),
                        SortOn=Excel.XlSortOn.xlSortOnValues, Order=Excel.XlSortOrder.xlAscending,
                        DataOption=Excel.XlSortDataOption.xlSortNormal)
    sort.SortFields.Add(Key=sheet.Range(sheet.Cells(r0, h_col), sheet.Cells(last_row, h_col)),
                        SortOn=Excel.XlSortOn.xlSortOnValues, Order=Excel.XlSortOrder.xlAscending,
                        DataOption=Excel.XlSortDataOption.xlSortNormal)
    sort.SetRange(data_rng)
    sort.Header = Excel.XlYesNoGuess.xlNo
    sort.MatchCase = False
    sort.Orientation = Excel.XlSortOrientation.xlSortColumns  # coerente con gli altri blocchi
    sort.Apply()

def run_cable_tray_separators_into_workbook(workbook):
    # prendi solo Cable Trays con MAN_Dividers > 0
    elems = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_CableTray) \
        .WhereElementIsNotElementType() \
        .ToElements()

    groups = {}
    for e in elems:
        if not SEP_dividers_ok(e):
            continue
        cat  = PAS_category_name(e) or "Cable Trays"
        tnm  = PAS_type_name(e) or ""
        desc = PAS_type_desc(e) or ""
        pref = PAS_type_param_text(e, "MAN_FamilyTypePrefix") or ""

        h_key, h_val = SEP_height_mm_key(e)
        if h_key <= 0:
            continue

        t_key = PAS_norm_text_strong(tnm)
        if not t_key:
            continue

        inner = groups.get(t_key)
        if inner is None:
            inner = {}; groups[t_key] = inner
        if h_key not in inner:
            inner[h_key] = [cat, t_key, desc, pref, h_key]
        else:
            # completa eventuali campi mancanti
            if not inner[h_key][2]: inner[h_key][2] = desc
            if not inner[h_key][3]: inner[h_key][3] = pref

    # Ordina per Type Name, poi per Height (mm)
    rows_tmp = []
    for t_key, by_h in groups.items():
        for h_key, vals in by_h.items():
            rows_tmp.append((t_key, float(h_key), vals))
    rows_tmp.sort(key=lambda x: (x[0] or "", x[1]))

    ordered = []
    current_keys = set()
    for t_key, h_val, vals in rows_tmp:
        ordered.append(vals)
        current_keys.add((PAS_norm_text_strong(vals[1]), round(float(vals[4]), SEP_KEY_MM_PREC)))

    sheet = None
    try:
        sheet = SEP_get_sheet_or_create(workbook, SEP_SHEET_NAME)
        headers = SEP_ensure_headers(sheet)
        existing, region = SEP_build_existing_index(sheet, headers)

        updates = []; appends = []
        for cat, tname, tdesc, pref, h_mm in ordered:
            tkey = PAS_norm_text_strong(tname)
            hkey = round(float(h_mm), SEP_KEY_MM_PREC)
            key = (tkey, hkey)
            row_vals = [cat, tname, tdesc, pref, hkey]
            if key in existing:
                updates.append((existing[key], row_vals))
            else:
                appends.append(row_vals)

        rows_to_delete = []; removed_keys = []
        for key, row in existing.items():
            if key not in current_keys:
                rows_to_delete.append(row); removed_keys.append(key)

        if updates: SEP_write_updates_batched(sheet, headers, updates)
        removed_count = SEP_delete_rows_batched(sheet, rows_to_delete) if rows_to_delete else 0

        added_count = 0
        if appends:
            start_row = SEP_first_empty_row_after_region(region)
            added_count = SEP_write_appends(sheet, start_row, headers, appends)

        SEP_sort_data_region(sheet, headers)

        print("[SEP PASSERELLE] Aggiunte:", added_count)
        if appends:
            preview = [(r[1], round(float(r[4]), 3)) for r in appends[:20]]
            print("[SEP PASSERELLE] Aggiunte (prime 20): {}".format(preview))
        print("[SEP PASSERELLE] Eliminate:", removed_count)
        if removed_keys:
            print("[SEP PASSERELLE] Eliminate (prime 20): {}".format(removed_keys[:20]))
    finally:
        try:
            if sheet: Marshal.ReleaseComObject(sheet)
        except: pass

# ============================================================
# ============ BLOCCO 3 — CAVIDOTTI (Conduits) =================
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
        try: return float(PAS_u(x).replace(",", ".").strip())
        except:
            m = re.search(r'(\d+(?:\.\d+)?)', PAS_u(x).replace(",", "."))
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
            out.append(PAS_u(val).strip())
        return out
    if isinstance(data, (tuple, list)):
        for row in data:
            try: val = row[0]
            except: val = row
            out.append(PAS_u(val).strip())
        return out
    out.append(PAS_u(data).strip())
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
        t = PAS_norm_text_strong(col_t[i] if i < len(col_t) else "")
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

def run_conduits_into_workbook(workbook):
    # Raccogli elementi Conduit (Cavidotti)
    elems = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_Conduit) \
        .WhereElementIsNotElementType() \
        .ToElements()

    groups = {}
    for e in elems:
        cat  = PAS_category_name(e) or "Conduits"
        tnm  = PAS_type_name(e) or ""
        if not tnm: continue
        desc = PAS_type_desc(e) or ""
        pref = PAS_type_param_text(e, "MAN_FamilyTypePrefix") or ""

        d_key, d_val = COND_outside_diam_mm_key(e)
        if d_key <= 0:  # ignora senza diametro
            continue

        t_key = PAS_norm_text_strong(tnm)
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
        current_keys.add((PAS_norm_text_strong(vals[1]), round(float(vals[4]), COND_KEY_MM_PREC)))

    sheet = None
    try:
        sheet = COND_get_sheet_or_create(workbook, COND_SHEET_NAME)
        headers = COND_ensure_headers(sheet)
        existing, region = COND_build_existing_index(sheet, headers)

        updates = []; appends = []
        for cat, tname, tdesc, pref, d_mm in ordered:
            tkey = PAS_norm_text_strong(tname)
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

        print("[CAVIDOTTI] Aggiunte:", added_count)
        if appends:
            preview = [(r[1], round(float(r[4]), 3)) for r in appends[:20]]
            print("[CAVIDOTTI] Aggiunte (prime 20): {}".format(preview))
        print("[CAVIDOTTI] Eliminate:", removed_count)
        if removed_keys:
            print("[CAVIDOTTI] Eliminate (prime 20): {}".format(removed_keys[:20]))
    finally:
        try:
            if sheet: Marshal.ReleaseComObject(sheet)
        except: pass

# ============================================================
# ===== BLOCCO 4 — QUADRI ELETTRICI (Electrical Equipment) =====
# ============================================================
EEQ_SHEET_NAME = "Quadri elettrici"
EEQ_HEADER_ROW = 3
EEQ_MIN_START_DATA_ROW = 5
EEQ_HEADERS = [
    "Category",
    "Family Name",
    "Type Name",
    "MAN_TypeDescription_IT",
    "Panel Name",
    "Level",
]
EEQ_EMPTY_RUN_STOP = 20

def EEQ_u(s):
    if s is None: return u""
    try: return s if isinstance(s, unicode) else unicode(s)
    except: return unicode(str(s))

def EEQ_norm_text(s):
    return EEQ_u(s).strip()

def EEQ_norm_text_strong(s):
    return u" ".join(EEQ_norm_text(s).split())

def EEQ_category_name(elem):
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

# cache per MAN_TypeDescription_IT
_EEQ_type_desc_cache = {}
def EEQ_type_desc(elem):
    try:
        p = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
        if not p: return ""
        tid = p.AsElementId()
        if not tid or tid.IntegerValue <= 0: return ""
        tid_i = tid.IntegerValue
        if tid_i in _EEQ_type_desc_cache: return _EEQ_type_desc_cache[tid_i]
        t = doc.GetElement(tid)
        val = ""
        if t:
            pp = t.LookupParameter("MAN_TypeDescription_IT")
            if pp: val = (pp.AsString() or "") or (pp.AsValueString() or "")
        _EEQ_type_desc_cache[tid_i] = val
        return val
    except:
        return ""

def EEQ_panel_name(elem):
    # Parametro istanza "Panel Name"
    try:
        p = elem.LookupParameter("Panel Name")
        if p:
            s = p.AsString()
            if s: return EEQ_norm_text(s)
            vs = p.AsValueString()
            if vs: return EEQ_norm_text(vs)
    except: pass
    return ""

def EEQ_level_name(elem):
    # FAMILY_LEVEL_PARAM -> Level element name
    try:
        p = elem.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
        if p:
            try:
                lid = p.AsElementId()
                if lid and lid.IntegerValue > 0:
                    lv = doc.GetElement(lid)
                    if lv and hasattr(lv, "Name") and lv.Name:
                        return EEQ_norm_text(lv.Name)
            except: pass
            vs = p.AsValueString()
            if vs: return EEQ_norm_text(vs)
    except: pass
    return ""

# ---------------------- Excel helpers -----------------------
def EEQ_get_sheet_or_create(workbook, name):
    try:
        return workbook.Worksheets.Item[name]
    except:
        sh = workbook.Worksheets.Add()
        sh.Name = name
        return sh

def EEQ_ensure_headers(sheet):
    last_col = sheet.Cells(EEQ_HEADER_ROW, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
    if last_col < 1: last_col = 1
    headers = {}
    for c in range(1, last_col + 1):
        v = sheet.Cells(EEQ_HEADER_ROW, c).Value2
        try:
            nm = v.strip() if isinstance(v, String) else None
        except:
            nm = None
        if nm: headers[nm] = c
    next_col = last_col + 1
    for h in EEQ_HEADERS:
        if h not in headers:
            sheet.Cells(EEQ_HEADER_ROW, next_col).Value2 = h
            headers[h] = next_col
            next_col += 1
    return headers

def EEQ_read_column_block(sheet, col, r0, r1):
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
            out.append(EEQ_u(val).strip())
        return out
    if isinstance(data, (tuple, list)):
        for row in data:
            try: val = row[0]
            except: val = row
            out.append(EEQ_u(val).strip())
        return out
    out.append(EEQ_u(data).strip())
    return out

def EEQ_detect_data_region(sheet, headers):
    fam_col = headers["Family Name"]
    typ_col = headers["Type Name"]
    pnl_col = headers["Panel Name"]
    r0 = EEQ_MIN_START_DATA_ROW
    step = 2000
    curr_start = r0
    last_data_row = r0 - 1
    empty_run = 0
    max_rows = sheet.Rows.Count
    while curr_start <= max_rows:
        r1_try = min(max_rows, curr_start + step - 1)
        f_vals = EEQ_read_column_block(sheet, fam_col, curr_start, r1_try)
        t_vals = EEQ_read_column_block(sheet, typ_col, curr_start, r1_try)
        p_vals = EEQ_read_column_block(sheet, pnl_col, curr_start, r1_try)
        block_len = max(len(f_vals), len(t_vals), len(p_vals))
        for i in range(block_len):
            f = f_vals[i] if i < len(f_vals) else ""
            t = t_vals[i] if i < len(t_vals) else ""
            p = p_vals[i] if i < len(p_vals) else ""
            if (f or t or p):
                last_data_row = curr_start + i
                empty_run = 0
            else:
                empty_run += 1
                if empty_run >= EEQ_EMPTY_RUN_STOP:
                    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)
        curr_start = r1_try + 1
    return (r0, last_data_row) if last_data_row >= r0 else (r0, r0-1)

def EEQ_build_existing_index(sheet, headers):
    (r0, r1) = EEQ_detect_data_region(sheet, headers)
    if r1 < r0:
        return {}, (r0, r1)
    fam_col = headers["Family Name"]; typ_col = headers["Type Name"]
    lvl_col = headers["Level"]; pnl_col = headers["Panel Name"]
    col_f = EEQ_read_column_block(sheet, fam_col, r0, r1)
    col_t = EEQ_read_column_block(sheet, typ_col, r0, r1)
    col_l = EEQ_read_column_block(sheet, lvl_col, r0, r1)
    col_p = EEQ_read_column_block(sheet, pnl_col, r0, r1)
    index = {}
    n = max(len(col_f), len(col_t), len(col_l), len(col_p))
    for i in range(n):
        f = EEQ_norm_text_strong(col_f[i] if i < len(col_f) else "")
        t = EEQ_norm_text_strong(col_t[i] if i < len(col_t) else "")
        l = EEQ_norm_text_strong(col_l[i] if i < len(col_l) else "")
        p = EEQ_norm_text_strong(col_p[i] if i < len(col_p) else "")
        if f or t or l or p:
            index[(f, t, l, p)] = r0 + i
    return index, (r0, r1)

def EEQ_first_empty_row_after_region(region_tuple):
    r0, r1 = region_tuple
    if r1 < r0: return EEQ_MIN_START_DATA_ROW
    return max(EEQ_MIN_START_DATA_ROW, r1 + 1)

def EEQ_chunk_consecutive_rows(sorted_rows):
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

def EEQ_delete_rows_batched(sheet, rows_to_delete):
    if not rows_to_delete: return 0
    rows_sorted = sorted(rows_to_delete, reverse=True)
    runs = EEQ_chunk_consecutive_rows(rows_sorted)
    count = 0
    for (r_start, r_end) in runs:
        rng = sheet.Range[sheet.Rows[r_start], sheet.Rows[r_end]]
        rng.EntireRow.Delete()
        count += (r_start - r_end + 1)
    return count

def EEQ_write_updates_batched(sheet, headers, updates):
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
        "Panel Name": headers["Panel Name"],
        "Level": headers["Level"],
    }
    keys = ["Category","Family Name","Type Name","MAN_TypeDescription_IT","Panel Name","Level"]
    for (r0, r1, vals_list) in runs:
        n = (r1 - r0 + 1)
        for j, key in enumerate(keys):
            col = col_map[key]
            rng = sheet.Range[sheet.Cells(r0, col), sheet.Cells(r1, col)]
            data = Array.CreateInstance(Object, n, 1)
            for i in range(n):
                data[i,0] = EEQ_u(vals_list[i][j])
            rng.Value2 = data
    return len(updates)

def EEQ_write_appends(sheet, start_row, headers, rows_data):
    if not rows_data: return 0
    if start_row < EEQ_MIN_START_DATA_ROW:
        start_row = EEQ_MIN_START_DATA_ROW
    cols = [headers[h] for h in EEQ_HEADERS]
    n_rows = len(rows_data); n_cols = len(cols)
    data = Array.CreateInstance(Object, n_rows, n_cols)
    for i in range(n_rows):
        r = rows_data[i]
        for j in range(n_cols):
            data[i, j] = EEQ_u(r[j])
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
                col_data[r, 0] = EEQ_u(rows_data[r][idx])
            col_rng.Value2 = col_data
    return n_rows

def EEQ_sort_data_region(sheet, headers):
    r0 = EEQ_MIN_START_DATA_ROW
    first_col = 1
    last_col = sheet.Cells(EEQ_HEADER_ROW, sheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

    fam_col = headers["Family Name"]
    typ_col = headers["Type Name"]
    lvl_col = headers["Level"]
    pnl_col = headers["Panel Name"]

    last_row = 0
    for col in (fam_col, typ_col, lvl_col, pnl_col):
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
    sort.SortFields.Add(Key=sheet.Range(sheet.Cells(r0, lvl_col), sheet.Cells(last_row, lvl_col)),
                        SortOn=Excel.XlSortOn.xlSortOnValues, Order=Excel.XlSortOrder.xlAscending,
                        DataOption=Excel.XlSortDataOption.xlSortNormal)
    sort.SortFields.Add(Key=sheet.Range(sheet.Cells(r0, pnl_col), sheet.Cells(last_row, pnl_col)),
                        SortOn=Excel.XlSortOn.xlSortOnValues, Order=Excel.XlSortOrder.xlAscending,
                        DataOption=Excel.XlSortDataOption.xlSortNormal)
    sort.SetRange(data_rng)
    sort.Header = Excel.XlYesNoGuess.xlNo
    sort.MatchCase = False
    sort.Orientation = Excel.XlSortOrientation.xlSortColumns
    sort.Apply()

def run_electrical_equipment_into_workbook(workbook):
    # Raccoglie solo gli Electrical Equipment con Family Name che inizia per "MAN_EEQ_PNB_SwitchBoard"
    elems = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_ElectricalEquipment) \
        .WhereElementIsNotElementType() \
        .ToElements()

    groups = {}
    for e in elems:
        fam  = EEQ_family_name(e) or ""
        if not fam or not fam.startswith("MAN_EEQ_PNB_SwitchBoard"):
            continue  # prefiltraggio fondamentale

        cat  = EEQ_category_name(e) or "Electrical Equipment"
        typ  = EEQ_type_name(e) or ""
        desc = EEQ_type_desc(e) or ""
        pnl  = EEQ_panel_name(e) or ""
        lvl  = EEQ_level_name(e) or ""

        fam_k = EEQ_norm_text_strong(fam)
        typ_k = EEQ_norm_text_strong(typ)
        lvl_k = EEQ_norm_text_strong(lvl)
        pnl_k = EEQ_norm_text_strong(pnl)

        key = (fam_k, typ_k, lvl_k, pnl_k)
        if key not in groups:
            groups[key] = [cat, fam_k, typ_k, desc, pnl_k, lvl_k]
        else:
            if not groups[key][3]:
                groups[key][3] = desc  # completa descrizione se mancante

    # Ordina per Family Name, Type Name, Level, Panel Name
    rows_tmp = []
    for (fam_k, typ_k, lvl_k, pnl_k), vals in groups.items():
        rows_tmp.append((fam_k, typ_k, lvl_k, pnl_k, vals))
    rows_tmp.sort(key=lambda x: (x[0] or "", x[1] or "", x[2] or "", x[3] or ""))

    ordered = []
    current_keys = set()
    for fam_k, typ_k, lvl_k, pnl_k, vals in rows_tmp:
        ordered.append(vals)
        current_keys.add((fam_k, typ_k, lvl_k, pnl_k))

    sheet = None
    try:
        sheet = EEQ_get_sheet_or_create(workbook, EEQ_SHEET_NAME)
        headers = EEQ_ensure_headers(sheet)
        existing, region = EEQ_build_existing_index(sheet, headers)

        updates = []; appends = []
        for cat, fam, typ, desc, pnl, lvl in ordered:
            key = (EEQ_norm_text_strong(fam), EEQ_norm_text_strong(typ),
                   EEQ_norm_text_strong(lvl), EEQ_norm_text_strong(pnl))
            row_vals = [cat, fam, typ, desc, pnl, lvl]
            if key in existing:
                updates.append((existing[key], row_vals))
            else:
                appends.append(row_vals)

        rows_to_delete = []; removed_keys = []
        for key, row in existing.items():
            if key not in current_keys:
                rows_to_delete.append(row); removed_keys.append(key)

        if updates: EEQ_write_updates_batched(sheet, headers, updates)
        removed_count = EEQ_delete_rows_batched(sheet, rows_to_delete) if rows_to_delete else 0

        added_count = 0
        if appends:
            start_row = EEQ_first_empty_row_after_region(region)
            added_count = EEQ_write_appends(sheet, start_row, headers, appends)

        EEQ_sort_data_region(sheet, headers)

        print("[QUADRI ELETTRICI] Aggiunte:", added_count)
        if appends:
            preview = [(r[1], r[2], r[5], r[4]) for r in appends[:20]]  # fam, type, level, panel
            print("[QUADRI ELETTRICI] Aggiunte (prime 20): {}".format(preview))
        print("[QUADRI ELETTRICI] Eliminate:", removed_count)
        if removed_keys:
            print("[QUADRI ELETTRICI] Eliminate (prime 20): {}".format(removed_keys[:20]))
    finally:
        try:
            if sheet: Marshal.ReleaseComObject(sheet)
        except: pass


# ============================================================
# ========== BLOCCO 5 — GENERALE (EF / LF / Fixtures) ======
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
    for bic in (BuiltInCategory.OST_CableTrayFitting,
                BuiltInCategory.OST_ElectricalEquipment,
                BuiltInCategory.OST_ElectricalFixtures,
                BuiltInCategory.OST_LightingDevices,
                BuiltInCategory.OST_LightingFixtures):
        elems.extend(
            list(FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType().ToElements())
        )

    groups = {}
    for e in elems:
        if not isinstance(e, FamilyInstance):
            continue
        try:
            if e.Category and e.Category.Id.IntegerValue == int(BuiltInCategory.OST_ElectricalEquipment):
                fam_raw = GEN_family_name(e) or ""
                fam_s = fam_raw.strip()
                if fam_s.startswith("MAN_EEQ_PNB_SwitchBoard") or fam_s.startswith("MAN_SEQ"):
                    continue
        except:
            pass
        
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
# =============== BLOCCO 6 — PIPE (Type→Diameter) ============
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
# ====== BLOCCO 7 — PIPE FITTINGS (Family/Type → MaxSize) ====
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
# ======== BLOCCO 8 — CANALI RIGIDI (Type → Max Dim) =========
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
# ===== BLOCCO 9 — DUCT FITTINGS (Family/Type → MaxSize) =====
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
# ========================= MAIN =============================
# ============================================================
def main():
    # 1) scegli cosa eseguire
    form = RunPickerForm()
    dr = form.ShowDialog()
    if dr != DialogResult.OK:
        return
    run_tray = form.chkTray.Checked
    run_tray_sep = form.chkTraySep.Checked
    run_cond = form.chkCond.Checked
    run_eeq  = form.chkEEQ.Checked
    run_gen  = form.chkGen.Checked
    run_pipe = form.chkPipe.Checked
    run_fit  = form.chkFit.Checked
    run_duct = form.chkDuct.Checked
    run_dft = form.chkDft.Checked


    if not (run_tray or run_tray_sep or run_cond or run_eeq or run_pipe or run_fit or run_gen or run_duct or run_dft):
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
        if run_tray:
            run_cable_trays_into_workbook(workbook)
        if run_tray_sep:
            run_cable_tray_separators_into_workbook(workbook)
        if run_cond:
            run_conduits_into_workbook(workbook)
        if run_eeq:
            run_electrical_equipment_into_workbook(workbook)
        if run_gen:
            run_general_into_workbook(workbook)
        if run_pipe:
            run_pipe_into_workbook(workbook)
        if run_fit:
            run_fittings_into_workbook(workbook)
        if run_duct:
            run_ducts_into_workbook(workbook)
        if run_dft:
            run_duct_fittings_into_workbook(workbook)




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
