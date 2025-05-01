# -*- coding: utf-8 -*-
__title__ = "Geberit"
__doc__ = """Version = 1.0
Date    = 12.04.2025
________________________________________________________________
Description:

This integrated script lets you:
  1. Select boundary detail lines that form a closed loop.
  2. Gather all elements in the active view whose bounding-box center is inside the boundary.
  3. Filter the gathered elements to only those in the categories of interest 
     (Pipes, Pipe Fittings, Pipe Tags, Text Notes).
  4. Display an editable grid so you can adjust the associated prefab (Comments) codes.
  5. Use the Text Note code as the base number.
       • For Pipes and Pipe Tags, the NewCode becomes: [Base].1, [Base].2, … (sorted left-to-right, bottom-to-up).
       • For Pipe Fittings, the NewCode is simply [Base].

If a Text Note is missing from the selected region, you have the option to place one.
________________________________________________________________
How-To:

1. [Hold ALT + CLICK] on the button to open its source folder.
2. Create your boundary (with detail lines) in the view.
3. Click the button and follow the prompts.
________________________________________________________________
Author: Emin Avdovic"""

# ==================================================
# Imports
# ==================================================
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import (
    FamilySymbol,
    FilteredElementCollector,
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    XYZ,
    Transaction,
    TextNote,
    TextNoteType,
    TextNoteOptions,
    IndependentTag,
    UV,
    Reference,
    TagMode,
    TagOrientation,
    ViewSheet,
    ViewDuplicateOption,
    ViewDiscipline,
)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.UI import UIDocument
from Autodesk.Revit.Exceptions import ArgumentException
from System.Collections.Generic import List

import clr
import System
import System.IO

clr.AddReference("System")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("RevitServices")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("WindowsBase")
from RevitServices.Persistence import DocumentManager
from System.Windows.Forms import (
    Form,
    ComboBox,
    ListBox,
    PictureBox,
    PictureBoxSizeMode,
    DataGridView,
    DataGridViewTextBoxColumn,
    DataGridViewAutoSizeColumnsMode,
    TextBox,
    Button,
    MessageBox,
    DialogResult,
    Label,
    ScrollBars,
    Application,
)
from System.Drawing import Image, Point, Color, Rectangle, Size
from System.IO import MemoryStream
from System.Windows.Forms import DataGridViewButtonColumn
from System import Array
import math, re, sys

# ==================================================
# Revit Document Setup
# ==================================================
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

# ==================================================
# Helper Functions
# ==================================================


# --- Boundary Selection Functions ---
class DetailLineSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        if elem.Category and elem.Category.Id.IntegerValue == int(
            BuiltInCategory.OST_Lines
        ):
            return True
        return False

    def AllowReference(self, ref, point):
        return False


def points_are_close(pt1, pt2, tol=1e-6):
    return (
        abs(pt1.X - pt2.X) < tol
        and abs(pt1.Y - pt2.Y) < tol
        and abs(pt1.Z - pt2.Z) < tol
    )


def order_segments_to_polygon(segments):
    if not segments:
        return None
    polygon = [segments[0][0], segments[0][1]]
    segments.pop(0)
    changed = True
    while segments and changed:
        changed = False
        last_pt = polygon[-1]
        for idx, seg in enumerate(segments):
            ptA, ptB = seg
            if points_are_close(last_pt, ptA):
                polygon.append(ptB)
                segments.pop(idx)
                changed = True
                break
            elif points_are_close(last_pt, ptB):
                polygon.append(ptA)
                segments.pop(idx)
                changed = True
                break
    if polygon and points_are_close(polygon[0], polygon[-1]):
        polygon.pop()
        return polygon
    else:
        return None


def is_point_inside_polygon(point, polygon):
    x = point.X
    y = point.Y
    inside = False
    n = len(polygon)
    j = n - 1
    for i in range(n):
        xi = polygon[i].X
        yi = polygon[i].Y
        xj = polygon[j].X
        yj = polygon[j].Y
        if ((yi > y) != (yj > y)) and (
            x < (xj - xi) * (y - yi) / ((yj - yi) or 1e-12) + xi
        ):
            inside = not inside
        j = i
    return inside


def select_boundary_and_gather():
    try:
        selection_refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            DetailLineSelectionFilter(),
            "Select boundary detail lines (click on the lines that form a closed loop)",
        )
    except Exception:
        return None
    if not selection_refs:
        return None

    segments = []
    for ref in selection_refs:
        elem = doc.GetElement(ref)
        try:
            curve = elem.GeometryCurve
            start = curve.GetEndPoint(0)
            end = curve.GetEndPoint(1)
            segments.append((start, end))
        except Exception:
            continue

    polygon = order_segments_to_polygon(segments[:])
    if polygon is None:
        MessageBox.Show(
            "The selected detail lines do not form a closed boundary.", "Error"
        )
        return None

    collector = (
        FilteredElementCollector(doc, uidoc.ActiveView.Id)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    elements_inside = []
    for elem in collector:
        bbox = elem.get_BoundingBox(uidoc.ActiveView)
        if bbox:
            center = XYZ(
                (bbox.Min.X + bbox.Max.X) / 2.0,
                (bbox.Min.Y + bbox.Max.Y) / 2.0,
                (bbox.Min.Z + bbox.Max.Z) / 2.0,
            )
            if is_point_inside_polygon(center, polygon):
                elements_inside.append(elem)

    MessageBox.Show(
        "Found {0} element(s) inside the selected boundary.".format(
            len(elements_inside)
        ),
        "Boundary Selection",
    )
    return elements_inside


# --- Parameter and Region Helpers ---
def convert_param_to_string(param_obj):
    if not param_obj:
        return ""
    try:
        val_str = param_obj.AsValueString()
        if val_str and val_str.strip() != "":
            return val_str
    except Exception:
        pass
    try:
        val_double = param_obj.AsDouble()
        val_mm = val_double * 304.8
        return str(int(round(val_mm))) + " mm"
    except Exception:
        return ""


def get_region_bounding_box(elements):
    valid_found = False
    overall_min_x = float("inf")
    overall_min_y = float("inf")
    overall_min_z = float("inf")
    overall_max_x = float("-inf")
    overall_max_y = float("-inf")
    overall_max_z = float("-inf")

    for el in elements:
        try:
            bbox = el.get_BoundingBox(uidoc.ActiveView)
        except:
            continue  # skip if element was just deleted
        if not bbox:
            continue
        if bbox is None:
            continue
        if (
            bbox.Min.X == float("inf")
            or bbox.Min.Y == float("inf")
            or bbox.Min.Z == float("inf")
        ):
            continue
        valid_found = True
        overall_min_x = min(overall_min_x, bbox.Min.X)
        overall_min_y = min(overall_min_y, bbox.Min.Y)
        overall_min_z = min(overall_min_z, bbox.Min.Z)
        overall_max_x = max(overall_max_x, bbox.Max.X)
        overall_max_y = max(overall_max_y, bbox.Max.Y)
        overall_max_z = max(overall_max_z, bbox.Max.Z)

    if not valid_found:
        return XYZ(0, 0, 0), XYZ(0, 0, 0)

    overall_min = XYZ(overall_min_x, overall_min_y, overall_min_z)
    overall_max = XYZ(overall_max_x, overall_max_y, overall_max_z)
    return overall_min, overall_max


def create_pipe_tags_for_untagged_pipes(doc, pipes, view):
    t = Transaction(doc, "Add Missing Pipe Tags")
    t.Start()
    for pipe in pipes:
        bbox = pipe.get_BoundingBox(view)
        if not bbox:
            continue
        center = XYZ(
            (bbox.Min.X + bbox.Max.X) / 2.0,
            (bbox.Min.Y + bbox.Max.Y) / 2.0,
            (bbox.Min.Z + bbox.Max.Z) / 2.0,
        )
        pipe_ref = Reference(pipe)
        IndependentTag.Create(
            doc,
            view.Id,
            pipe_ref,
            True,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            UV(center.X, center.Y),
        )
    t.Commit()


# ==================================================
# UI Class: ElementEditorForm
# ==================================================
class ElementEditorForm(Form):
    def __init__(self, elements_data, region_elements=None):
        self.Text = "Edit Element Codes"
        self.Width = 950
        self.Height = 500
        self.regionElements = region_elements

        self.dataGrid = DataGridView()
        self.dataGrid.SelectionChanged += self.on_row_selected
        self.dataGrid.Location = Point(10, 10)
        self.dataGrid.Size = Size(900, 350)
        self.dataGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        self.dataGrid.CellContentClick += self.dataGrid_CellContentClick
        self.Controls.Add(self.dataGrid)

        # Columns
        self.colId = DataGridViewTextBoxColumn()
        self.colId.Name = "Id"
        self.colId.HeaderText = "Element Id"
        self.colId.ReadOnly = True
        self.colCategory = DataGridViewTextBoxColumn()
        self.colCategory.Name = "Category"
        self.colCategory.HeaderText = "Category"
        self.colCategory.ReadOnly = True
        self.colName = DataGridViewTextBoxColumn()
        self.colName.Name = "Name"
        self.colName.HeaderText = "Name"
        self.colName.ReadOnly = True
        self.colArticle = DataGridViewTextBoxColumn()
        self.colArticle.Name = "GEB_Article_Number"
        self.colArticle.HeaderText = "GEB Article No."
        self.colArticle.ReadOnly = True
        self.colDefaultCode = DataGridViewTextBoxColumn()
        self.colDefaultCode.Name = "DefaultCode"
        self.colDefaultCode.HeaderText = "Default Code"
        self.colDefaultCode.ReadOnly = True
        self.colNewCode = DataGridViewTextBoxColumn()
        self.colNewCode.Name = "NewCode"
        self.colNewCode.HeaderText = "New Code"
        self.colNewCode.ReadOnly = False
        self.colOD = DataGridViewTextBoxColumn()
        self.colOD.Name = "OutsideDiameter"
        self.colOD.HeaderText = "Outside Diameter"
        self.colOD.ReadOnly = True
        self.colLength = DataGridViewTextBoxColumn()
        self.colLength.Name = "Length"
        self.colLength.HeaderText = "Length"
        self.colLength.ReadOnly = True
        self.colSize = DataGridViewTextBoxColumn()
        self.colSize.Name = "Size"
        self.colSize.HeaderText = "Size"
        self.colSize.ReadOnly = True

        self.colTagStatus = DataGridViewButtonColumn()
        self.colTagStatus.Name = "TagStatus"
        self.colTagStatus.HeaderText = "Tags"
        self.colTagStatus.UseColumnTextForButtonValue = False

        # Place Text Note
        self.btnPlaceTextNote = Button()
        self.btnPlaceTextNote.Text = "Place Text Note"
        self.btnPlaceTextNote.Location = Point(10, 370)
        self.btnPlaceTextNote.Width = 150
        self.btnPlaceTextNote.Click += self.btnPlaceTextNote_Click
        self.Controls.Add(self.btnPlaceTextNote)

        self.txtTextNoteCode = TextBox()
        self.txtTextNoteCode.Location = Point(170, 370)
        self.txtTextNoteCode.Width = 150
        self.Controls.Add(self.txtTextNoteCode)

        # Auto‑fill
        self.btnAutoFill = Button()
        self.btnAutoFill.Text = "Auto-fill Pipe Tag Codes"
        self.btnAutoFill.Location = Point(10, 410)
        self.btnAutoFill.Width = 200
        self.btnAutoFill.Click += self.autoFillPipeTagCodes
        self.Controls.Add(self.btnAutoFill)

        # OK / Cancel
        self.btnOK = Button()
        self.btnOK.Text = "OK"
        self.btnOK.Location = Point(600, 420)
        self.btnOK.DialogResult = DialogResult.OK
        self.btnOK.Click += self.okButton_Click
        self.Controls.Add(self.btnOK)

        self.btnCancel = Button()
        self.btnCancel.Text = "Cancel"
        self.btnCancel.Location = Point(720, 420)
        self.btnCancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(self.btnCancel)

        self.textNotePlaced = False
        self.Result = None

        # Add columns
        self.dataGrid.Columns.AddRange(
            Array[DataGridViewTextBoxColumn](
                [
                    self.colId,
                    self.colCategory,
                    self.colName,
                    self.colDefaultCode,
                    self.colNewCode,
                    self.colOD,
                    self.colLength,
                    self.colSize,
                ]
            )
        )
        self.dataGrid.Columns.Add(self.colArticle)
        self.dataGrid.Columns.Add(self.colTagStatus)

        # Populate rows
        for ed in elements_data:
            row_idx = self.dataGrid.Rows.Add()
            row = self.dataGrid.Rows[row_idx]
            row.Cells["Id"].Value = ed["Id"]
            row.Cells["Category"].Value = ed["Category"]
            row.Cells["Name"].Value = ed["Name"]
            row.Cells["GEB_Article_Number"].Value = ed.get("GEB_Article_Number", "")
            row.Cells["DefaultCode"].Value = ed["DefaultCode"]
            row.Cells["NewCode"].Value = ed["NewCode"]
            row.Cells["OutsideDiameter"].Value = ed["OutsideDiameter"]
            row.Cells["Length"].Value = ed["Length"]
            row.Cells["Size"].Value = ed.get("Size", "")

            # TagStatus logic
            cat = ed["Category"]
            status = ed["TagStatus"]
            if cat in ("Pipes", "Pipe Fittings"):
                # if there _is_ already a tag on this element, offer Remove
                if status == "Yes":
                    row.Cells["TagStatus"].Value = "Remove Tag"
                    row.Cells["TagStatus"].ReadOnly = False
                else:
                    row.Cells["TagStatus"].Value = "Add/Place Tag"
                    row.Cells["TagStatus"].ReadOnly = False
            elif cat == "Pipe Tags":
                # always allow removal of the tag itself
                row.Cells["TagStatus"].Value = "Remove Tag"
                row.Cells["TagStatus"].ReadOnly = False
            else:
                row.Cells["TagStatus"].Value = ""
            if ed["Category"] == "Pipes":
                row.DefaultCellStyle.BackColor = Color.LightBlue
            elif ed["Category"] == "Pipe Tags":
                row.DefaultCellStyle.BackColor = Color.LightGreen
            elif ed["Category"] == "Pipe Fittings":
                row.DefaultCellStyle.BackColor = Color.LightGoldenrodYellow
            elif ed["Category"] == "Text Notes":
                row.DefaultCellStyle.BackColor = Color.LightGray

    def _add_row(self, data):
        """Helper to append a new DataGridView row from a dict."""
        idx = self.dataGrid.Rows.Add()
        row = self.dataGrid.Rows[idx]
        for k, v in data.items():
            row.Cells[k].Value = v
        # keep the new tag’s button column read-only
        row.Cells["TagStatus"].Value = "Remove Tag"
        row.Cells["TagStatus"].ReadOnly = True

    def btnPlaceTextNote_Click(self, sender, event):
        text_note_code = self.txtTextNoteCode.Text.strip()
        if text_note_code == "":
            MessageBox.Show("Please enter a Text Note Code.", "Error")
            return
        if not self.regionElements or len(self.regionElements) == 0:
            MessageBox.Show(
                "Region elements not available to compute location.", "Error"
            )
            return
        (region_min, region_max) = get_region_bounding_box(self.regionElements)
        corner = region_min
        ttn = Transaction(doc, "Place Text Note")
        ttn.Start()
        note_type = FilteredElementCollector(doc).OfClass(TextNoteType).FirstElement()
        if note_type:
            opts = TextNoteOptions(note_type.Id)
            new_note = TextNote.Create(
                doc, doc.ActiveView.Id, corner, text_note_code, opts
            )
            if new_note:
                MessageBox.Show("Text Note created successfully.", "Success")
                self.textNotePlaced = True
        else:
            MessageBox.Show("No TextNoteType found.", "Error")
        ttn.Commit()

    def autoFillPipeTagCodes(self, sender, event):
        # 1) Parse base
        raw = self.txtTextNoteCode.Text.strip()
        m = re.search(r"([\d\.]+)", raw)
        if not m:
            MessageBox.Show("Could not parse base code from text note.", "Error")
            return
        base = m.group(1)  # e.g. "4.1.1"
        parts = base.split(".")
        if len(parts) >= 3:
            prefix = parts[0] + "." + parts[1]  # "4.1"
            try:
                base_n = int(parts[2])  # 1
            except:
                base_n = 0
        else:
            prefix, base_n = base, 0

        # 2) Collect indices
        fit_rows, pipe_rows, tag_rows = [], [], []
        for i in range(self.dataGrid.Rows.Count):
            cat = self.dataGrid.Rows[i].Cells["Category"].Value
            if cat == "Pipe Fittings":
                fit_rows.append(i)
            elif cat == "Pipes":
                pipe_rows.append(i)
            elif cat == "Pipe Tags":
                tag_rows.append(i)

        # 3) Compute centers for fittings & cluster nested ones
        fit_centers = []
        for idx in fit_rows:
            rid = int(str(self.dataGrid.Rows[idx].Cells["Id"].Value))
            elem = doc.GetElement(ElementId(rid))
            bbox = elem.get_BoundingBox(uidoc.ActiveView) or elem.get_BoundingBox(None)
            if bbox:
                ctr = XYZ(
                    (bbox.Min.X + bbox.Max.X) * 0.5,
                    (bbox.Min.Y + bbox.Max.Y) * 0.5,
                    (bbox.Min.Z + bbox.Max.Z) * 0.5,
                )
            else:
                ctr = XYZ(0, 0, 0)
            fit_centers.append((idx, ctr))

        clusters = []
        tol = 0.01  # about 3 mm in model units — instead of 1e‑6
        for idx, ctr in fit_centers:
            placed = False
            for cl in clusters:
                # compare against the cluster’s first center
                if abs(ctr.X - cl[0][1].X) < tol and abs(ctr.Y - cl[0][1].Y) < tol:
                    cl.append((idx, ctr))
                    placed = True
                    break
            if not placed:
                clusters.append([(idx, ctr)])
        # sort clusters left→right, down→up
        clusters.sort(key=lambda c: (c[0][1].X, c[0][1].Y))

        # 4) Number clusters starting at base_n + 0
        for i, cl in enumerate(clusters):
            code = prefix + "." + str(base_n + i)
            for idx, _ in cl:
                self.dataGrid.Rows[idx].Cells["NewCode"].Value = code

        # 5) Pipes sorted and numbered: full base + .1,.2...
        pipe_centers = []
        for idx in pipe_rows:
            rid = int(str(self.dataGrid.Rows[idx].Cells["Id"].Value))
            elem = doc.GetElement(ElementId(rid))
            bbox = elem.get_BoundingBox(uidoc.ActiveView)
            if bbox:
                ctr = XYZ(
                    (bbox.Min.X + bbox.Max.X) * 0.5,
                    (bbox.Min.Y + bbox.Max.Y) * 0.5,
                    (bbox.Min.Z + bbox.Max.Z) * 0.5,
                )
            else:
                ctr = XYZ(0, 0, 0)
            pipe_centers.append((idx, ctr))

        pipe_centers.sort(key=lambda x: (x[1].X, x[1].Y))
        for i, (idx, _) in enumerate(pipe_centers, 1):
            code = base + "." + str(i)
            self.dataGrid.Rows[idx].Cells["NewCode"].Value = code

        # 6) Mirror pipe numbering onto pipe‐tag rows (same count)
        for i, (idx, _) in enumerate(pipe_centers, 1):
            if i - 1 < len(tag_rows):
                trow = tag_rows[i - 1]
                code = base + "." + str(i)
                self.dataGrid.Rows[trow].Cells["NewCode"].Value = code

    def dataGrid_CellContentClick(self, sender, e):
        col = self.dataGrid.Columns[e.ColumnIndex].Name
        if col != "TagStatus":
            return

        row = self.dataGrid.Rows[e.RowIndex]
        cat = row.Cells["Category"].Value
        val = row.Cells["TagStatus"].Value

        # --- ADD/REMOVE ON PIPES & FITTINGS ---
        if cat in ("Pipes", "Pipe Fittings"):
            host_id = int(str(row.Cells["Id"].Value))
            host = doc.GetElement(ElementId(host_id))

            # --- ADD TAG ---
            if val == "Add/Place Tag":
                tr = Transaction(doc, "Add Tag")
                tr.Start()
                bb = host.get_BoundingBox(uidoc.ActiveView)
                if bb:
                    ctr = XYZ(
                        (bb.Min.X + bb.Max.X) / 2.0,
                        (bb.Min.Y + bb.Max.Y) / 2.0,
                        (bb.Min.Z + bb.Max.Z) / 2.0,
                    )
                    ref = Reference(host)
                    new_tag = IndependentTag.Create(
                        doc,
                        doc.ActiveView.Id,
                        ref,
                        True,
                        TagMode.TM_ADDBY_CATEGORY,
                        TagOrientation.Horizontal,
                        ctr,
                    )
                tr.Commit()

                # flip button
                row.Cells["TagStatus"].Value = "Remove Tag"

                # add the new‐tag row here
                te = doc.GetElement(new_tag.Id)
                if te:
                    data = {
                        "Id": str(te.Id),
                        "Category": "Pipe Tags",
                        "Name": te.Name or "",
                        "DefaultCode": host.LookupParameter("Comments").AsString()
                        or "",
                        "NewCode": row.Cells["NewCode"].Value,
                        "OutsideDiameter": row.Cells["OutsideDiameter"].Value,
                        "Length": row.Cells["Length"].Value,
                        "Size": "",
                        "GEB_Article_Number": "",
                        "TagStatus": "Yes",
                    }
                    self._add_row(data)
                return

            # --- REMOVE TAG ---
            if cat in ("Pipes", "Pipe Fittings") and val == "Remove Tag":
                host_id = int(str(row.Cells["Id"].Value))
                deleted_id = None

                # find the matching tag element
                for t in (
                    FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_PipeTags)
                    .WhereElementIsNotElementType()
                    .ToElements()
                ):
                    # unwrap LinkElementId if present
                    tagged = (
                        t.GetTaggedElementIds()
                        if hasattr(t, "GetTaggedElementIds")
                        else [t.TaggedElementId]
                    )
                    for rid in tagged:
                        # New: handle LinkElementId vs plain ElementId
                        eid = (
                            rid.HostElementId.IntegerValue
                            if hasattr(rid, "HostElementId")
                            else rid.IntegerValue
                        )
                        if eid == host_eid:
                            deleted_id = t.Id.IntegerValue
                            tag_elem_id = t.Id
                            break
                    if deleted_id:
                        break

                if deleted_id:
                    # unsubscribe row-selection highlight
                    self.dataGrid.SelectionChanged -= self.on_row_selected
                    # delete the tag in Revit
                    tr = Transaction(doc, "Remove Tag")
                    tr.Start()
                    doc.Delete(tag_elem_id)
                    tr.Commit()
                    # flip the host's button back to Add/Place
                    row.Cells["TagStatus"].Value = "Add/Place Tag"
                    row.Cells["TagStatus"].ReadOnly = False
                    # remove its Pipe-Tags row
                    for i in range(self.dataGrid.Rows.Count):
                        r2 = self.dataGrid.Rows[i]
                        if (
                            r2.Cells["Category"].Value == "Pipe Tags"
                            and int(str(r2.Cells["Id"].Value)) == deleted_id
                        ):
                            self.dataGrid.Rows.RemoveAt(i)
                            break
                    self.dataGrid.SelectionChanged += self.on_row_selected
                return

        # --- REMOVE AN ORPHAN PIPE-TAG ROW ---
        if cat == "Pipe Tags" and val == "Remove Tag":
            # Figure out who the host pipe was (while the tag still exists)
            tag_id = ElementId(int(str(row.Cells["Id"].Value)))
            self.dataGrid.SelectionChanged -= self.on_row_selected
            try:
                tag_elem = doc.GetElement(tag_id)
                host_id = None
                if tag_elem:
                    if hasattr(tag_elem, "GetTaggedElementIds"):
                        ids = tag_elem.GetTaggedElementIds()
                        if ids and ids.Count:
                            # if this is a LinkedElementId, use its HostElementId
                            rid = ids[0]
                            host_id = (
                                rid.HostElementId.IntegerValue
                                if hasattr(rid, "HostElementId")
                                else rid.IntegerValue
                            )
                    elif hasattr(tag_elem, "TaggedElementId"):
                        rid = tag_elem.TaggedElementId
                        host_id = (
                            rid.HostElementId.IntegerValue
                            if hasattr(rid, "HostElementId")
                            else rid.IntegerValue
                        )

                # Delete the Tag
                tr = Transaction(doc, "Remove Pipe-Tag")
                tr.Start()
                doc.Delete(tag_id)
                tr.Commit()

                # Remove the tags row
                self.dataGrid.Rows.RemoveAt(e.RowIndex)

                # Now find the pipe's row and flip it back to "Add/Place Tag"
                if host_id:
                    for i in range(self.dataGrid.Rows.Count):
                        pr = self.dataGrid.Rows[i]
                        if int(str(pr.Cells["Id"].Value)) == host_id and pr.Cells[
                            "Category"
                        ].Value in ("Pipes", "Pipe Fittings"):
                            pr.Cells["TagStatus"].Value = "Add/Place Tag"
                            pr.Cells["TagStatus"].ReadOnly = False
                            break
            finally:
                # rehook highlight
                self.dataGrid.SelectionChanged += self.on_row_selected
            return

        # --- PIPE FITTINGS ---
        elif cat == "Pipe Fittings" and val == "Add/Place Tag":
            # from Autodesk.Revit.DB import IndependentTag, Transaction, Reference
            elem_id = ElementId(int(str(row.Cells["Id"].Value)))
            fitting_elem = doc.GetElement(elem_id)

            t3 = Transaction(doc, "Add Fitting Tag")
            t3.Start()
            bbox = fitting_elem.get_BoundingBox(uidoc.ActiveView)
            if bbox:
                center = XYZ(
                    (bbox.Min.X + bbox.Max.X) / 2.0,
                    (bbox.Min.Y + bbox.Max.Y) / 2.0,
                    (bbox.Min.Z + bbox.Max.Z) / 2.0,
                )
                ref = Reference(fitting_elem)
                new_tag = IndependentTag.Create(
                    doc,
                    doc.ActiveView.Id,
                    ref,
                    True,
                    TagMode.TM_ADDBY_CATEGORY,
                    TagOrientation.Horizontal,
                    center,
                )
            t3.Commit()

            row.Cells["TagStatus"].Value = "Yes"

            new_tag_elem = doc.GetElement(new_tag.Id)
            if new_tag_elem:
                tag_dict = {
                    "Id": str(new_tag_elem.Id),
                    "Category": "Pipe Tags",
                    "Name": new_tag_elem.Name or "",
                    "DefaultCode": fitting_elem.LookupParameter("Comments").AsString()
                    or "",
                    "NewCode": row.Cells["NewCode"].Value,
                    "OutsideDiameter": row.Cells["OutsideDiameter"].Value,
                    "Length": row.Cells["Length"].Value,
                    "TagStatus": "Yes",
                }
                new_row_idx = self.dataGrid.Rows.Add()
                new_row = self.dataGrid.Rows[new_row_idx]
                for key, v in tag_dict.items():
                    new_row.Cells[key].Value = v
                new_row.Cells["TagStatus"].Value = "Remove Tag"
                new_row.Cells["TagStatus"].ReadOnly = True

    def okButton_Click(self, sender, event):
        updated_data = []
        for row in self.dataGrid.Rows:
            entry = {
                "Id": row.Cells["Id"].Value,
                "Category": row.Cells["Category"].Value,
                "Name": row.Cells["Name"].Value,
                "DefaultCode": row.Cells["DefaultCode"].Value,
                "NewCode": row.Cells["NewCode"].Value,
                "OutsideDiameter": row.Cells["OutsideDiameter"].Value,
                "Length": row.Cells["Length"].Value,
                "TagStatus": row.Cells["TagStatus"].Value,
            }
            updated_data.append(entry)

        self.Result = {
            "Elements": updated_data,
            "TextNotePlaced": self.textNotePlaced,
            "TextNote": self.txtTextNoteCode.Text.strip(),
        }
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_row_selected(self, sender, event):
        """When the user clicks or arrows to a row, select that element in Revit."""
        row = self.dataGrid.CurrentRow
        if not row:
            return
        id_val = row.Cells["Id"].Value
        if not id_val:
            return

        # try to parse and highlight, but swallow any invalid-object errors
        try:
            eid = ElementId(int(str(id_val)))
            elem = doc.GetElement(eid)
            # guard against deleted/invalid elements
            if elem and elem.IsValidObject:
                uidoc.Selection.SetElementIds(List[ElementId]([eid]))
        except:
            return


def show_element_editor(elements_data, region_elements=None):
    form = ElementEditorForm(elements_data, region_elements)
    if form.ShowDialog() == DialogResult.OK:
        return form.Result
    return None


# ==================================================
# Filter Gathered Elements to Relevant Categories
# ==================================================
def filter_relevant_elements(gathered_elements):
    """
    Build a list of dicts with keys:
     "Id","Category","Name","DefaultCode","NewCode",
     "OutsideDiameter","Length","GEB_Article_Number","TagStatus"
    """
    relevant = []

    pipe_ids = {
        e.Id.IntegerValue
        for e in gathered_elements
        if e.Category and e.Category.Name == "Pipes"
    }
    # grab all tags in the view
    all_pipe_tags = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_PipeTags)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    # pull in any tags whose host pipe was in our region
    for tag in all_pipe_tags:
        try:
            host = None
            if hasattr(tag, "GetTaggedElementIds"):
                ids = tag.GetTaggedElementIds()
                if ids and ids.Count > 0:
                    host = doc.GetElement(ids[0])
            elif hasattr(tag, "TaggedElementId"):
                host = doc.GetElement(tag.TaggedElementId)
            if host and host.Id.IntegerValue in pipe_ids:
                # and only if we haven't already added it in gathered_elements
                if not any(str(tag.Id) == d["Id"] for d in relevant):
                    # build your dict exactly like you do for pipe‑tags below
                    relevant.append(
                        {
                            "Id": str(tag.Id),
                            "Category": "Pipe Tags",
                            "Name": tag.Name or "",
                            "DefaultCode": host.LookupParameter("Comments").AsString()
                            or "",
                            "NewCode": host.LookupParameter("Comments").AsString()
                            or "",
                            "OutsideDiameter": convert_param_to_string(
                                host.LookupParameter("Outside Diameter")
                            ),
                            "Length": convert_param_to_string(
                                host.LookupParameter("Length")
                            ),
                            "Size": "",  # if you want
                            "GEB_Article_Number": "",
                            "TagStatus": "Yes",
                        }
                    )
        except:
            pass

    for e in gathered_elements:
        if not e.Category:
            continue
        cat = e.Category.Name
        if cat not in ("Pipes", "Pipe Fittings", "Pipe Tags", "Text Notes"):
            continue

        com = e.LookupParameter("Comments")
        default_code = com.AsString() if com and com.AsString() else ""

        # initialize
        outside_diam = ""
        length_val = ""
        art_num = ""
        tag_status = ""

        # --- Pipes ---
        if cat == "Pipes":
            odp = e.LookupParameter("Outside Diameter")
            lp = e.LookupParameter("Length")
            outside_diam = convert_param_to_string(odp)
            length_val = convert_param_to_string(lp)

            # detect existing tags
            tag_status = "No"
            for tag in all_pipe_tags:
                try:
                    tagged_ids = []
                    if hasattr(tag, "GetTaggedElementIds"):
                        tagged_ids = tag.GetTaggedElementIds()
                    elif hasattr(tag, "TaggedElementId"):
                        tagged_ids = [tag.TaggedElementId]
                    for tid in tagged_ids:
                        if tid and tid.IntegerValue == e.Id.IntegerValue:
                            tag_status = "Yes"
                            break
                    if tag_status == "Yes":
                        break
                except:
                    pass

        # --- Pipe Fittings ---
        elif cat == "Pipe Fittings":
            # diameter (try several names)
            for pname in ("Outside Diameter", "Diameter", "Nominal Diameter"):
                p = e.LookupParameter(pname)
                if p:
                    outside_diam = convert_param_to_string(p)
                    break
            # length
            lp = e.LookupParameter("Length")
            length_val = convert_param_to_string(lp)
            # GEB article
            ap = e.LookupParameter("GEB_Article_Number")
            art_num = ap.AsString() if ap and ap.AsString() else ""

            # only the specific fitting gets Add/Place Tag
            if e.Name and e.Name.find("DN") >= 0:
                tag_status = "No"
            else:
                tag_status = ""

        # --- Pipe Tags ---
        elif cat == "Pipe Tags":
            tag_status = "Yes"
            host = None
            try:
                if hasattr(e, "GetTaggedElementIds"):
                    ids = e.GetTaggedElementIds()
                    if ids and ids.Count > 0:
                        host = doc.GetElement(ids[0])
                if not host and hasattr(e, "TaggedElementId"):
                    host = doc.GetElement(e.TaggedElementId)
            except:
                host = None

            if host:
                odp = host.LookupParameter("Outside Diameter")
                lp = host.LookupParameter("Length")
                outside_diam = convert_param_to_string(odp)
                length_val = convert_param_to_string(lp)

        # --- Text Notes & others ---
        else:
            tag_status = ""

        size_val = ""
        if cat == "Pipe Fittings":
            param_size = e.LookupParameter("Size")
            size_val = convert_param_to_string(param_size)

        relevant.append(
            {
                "Id": str(e.Id),
                "Category": cat,
                "Name": e.Name if hasattr(e, "Name") else "",
                "DefaultCode": default_code,
                "NewCode": default_code,
                "OutsideDiameter": outside_diam,
                "Length": length_val,
                "Size": size_val,
                "GEB_Article_Number": art_num,
                "TagStatus": tag_status,
            }
        )

    return relevant


# ==================================================
# MAIN WORKFLOW
# ==================================================
gathered_elements = select_boundary_and_gather()
if gathered_elements is None or len(gathered_elements) == 0:
    MessageBox.Show("No elements were gathered. Operation cancelled.", "Error")
    sys.exit("Operation cancelled by the user.")

filtered_elements = filter_relevant_elements(gathered_elements)
if len(filtered_elements) == 0:
    MessageBox.Show("No relevant elements found in the selected region.", "Error")
    sys.exit("Operation cancelled by the user.")

result = show_element_editor(filtered_elements, region_elements=gathered_elements)
if result is None:
    sys.exit("Operation cancelled by the user.")

# --- Renumber Pipes based on region order (sorted left-to-right, bottom-to-up) ---
if not result.get("TextNotePlaced", False):
    base_raw = result.get("TextNote", "").strip()
    m = re.search(r"([\d\.]+)", base_raw)
    base = m.group(1) if m else "0"

    pipe_entries = []
    for idx, eData in enumerate(result["Elements"]):
        if eData["Category"] == "Pipes":
            elem = doc.GetElement(ElementId(int(str(eData["Id"]))))
            if elem:
                bbox = elem.get_BoundingBox(uidoc.ActiveView)
                if bbox:
                    center = XYZ(
                        (bbox.Min.X + bbox.Max.X) / 2.0,
                        (bbox.Min.Y + bbox.Max.Y) / 2.0,
                        (bbox.Min.Z + bbox.Max.Z) / 2.0,
                    )
                    pipe_entries.append((idx, center))
    pipe_entries.sort(key=lambda x: (x[1].X, x[1].Y))

    ctr = 1
    for i, _ in pipe_entries:
        result["Elements"][i]["NewCode"] = base + "." + str(ctr)
        ctr += 1

    for eData in result["Elements"]:
        if eData["Category"] == "Pipe Fittings":
            eData["NewCode"] = base

# --- Update the elements' "Comments" from the DataGridView ---
t = Transaction(doc, "Update Comments")
t.Start()
for eData in result["Elements"]:
    if not eData["Id"] or str(eData["Id"]).lower() == "none":
        continue
    elem = doc.GetElement(ElementId(int(str(eData["Id"]))))
    if elem:
        p = elem.LookupParameter("Comments")
        if p and not p.IsReadOnly:
            p.Set(str(eData["NewCode"]))
t.Commit()

# --- Place the text note if not already placed ---
if not result.get("TextNotePlaced", False):
    (region_min, region_max) = get_region_bounding_box(gathered_elements)
    view = doc.ActiveView
    corner = region_min
    ttn = Transaction(doc, "Place Text Note at Region Corner")
    ttn.Start()
    nt = FilteredElementCollector(doc).OfClass(TextNoteType).FirstElement()
    if nt:
        opts = TextNoteOptions(nt.Id)
        TextNote.Create(
            doc, doc.ActiveView.Id, corner, result.get("TextNote", base), opts
        )
    ttn.Commit()

region_min, region_max = get_region_bounding_box(gathered_elements)

orig = uidoc.ActiveView
if orig.ViewType != ViewType.FloorPlan:
    MessageBox.Show("Active view is not a Floor Plan!", "Error")
    sys.exit()

tx = Transaction(doc, "Create Cropped Plan View")
tx.Start()
new_id = orig.Duplicate(ViewDuplicateOption.Duplicate)
new_view = doc.GetElement(new_id)

# remove any view template and set scale to 1:25
new_view.ViewTemplateId = ElementId.InvalidElementId

# force it to Coordination
new_view.Discipline = ViewDiscipline.Coordination

new_view.Scale = 25

# naming, cropping, discipline etc...
m = re.search(r"([\d\.]+)", result["TextNote"])
base = m.group(1) if m else result["TextNote"].strip()  # "5.1.1"
try:
    new_view.Name = base
except ArgumentException:
    MessageBox.Show(
        "A view named '{0}' already exists!\n\n"
        "Please pick a different code in the text-node editor.".format(base),
        "Duplicate View Name",
    )
    tx.RollBack()
    sys.exit("Duplicate View Name")

new_view.CropBoxActive = True
new_view.CropBoxVisible = True
bb = BoundingBoxXYZ()
bb.Min = region_min
bb.Max = region_max
new_view.CropBox = bb

tx.Commit()

# -----------------------------------------
# 2) SHOW TITLE-BLOCK PICKER, THEN CREATE SHEET
# -----------------------------------------

# collect title‑blocks
all_tbs = list(
    FilteredElementCollector(doc)
    .OfCategory(BuiltInCategory.OST_TitleBlocks)
    .OfClass(FamilySymbol)
    .ToElements()
)

if not all_tbs:
    MessageBox.Show("No title‑block types found.", "Error")
    sys.exit()


class TBPicker(Form):
    def __init__(self, tbs):
        self.tbs = tbs
        self.Text = "Choose a Title‑Block"
        self.ClientSize = Size(300, 350)

        # ListBox
        self.lb = ListBox()
        self.lb.Bounds = Rectangle(10, 10, 280, 280)
        for sym in tbs:
            fam = sym.FamilyName
            type_name = (
                sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() or ""
            )
            self.lb.Items.Add(fam + " - " + type_name)
        self.Controls.Add(self.lb)

        # OK / Cancel
        ok = Button(Text="OK", DialogResult=DialogResult.OK, Location=Point(10, 300))
        ca = Button(
            Text="Cancel", DialogResult=DialogResult.Cancel, Location=Point(100, 300)
        )
        self.Controls.Add(ok)
        self.Controls.Add(ca)
        self.AcceptButton = ok
        self.CancelButton = ca


existing_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
existing_numbers = {s.SheetNumber for s in existing_sheets}

# show the picker
picker = TBPicker(all_tbs)
if picker.ShowDialog() != DialogResult.OK or picker.lb.SelectedIndex < 0:
    MessageBox.Show("Sheet creation cancelled.", "Info")
    sys.exit()

title_block = all_tbs[picker.lb.SelectedIndex]

# ————————————————————————————————
# 3) Create A3 sheets, skipping duplicates
# ————————————————————————————————

# b) for each base, only create if it’s not already on a sheet
for base in {base}:  # e.g. 5.1.1
    if base in existing_numbers:
        MessageBox.Show(
            "Sheet 'prefab {0}' already exists!\n\n"
            "Please pick a different code in the text-note editor.".format(base),
            "Duplicate Sheet",
        )
        sys.exit("Duplicate sheet number")
    t3 = Transaction(doc, "Create 3D callout")
    t3.Start()
    sheet = ViewSheet.Create(doc, title_block.Id)
    sheet.SheetNumber = base
    sheet.Name = "Prefab " + base

    o = sheet.Outline
    center = XYZ((o.Min.U + o.Max.U) / 2, (o.Min.V + o.Max.V) / 2, 0)
    Viewport.Create(doc, sheet.Id, new_view.Id, center)

    # ------------------------------------------
    # 4) Create & place 3D callout
    # ------------------------------------------
    all3ds = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
    all3d_views = FilteredElementCollector(doc).OfClass(View3D).ToElements()
    # 1. pick a 3D ViewFamilyType
    prefix = "{} - Sheet".format(base)
    existing_count = sum(1 for v in all3d_views if v.Name.startswith(prefix))

    v3d_type = next(v for v in all3ds if v.ViewFamily == ViewFamily.ThreeDimensional)

    # split off the last number of the base code
    parts = base.split(".")
    major = ".".join(parts[:-1])
    last = int(parts[-1])
    new_last = last + existing_count
    sheet_suffix = "{}.{}".format(major, new_last)

    # 2. create an isometric 3D view
    view3d = View3D.CreateIsometric(doc, v3d_type.Id)
    view3d.Name = "{} - Sheet {}".format(base, sheet_suffix)
    # force it into the Architectural branch of the browser
    view3d.Discipline = ViewDiscipline.Architectural
    view3d.Scale = 25

    # apply your A00_Algemeen 3D View Template
    tmpl = next(
        (v for v in all3d_views if v.IsTemplate and v.Name == "S4R_A00_Algemeen_3D"),
        None,
    )
    if tmpl:
        view3d.ViewTemplateId = tmpl.Id

    param = view3d.get_Parameter(BuiltInParameter.VIEW_DISCIPLINE)
    if not param.IsReadOnly:
        param.Set(int(ViewDiscipline.Architectural))

    # 3. use the same region bounding box you computed earlier
    section_bb = BoundingBoxXYZ()
    section_bb.Min = region_min
    section_bb.Max = region_max
    view3d.SetSectionBox(section_bb)

    # 4. position the 3D viewport on the sheet (to the right of the floor plan)
    o = sheet.Outline
    # push it over 1/3 of the sheet width, and up a bit
    u = o.Min.U + (o.Max.U - o.Min.U) * 0.65
    v = o.Min.V + (o.Max.V - o.Min.V) * 0.35

    Viewport.Create(doc, sheet.Id, view3d.Id, XYZ(u, v, 0))

    t3.Commit()
