"""edit_table, edit_page_layout, edit_fields."""

from __future__ import annotations

from word_mcp.errors import InvalidArgument, TargetNotFound
from word_mcp.history import ChangeStack
from word_mcp.operations import gate_delete, gate_range
from word_mcp.session import require_writable

WD_CAPTION_ABOVE = -1
WD_CAPTION_BELOW = 1


def find_table(session, table_index: int, doc):
    tables = session.guard(lambda: list(doc.Tables))
    if table_index < 1 or table_index > len(tables):
        raise TargetNotFound(
            f"There is no table {table_index}.", tables=len(tables)
        )
    return tables[table_index - 1]


def edit_table(
    session,
    stack: ChangeStack,
    resolve,
    action: str,
    spec=None,
    table_index: int = 1,
    row: int = 0,
    column: int = 0,
    rows: int = 1,
    columns: int = 1,
    values: list | None = None,
    formatting: dict | None = None,
    label: str = "Table",
    title: str = "",
    position: str = "below",
    confirm: bool = False,
) -> dict:
    doc = session.doc()
    require_writable(doc)

    if action == "insert":
        return _insert_table(session, resolve, spec, rows, columns, values)

    table = find_table(session, table_index, doc)

    if action == "add_rows":
        session.guard(lambda: table.Rows.Add())
        return {"table": table_index, "rows": int(session.guard(lambda: table.Rows.Count))}

    if action == "add_columns":
        session.guard(lambda: table.Columns.Add())
        return {"table": table_index, "columns": int(session.guard(lambda: table.Columns.Count))}

    if action == "delete_row":
        _check_row(session, table, row)
        session.guard(lambda: table.Rows(row).Delete())
        return {"table": table_index, "deleted_row": row}

    if action == "delete_column":
        _check_column(session, table, column)
        session.guard(lambda: table.Columns(column).Delete())
        return {"table": table_index, "deleted_column": column}

    if action == "merge_cells":
        _check_row(session, table, row)
        _check_column(session, table, column)
        session.guard(lambda: table.Cell(row, column).Merge(table.Cell(row, column + 1)))
        return {"table": table_index, "merged": [row, column]}

    if action == "split_cell":
        _check_row(session, table, row)
        _check_column(session, table, column)
        session.guard(lambda: table.Cell(row, column).Split(rows, columns))
        return {"table": table_index, "split": [row, column]}

    if action == "format_cell":
        _check_row(session, table, row)
        _check_column(session, table, column)
        cell = session.guard(lambda: table.Cell(row, column))
        _apply_cell_formatting(session, cell, formatting or {})
        return {"table": table_index, "cell": [row, column], "applied": formatting or {}}

    if action == "style":
        name = (formatting or {}).get("style")
        if not name:
            raise InvalidArgument("action='style' needs formatting={'style': '<name>'}.")
        session.guard(lambda: setattr(table, "Style", name))
        return {"table": table_index, "style": name}

    if action == "header_row":
        session.guard(lambda: setattr(table.Rows(1), "HeadingFormat", True))
        return {"table": table_index, "header_row_repeat": True}

    if action == "caption":
        return _caption_table(session, table, label, title, position)

    raise InvalidArgument(
        f"Unknown action {action!r}.",
        allowed=[
            "insert", "add_rows", "add_columns", "delete_row", "delete_column",
            "merge_cells", "split_cell", "format_cell", "style", "header_row", "caption",
        ],
    )


def _insert_table(session, resolve, spec, rows, columns, values) -> dict:
    doc = session.doc()
    match = resolve(session, spec)
    rng = match.handle
    table = session.guard(lambda: doc.Tables.Add(rng, rows, columns))
    if values:
        for r, row_values in enumerate(values, start=1):
            for c, value in enumerate(row_values, start=1):
                if r > rows or c > columns:
                    break
                session.guard(lambda t=table, rr=r, cc=c, v=value: setattr(t.Cell(rr, cc).Range, "Text", str(v)))
    return {
        "inserted_table": True,
        "rows": rows,
        "columns": columns,
        "at": match.describe(),
        "tables_now": len(session.guard(lambda: list(doc.Tables))),
    }


def _check_row(session, table, row: int) -> None:
    if row < 1 or row > int(session.guard(lambda: table.Rows.Count)):
        raise InvalidArgument(f"Row {row} is out of range.")


def _check_column(session, table, column: int) -> None:
    if column < 1 or column > int(session.guard(lambda: table.Columns.Count)):
        raise InvalidArgument(f"Column {column} is out of range.")


def _apply_cell_formatting(session, cell, formatting: dict) -> None:
    for prop, value in formatting.items():
        if prop == "background":
            session.guard(lambda v=value: setattr(cell.Shading, "BackgroundPatternColor", v))
        elif prop == "width":
            session.guard(lambda v=value: setattr(cell, "Width", v))
        elif prop == "font":
            for f_prop, f_value in dict(value).items():
                session.guard(lambda p=f_prop, v=f_value: setattr(cell.Range.Font, p, v))
        else:
            raise InvalidArgument(f"Unknown cell property {prop!r}.")


def _caption_table(session, table, label: str, title: str, position: str) -> dict:
    """Insert a caption paragraph after or before the table.

    Word's captions are SEQ fields, so they renumber on their own and a table of
    figures can collect them.
    """
    doc = session.doc()
    # Caption the table's own range; InsertCaption adds the caption as its own
    # paragraph, which is then located so it can be styled and its text reported.
    end = int(session.guard(lambda: table.Range.End))
    session.guard(
        lambda: table.Range.InsertCaption(
            Label=label,
            Title=title,
            Position=WD_CAPTION_BELOW if position == "below" else WD_CAPTION_ABOVE,
        )
    )
    caption_para = None
    for candidate in session.guard(lambda: list(doc.Paragraphs)):
        if int(session.guard(lambda c=candidate: c.Range.Start)) == end:
            caption_para = candidate.Range
            break
    if caption_para is not None:
        session.guard(lambda: setattr(caption_para, "Style", "Caption"))
    text = (
        str(session.guard(lambda: caption_para.Text)).strip()
        if caption_para is not None
        else ""
    )
    return {
        "caption_added": True,
        "label": label,
        "title": title,
        "position": position,
        "text": text,
    }


WD_ORIENT_PORTRAIT = 0
WD_ORIENT_LANDSCAPE = 1
WD_SECTION_BREAK_NEXT_PAGE = 2
WD_SECTION_BREAK_CONTINUOUS = 3
WD_HEADER_FOOTER_PRIMARY = 1
WD_HEADER_FOOTER_FIRST_PAGE = 2
WD_HEADER_FOOTER_EVEN = 3

LAYOUT_POINT_PROPERTIES = {
    "top_margin": "TopMargin",
    "bottom_margin": "BottomMargin",
    "left_margin": "LeftMargin",
    "right_margin": "RightMargin",
    "page_width": "PageWidth",
    "page_height": "PageHeight",
    "header_distance": "HeaderDistance",
    "footer_distance": "FooterDistance",
    "gutter": "Gutter",
}


def edit_page_layout(
    session,
    resolve,
    action: str,
    spec=None,
    section: int = 1,
    orientation: str = "",
    points: dict | None = None,
    columns: int = 1,
    side: str = "header",
    which: str = "primary",
    text: str = "",
    different_first_page: bool | None = None,
    odd_and_even: bool | None = None,
) -> dict:
    doc = session.doc()
    require_writable(doc)
    sections = session.guard(lambda: list(doc.Sections))
    if section < 1 or section > len(sections):
        raise InvalidArgument(f"Section {section} is out of range.", sections=len(sections))
    target_section = sections[section - 1]
    setup = session.guard(lambda: target_section.PageSetup)

    if action == "page_setup":
        if orientation:
            value = WD_ORIENT_LANDSCAPE if orientation == "landscape" else WD_ORIENT_PORTRAIT
            session.guard(lambda: setattr(setup, "Orientation", value))
        for key, value in (points or {}).items():
            if key not in LAYOUT_POINT_PROPERTIES:
                raise InvalidArgument(f"Unknown page setup property {key!r}.", allowed=sorted(LAYOUT_POINT_PROPERTIES))
            session.guard(lambda k=key, v=value: setattr(setup, LAYOUT_POINT_PROPERTIES[k], v))
        if columns > 1:
            session.guard(lambda: setup.TextColumns.SetCount(columns))
        return {
            "section": section,
            "orientation": "landscape" if session.guard(lambda: setup.Orientation) == WD_ORIENT_LANDSCAPE else "portrait",
            "columns": columns,
            "applied": points or {},
        }

    if action == "section_break":
        match = resolve(session, spec)
        kind = WD_SECTION_BREAK_CONTINUOUS if columns == 0 else WD_SECTION_BREAK_NEXT_PAGE
        session.guard(lambda: match.handle.InsertBreak(kind))
        return {"section_break": "continuous" if columns == 0 else "next_page", "at": match.describe()}

    if action == "header_footer":
        index = {
            "primary": WD_HEADER_FOOTER_PRIMARY,
            "first_page": WD_HEADER_FOOTER_FIRST_PAGE,
            "even": WD_HEADER_FOOTER_EVEN,
        }.get(which)
        if index is None:
            raise InvalidArgument(f"Unknown header or footer {which!r}.")
        collection = (
            session.guard(lambda: target_section.Headers(index))
            if side == "header"
            else session.guard(lambda: target_section.Footers(index))
        )
        session.guard(lambda: setattr(collection.Range, "Text", text))
        return {"side": side, "which": which, "text": text}

    if action == "flags":
        if different_first_page is not None:
            session.guard(lambda: setattr(setup, "DifferentFirstPageHeaderFooter", different_first_page))
        if odd_and_even is not None:
            session.guard(lambda: setattr(setup, "OddAndEvenPagesHeaderFooter", odd_and_even))
        return {"different_first_page": different_first_page, "odd_and_even": odd_and_even}

    raise InvalidArgument(
        f"Unknown action {action!r}.", allowed=["page_setup", "section_break", "header_footer", "flags"]
    )


WD_FIELD_PAGE = 33
WD_FIELD_DATE = 31
WD_FIELD_SEQUENCE = 12
WD_FIELD_TOC = 37


def edit_fields(
    session,
    resolve,
    action: str,
    spec=None,
    field: str = "page",
    toc_options: dict | None = None,
) -> dict:
    doc = session.doc()
    require_writable(doc)

    if action == "insert":
        match = resolve(session, spec)
        kinds = {"page": WD_FIELD_PAGE, "date": WD_FIELD_DATE, "sequence": WD_FIELD_SEQUENCE}
        if field not in kinds:
            raise InvalidArgument(f"Unknown field {field!r}.", allowed=sorted(kinds))
        inserted = session.guard(lambda: match.handle.Fields.Add(match.handle, kinds[field]))
        return {"inserted": field, "at": match.describe(), "code": str(session.guard(lambda: inserted.Code.Text))}

    if action == "update_all":
        session.guard(lambda: doc.Fields.Update())
        return {"updated": "all fields"}

    if action == "update":
        match = resolve(session, spec)
        fields = session.guard(lambda: list(match.handle.Fields))
        for item in fields:
            session.guard(lambda f=item: f.Update())
        return {"updated": len(fields), "at": match.describe()}

    if action == "toc":
        options = toc_options or {}
        match = resolve(session, spec)
        toc = session.guard(
            lambda: doc.TablesOfContents.Add(
                Range=match.handle,
                UseHeadingStyles=True,
                UpperHeadingLevel=options.get("upper_heading_level", 1),
                LowerHeadingLevel=options.get("lower_heading_level", 3),
                RightAlignPageNumbers=True,
                IncludePageNumbers=True,
                UseHyperlinks=True,
            )
        )
        session.guard(lambda: toc.Update())
        return {"toc_inserted": True, "at": match.describe()}

    if action == "table_of_figures":
        options = toc_options or {}
        match = resolve(session, spec)
        label = options.get("label", "Figure")
        tof = session.guard(
            lambda: doc.TablesOfFigures.Add(
                Range=match.handle,
                Caption=label,
                IncludeLabel=True,
                RightAlignPageNumbers=True,
                IncludePageNumbers=True,
                UseHyperlinks=True,
            )
        )
        session.guard(lambda: tof.Update())
        return {"table_of_figures_inserted": True, "label": label, "at": match.describe()}

    raise InvalidArgument(
        f"Unknown action {action!r}.",
        allowed=["insert", "update", "update_all", "toc", "table_of_figures"],
    )
