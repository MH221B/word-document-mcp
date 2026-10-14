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
    anchor = session.guard(lambda: doc.Range(table.Range.End, table.Range.End))
    session.guard(lambda: anchor.InsertAfter("\r"))
    para = session.guard(lambda: doc.Range(anchor.Start, anchor.Start + 1).Paragraphs(1).Range)
    session.guard(
        lambda: para.InsertCaption(
            Label=label,
            Title=title,
            Position=WD_CAPTION_BELOW if position == "below" else WD_CAPTION_ABOVE,
        )
    )
    return {
        "caption_added": True,
        "label": label,
        "title": title,
        "position": position,
        "text": str(session.guard(lambda: para.Text)).strip(),
    }
