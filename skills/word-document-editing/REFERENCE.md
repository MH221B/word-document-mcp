# word-mcp reference

Every tool returns JSON. Success is the operation's own payload. Failure is
`{"error": {"code": ..., "message": ..., ...}}`, where the extra keys carry the
useful detail, such as `candidates` on `ambiguous_target`.

## Targeting

A `target` argument is an object with any of these keys:

| Key | Meaning |
|---|---|
| `token` | An anchor from an earlier read. Most precise, and immune to text shifting. |
| `text` | A literal or wildcard pattern. |
| `style` | Restrict matches to a paragraph style. |
| `heading_path` | Restrict matches to one heading, using the `"Overview > Risks"` form. |
| `container` | `body` (default), `header`, or `footer`. Section 1 only. Offsets inside a header or footer belong to that story, so targets and edits there stay in it. |
| `wildcards` | Treat `text` as a Word wildcard pattern. |
| `match_case` | Case-sensitive matching. |
| `occurrence` | Pick the nth match instead of failing on ambiguity. |

`heading_path` lists **ancestors only**, nearest last: headings with a smaller
outline level. A heading does not name itself, and a previous sibling at the
same level is not an ancestor. So a match under a level-2 heading reports
`"Top > Sub"`, and a level-1 heading reports `""`.

`find_text` collapses whitespace and folds curly quotes to straight when it
decides whether a candidate `matched`; the Word search itself uses the text as
given. Searches stop at 50 candidates. `footnote`, `endnote`, and `comment`
containers are accepted by the spec but not implemented in search, which reports
`invalid_argument`.

There is no paragraph-index addressing. Use the tokens from `document_outline`
or `find_text`.

## Connection

| Tool | Parameters | Notes |
|---|---|---|
| `list_open_documents` | none | Name, path, read-only, protected. Every Word instance on the machine is searched, because a machine runs several `WINWORD` processes and the document you see is not always in the one `GetActiveObject` returns. |
| `connect_document` | `path_or_name` | Pins one document. Ambiguous names return `document_ambiguous`. |
| `disconnect_document` | none | Forgets the pin. |

## Inspection

| Tool | Parameters | Returns |
|---|---|---|
| `document_outline` | `depth=9`, `refresh=false` | Headings with tokens, tables with size, sections with page size, images. Structure is cached, so calling it again on an unchanged document is cheap. `refresh=true` rebuilds after a structural change made by hand in Word. |
| `read_range` | `target`, `include_formatting=False` | Text, and optionally the formatting block. |
| `get_formatting` | `target` | Paragraph and character style, font and paragraph snapshots, and `deviations_from_style`. |
| `find_text` | `text`, `container`, `heading_path`, `style`, `wildcards`, `match_case`, `limit=20` | Candidates with context. Never fails on ambiguity. |
| `inspect_document` | none | Template path, protection, track-changes state, styles, comments, revisions, bookmarks, content controls, properties. |
| `get_ooxml` | `target` | Flat WordprocessingML for the range. |

`deviations_from_style` lists only the properties where direct formatting differs
from the style definition, so an empty dict means the paragraph is
template-conformant.

## Editing

`edit_text` (`target`, `action`, `text`, `position`, `format_mode`, `formatting`, `confirm`)
Actions: `insert`, `replace`, `delete`.

`format_text` (`target`, `action`, `formatting`, `source`, `confirm`)
Actions: `character`, `paragraph`, `clear`, `copy_from` where `source` is another target.

`manage_style` (`action`, `target`, `name`, `based_on`, `formatting`, `confirm`)
Actions: `list`, `create`, `update`, `delete`, `apply`.

`manage_list` (`target`, `action`, `list_format`, `level`, `start_at`)
Actions: `apply`, `continue`, `restart`, `set_level`.

`edit_table` (`action`, `target`, `table_index`, `row`, `column`, `rows`, `columns`, `values`, `formatting`, `label`, `title`, `position`)
Actions: `insert`, `add_rows`, `add_columns`, `delete_row`, `delete_column`, `merge_cells`, `split_cell`, `format_cell`, `style`, `header_row`, `caption`.
Cell formatting keys are `background`, `width`, and `font`.

`edit_page_layout` (`action`, `target`, `section`, `orientation`, `points`, `columns`, `side`, `which`, `text`, `different_first_page`, `odd_and_even`)
Actions: `page_setup`, `section_break`, `header_footer`, `flags`.
`points` keys: `top_margin`, `bottom_margin`, `left_margin`, `right_margin`, `page_width`, `page_height`, `header_distance`, `footer_distance`, `gutter`.

`edit_fields` (`action`, `target`, `field`, `toc_options`)
Actions: `insert` (`field` is `page`, `date`, or `sequence`), `update`, `update_all`, `toc`, `table_of_figures`.

`edit_media` (`action`, `target`, `image_index`, `path`, `width`, `height`, `wrap`, `label`, `title`, `position`)
Actions: `insert`, `size`, `wrap`, `caption`. `wrap` is `inline`, `square`, `tight`, `through`, or `top_bottom`.

`edit_links` (`action`, `target`, `address`, `sub_address`, `screen_tip`, `display_text`, `name`, `text`)
Actions: `hyperlink`, `bookmark_add`, `bookmark_find`, `footnote`, `endnote`.

`edit_content_controls` (`action`, `target`, `control_index`, `control_type`, `tag`, `title`, `value`, `items`)
Actions: `insert`, `set_value`, `list`. `control_type` is `rich_text`, `text`, `dropdown`, `date`, or `checkbox`.

`edit_annotations` (`action`, `target`, `text`, `comment_index`, `done`, `mode`, `confirm`)
Actions: `add_comment`, `reply_comment`, `resolve_comment`, `delete_comment`, `track_changes` (`mode` is `on`, `off`, or `mine`), `accept_all`, `reject_all`.

`set_document_properties` (`properties`)
Built-in names: `Title`, `Author`, `Subject`, `Category`, `Comments`, `Keywords`, `Company`. Anything else becomes a custom property.

`save_document` (none)
Saves the pinned document. Nothing calls it implicitly, and AutoSave documents save themselves.

`revert_change` (`operation_id`)
Reverts the newest change only. Restores text and formatting, not deleted images, fields, or comments.

`insert_ooxml` (`target`, `ooxml`, `confirm`)
Inserts a WordprocessingML fragment. Flat `w:` markup, not the `pkg:package`
wrapper the Office.js API uses. Prefer the typed tools, and use this for
constructs with no object model.

## Errors

| Code | What to do |
|---|---|
| `word_not_running` | Ask the user to start Word. |
| `document_not_found` | The pin was lost, or the document was renamed or moved. Reconnect. |
| `document_ambiguous` | Two open documents share the name. Pass a path. |
| `ambiguous_target` | Narrow the target using `candidates`. |
| `target_not_found` | Re-read the document (the text may have changed), or use a valid table, image, comment, or control index. |
| `stale_token` | The anchored range was deleted. Read again for a fresh token. |
| `word_busy` | A dialog or an active edit is blocking Word. Wait for the user. |
| `word_disconnected` | Word closed. Reconnect after it reopens. |
| `protected_document` | Editing restrictions are on. Report them. |
| `read_only_document` | The document opened read-only. Report it. |
| `confirmation_required` | Re-issue with `confirm=true` once the user agrees. |
| `invalid_argument` | The message names the problem and usually the allowed values. |
| `word_error` | A Word COM failure the server could not classify, or the busy timeout. The message carries Word's error. |

## Recipes

**Watermark.** There is no watermark object model. Target a header range and
`insert_ooxml` a WordArt fragment.

**Page number in a footer.** Write placeholder text with
`edit_page_layout("header_footer", side="footer", which="primary", text="NN")`,
insert the field over it with `edit_fields("insert", target={"text": "NN",
"container": "footer"}, field="page")`, then prepend the surrounding words with
`edit_text` at the returned token, because the field insert consumes the text it
was aimed at. Footer and header offsets are story-relative, so these targets land
in the footer rather than the body.

**Table of contents and table of figures.** `edit_fields("toc", target={...})` at the
insertion point, with `toc_options` keys `upper_heading_level` and
`lower_heading_level`. Collect captions with `edit_fields("table_of_figures",
toc_options={"label": "Figure"})`. Refresh both with `edit_fields("update_all")`.

**Cross-reference.** Captions are SEQ fields. No tool inserts a reference (REF)
field — `edit_fields` insert supports only `page`, `date`, and `sequence` — so
`get_ooxml` an existing reference and re-insert its markup with `insert_ooxml`.
Caption numbering that follows the document scheme renumbers itself.

**Numbering a list.** `manage_list("apply", list_format="number")` then
`manage_list("restart", start_at=3, level=1)`. The returned `restarted_at` is the
value Word actually used, which is worth checking rather than assuming. `restart`
sets `StartAt` on the shared number gallery template, which lasts for the session.

**Formatting audit.** `inspect_document` for the style inventory, then
`get_formatting` per paragraph and read `deviations_from_style`. Clear the drift
with `format_text(action="clear")`, or re-apply the style with
`manage_style("apply", name=...)`.

**Insert text that matches the template.** Find the paragraph you are extending,
then `edit_text(..., action="insert", position="end", format_mode="style_only")`. The
operation result reports the resulting style, so confirm it is the one you wanted.
