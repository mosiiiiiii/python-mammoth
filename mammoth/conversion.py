# coding=utf-8

from __future__ import unicode_literals

from functools import partial

import cobble

from . import documents, results, html_paths, images, writers, html
from .docx.files import InvalidFileReferenceError
from .lists import find_index


def convert_document_element_to_html(element,
        style_map=None,
        convert_image=None,
        id_prefix=None,
        output_format=None,
        ignore_empty_paragraphs=True):

    if style_map is None:
        style_map = []

    if id_prefix is None:
        id_prefix = ""

    if convert_image is None:
        convert_image = images.data_uri

    if isinstance(element, documents.Document):
        comments = dict(
            (comment.comment_id, comment)
            for comment in element.comments
        )
    else:
        comments = {}

    messages = []
    converter = _DocumentConverter(
        messages=messages,
        style_map=style_map,
        convert_image=convert_image,
        id_prefix=id_prefix,
        ignore_empty_paragraphs=ignore_empty_paragraphs,
        note_references=[],
        comments=comments,
    )
    context = _ConversionContext(is_table_header=False)
    nodes = converter.visit(element, context)

    writer = writers.writer(output_format)
    html.write(writer, html.collapse(html.strip_empty(nodes)))
    return results.Result(writer.as_string(), messages)


@cobble.data
class _ConversionContext(object):
    is_table_header = cobble.field()

    def copy(self, **kwargs):
        return cobble.copy(self, **kwargs)


class _DocumentConverter(documents.element_visitor(args=1)):
    def __init__(self, messages, style_map, convert_image, id_prefix, ignore_empty_paragraphs, note_references, comments):
        self._messages = messages
        self._style_map = style_map
        self._id_prefix = id_prefix
        self._ignore_empty_paragraphs = ignore_empty_paragraphs
        self._note_references = note_references
        self._referenced_comments = []
        self._convert_image = convert_image
        self._comments = comments
        # Track list counters so we can emit <li value="N"> when a list is
        # interrupted and later resumed. The key contains the numbering
        # format and the nesting level so independent lists don’t interfere.
        self._list_counters = {}
        # Track counters for numbered headings (outline numbering) so we can
        # recreate prefixes such as “1.”, “6.2”, … when headings are also
        # part of a numbered list/outline. Keys are the numbering definition
        # (abstract_num_id or num_id) so independent numbering sequences do
        # not interfere with each other. Each value is another mapping of
        # level -> current counter.
        self._heading_counters = {}

    def visit_image(self, image, context):
        try:
            return self._convert_image(image)
        except InvalidFileReferenceError as error:
            self._messages.append(results.warning(str(error)))
            return []

    def visit_document(self, document, context):
        nodes = self._visit_all(document.children, context)
        notes = [
            document.notes.resolve(reference)
            for reference in self._note_references
        ]
        notes_list = html.element("ol", {}, self._visit_all(notes, context))
        comments = html.element("dl", {}, [
            html_node
            for referenced_comment in self._referenced_comments
            for html_node in self.visit_comment(referenced_comment, context)
        ])
        return nodes + [notes_list, comments]


    def visit_paragraph(self, paragraph, context):
        # Headings can also be numbered in Word using outline numbering. In
        # such cases we *don’t* want to wrap them in <ol>/<li>; instead we
        # keep them as headings (h1, h2, …) but prepend the computed number
        # as plain text so the resulting HTML matches what is shown in Word.

        is_numbered_heading = (
            paragraph.numbering is not None and self._is_heading(paragraph)
        )

        # Normal list paragraphs (that aren’t headings) are still converted
        # to <ul>/<ol>/<li>.
        if paragraph.numbering is not None and not is_numbered_heading:
            html_path = self._list_html_path(paragraph.numbering)

            def children():
                content = self._visit_all(paragraph.children, context)
                if self._ignore_empty_paragraphs:
                    return content
                else:
                    return [html.force_write] + content

            html_path = self._add_alignment_to_path(html_path, paragraph.alignment)
            return html_path.wrap(children)

        def children():
            content = self._visit_all(paragraph.children, context)

            # Prepend numbering prefix for numbered headings so that the HTML
            # contains the visible outline numbers (e.g. “6.2”).
            if is_numbered_heading:
                prefix = self._heading_number_prefix(paragraph.numbering)
                if prefix:
                    content = [html.text(prefix + " ")] + content

            if self._ignore_empty_paragraphs:
                return content
            else:
                return [html.force_write] + content

        html_path = self._find_html_path_for_paragraph(paragraph)
        html_path = self._add_alignment_to_path(html_path, paragraph.alignment)
        return html_path.wrap(children)


    def visit_run(self, run, context):
        nodes = lambda: self._visit_all(run.children, context)
        paths = []
        if run.highlight is not None:
            style = self._find_style(Highlight(color=run.highlight), "highlight")
            if style is not None:
                paths.append(style.html_path)
        if run.is_small_caps:
            paths.append(self._find_style_for_run_property("small_caps"))
        if run.is_all_caps:
            paths.append(self._find_style_for_run_property("all_caps"))
        if run.is_strikethrough:
            paths.append(self._find_style_for_run_property("strikethrough", default="s"))
        if run.is_underline:
            paths.append(self._find_style_for_run_property("underline"))
        if run.vertical_alignment == documents.VerticalAlignment.subscript:
            paths.append(html_paths.element(["sub"], fresh=False))
        if run.vertical_alignment == documents.VerticalAlignment.superscript:
            paths.append(html_paths.element(["sup"], fresh=False))
        if run.is_italic:
            paths.append(self._find_style_for_run_property("italic", default="em"))
        if run.is_bold:
            paths.append(self._find_style_for_run_property("bold", default="strong"))
        paths.append(self._find_html_path_for_run(run))

        for path in paths:
            nodes = partial(path.wrap, nodes)

        return nodes()


    def _find_style_for_run_property(self, element_type, default=None):
        style = self._find_style(None, element_type)
        if style is not None:
            return style.html_path
        elif default is not None:
            return html_paths.element(default, fresh=False)
        else:
            return html_paths.empty


    def visit_text(self, text, context):
        return [html.text(text.value)]


    def visit_hyperlink(self, hyperlink, context):
        if hyperlink.anchor is None:
            href = hyperlink.href
        else:
            href = "#{0}".format(self._html_id(hyperlink.anchor))

        attributes = {"href": href}
        if hyperlink.target_frame is not None:
            attributes["target"] = hyperlink.target_frame

        nodes = self._visit_all(hyperlink.children, context)
        return [html.collapsible_element("a", attributes, nodes)]


    def visit_checkbox(self, checkbox, context):
        attributes = {"type": "checkbox"}

        if checkbox.checked:
            attributes["checked"] = "checked"

        return [html.element("input", attributes)]


    def visit_bookmark(self, bookmark, context):
        element = html.collapsible_element(
            "a",
            {"id": self._html_id(bookmark.name)},
            [html.force_write])
        return [element]


    def visit_tab(self, tab, context):
        return [html.text("\t")]

    _default_table_path = html_paths.path([html_paths.element(["table"], fresh=True)])

    def visit_table(self, table, context):
        return self._find_html_path(table, "table", self._default_table_path) \
            .wrap(lambda: self._convert_table_children(table, context))

    def _convert_table_children(self, table, context):
        body_index = find_index(
            lambda child: not isinstance(child, documents.TableRow) or not child.is_header,
            table.children,
        )
        if body_index is None:
            body_index = len(table.children)

        if body_index == 0:
            children = self._visit_all(table.children, context.copy(is_table_header=False))
        else:
            head_rows = self._visit_all(table.children[:body_index], context.copy(is_table_header=True))
            body_rows = self._visit_all(table.children[body_index:], context.copy(is_table_header=False))
            children = [
                html.element("thead", {}, head_rows),
                html.element("tbody", {}, body_rows),
            ]

        return [html.force_write] + children


    def visit_table_row(self, table_row, context):
        return [html.element("tr", {}, [html.force_write] + self._visit_all(table_row.children, context))]


    def visit_table_cell(self, table_cell, context):
        if context.is_table_header:
            tag_name = "th"
        else:
            tag_name = "td"
        attributes = {}
        if table_cell.colspan != 1:
            attributes["colspan"] = str(table_cell.colspan)
        if table_cell.rowspan != 1:
            attributes["rowspan"] = str(table_cell.rowspan)
        nodes = [html.force_write] + self._visit_all(table_cell.children, context)
        return [
            html.element(tag_name, attributes, nodes)
        ]


    def visit_break(self, break_, context):
        return self._find_html_path_for_break(break_).wrap(lambda: [])


    def _find_html_path_for_break(self, break_):
        style = self._find_style(break_, "break")
        if style is not None:
            return style.html_path
        elif break_.break_type == "line":
            return html_paths.path([html_paths.element("br", fresh=True)])
        else:
            return html_paths.empty


    def visit_note_reference(self, note_reference, context):
        self._note_references.append(note_reference)
        note_number = len(self._note_references)
        return [
            html.element("sup", {}, [
                html.element("a", {
                    "href": "#" + self._note_html_id(note_reference),
                    "id": self._note_ref_html_id(note_reference),
                }, [html.text("[{0}]".format(note_number))])
            ])
        ]


    def visit_note(self, note, context):
        note_body = self._visit_all(note.body, context) + [
            html.collapsible_element("p", {}, [
                html.text(" "),
                html.element("a", {"href": "#" + self._note_ref_html_id(note)}, [
                    html.text(_up_arrow)
                ]),
            ])
        ]
        return [
            html.element("li", {"id": self._note_html_id(note)}, note_body)
        ]


    def visit_comment_reference(self, reference, context):
        def nodes():
            comment = self._comments[reference.comment_id]
            count = len(self._referenced_comments) + 1
            label = "[{0}{1}]".format(_comment_author_label(comment), count)
            self._referenced_comments.append((label, comment))
            return [
                # TODO: remove duplication with note references
                html.element("a", {
                    "href": "#" + self._referent_html_id("comment", reference.comment_id),
                    "id": self._reference_html_id("comment", reference.comment_id),
                }, [html.text(label)])
            ]

        html_path = self._find_html_path(
            None,
            "comment_reference",
            default=html_paths.ignore,
        )

        return html_path.wrap(nodes)

    def visit_comment(self, referenced_comment, context):
        label, comment = referenced_comment
        # TODO remove duplication with notes
        body = self._visit_all(comment.body, context) + [
            html.collapsible_element("p", {}, [
                html.text(" "),
                html.element("a", {"href": "#" + self._reference_html_id("comment", comment.comment_id)}, [
                    html.text(_up_arrow)
                ]),
            ])
        ]
        return [
            html.element(
                "dt",
                {"id": self._referent_html_id("comment", comment.comment_id)},
                [html.text("Comment {0}".format(label))],
            ),
            html.element("dd", {}, body),
        ]


    def _visit_all(self, elements, context):
        return [
            html_node
            for element in elements
            for html_node in self.visit(element, context)
        ]


    def _find_html_path_for_paragraph(self, paragraph):
        # List paragraphs are handled separately in visit_paragraph.
        if paragraph.numbering is not None and not self._is_heading(paragraph):
            return self._list_html_path(paragraph.numbering)

        default = html_paths.path([html_paths.element("p", fresh=True)])
        return self._find_html_path(paragraph, "paragraph", default, warn_unrecognised=True)

    def _find_html_path_for_run(self, run):
        return self._find_html_path(run, "run", default=html_paths.empty, warn_unrecognised=True)

    def _list_html_path(self, numbering):
        """Return an HtmlPath that produces the correct <ul>/<ol>/<li>
        hierarchy for the given Word numbering level."""

        level = int(numbering.level_index)
        ordered = numbering.is_ordered

        list_tag = "ol" if ordered else "ul"

        list_attributes = {}

        # For ordered lists, set the HTML ‘type’ attribute when Word uses
        # letters or roman numerals.
        if ordered and numbering.num_fmt:
            type_char = _num_fmt_to_ol_type(numbering.num_fmt)
            if type_char:
                list_attributes["type"] = type_char

        # Increment running counter and decide whether the current <li>
        # needs its explicit value attribute.
        li_attributes = {}
        if ordered:
            key_source = numbering.abstract_num_id or numbering.num_id
            counter_key = (key_source, numbering.level_index)

            start_val = numbering.start or 1
            if counter_key not in self._list_counters:
                 self._list_counters[counter_key] = start_val - 1

            # Increment
            self._list_counters[counter_key] += 1
            current_value = self._list_counters[counter_key]

            # If this is the first list item and start >1, put start on <ol>.
            if current_value == start_val and start_val > 1:
                list_attributes["start"] = str(start_val)
            elif current_value != start_val:
                # Non-first items interrupted list; use value attr to preserve
                li_attributes["value"] = str(current_value)

        elements = []

        # For ancestor levels we don’t know whether they’re ordered or not,
        # so accept both ul and ol.
        for _ in range(level):
            elements.append(html_paths.element(["ul", "ol"], fresh=False))
            elements.append(html_paths.element("li", fresh=False))

        # The actual list for this level
        elements.append(html_paths.element([list_tag], attributes=list_attributes, fresh=False))
        # The list item
        elements.append(html_paths.element("li", attributes=li_attributes, fresh=True))

        return html_paths.path(elements)

    def _add_alignment_to_path(self, html_path, alignment):
        """Return *html_path* with inline text-alignment style applied to the
        innermost element if *alignment* is not None.*

        Word stores paragraph alignment in <w:jc w:val="...">.  Values we
        care about are ``left``, ``center``, ``right`` and ``both`` (justify).
        When converting to HTML we keep left-aligned paragraphs unchanged
        (that is the browser default).  For the other alignments we append a
        ``style="text-align: …;"`` declaration to the innermost element of
        *html_path* (usually <p> or <li>).  If the element already has a
        ``style`` attribute we append to it, otherwise we add a new one.
        """
        if alignment is None:
            return html_path

        ALIGN_MAP = {
            "center": "center",
            "right": "right",
            "both": "justify",
            "justify": "justify",
            "end": "right",
            "start": "left",
            # Treat explicit "left" the same as None (browser default)
        }
        css_value = ALIGN_MAP.get(alignment)
        if css_value in (None, "left"):
            return html_path

        # html_path.elements is an ordered list where the last element is the
        # innermost element (e.g. <p> or <li>).  Create a shallow copy so we
        # don’t mutate the original path shared elsewhere.
        elements = list(html_path.elements)
        if not elements:
            return html_path

        innermost = elements[-1]
        # Copy existing attributes and append/merge style
        attrs = dict(innermost.tag.attributes)
        existing_style = attrs.get("style", "").strip()
        align_declaration = "text-align: {0};".format(css_value)
        if existing_style:
            # Ensure a trailing semicolon for proper separation
            if not existing_style.endswith(";"):
                existing_style += ";"
            attrs["style"] = "{0} {1}".format(existing_style, align_declaration)
        else:
            attrs["style"] = align_declaration

        # Rebuild HtmlPathElement with updated attributes, preserving other
        # Tag properties.
        new_tag = html.tag(
            tag_names=innermost.tag.tag_names,
            attributes=attrs,
            collapsible=innermost.tag.collapsible,
            separator=innermost.tag.separator,
        )
        elements[-1] = html_paths.HtmlPathElement(new_tag)
        return html_paths.HtmlPath(elements)

    def _is_heading(self, paragraph):
        """Return True if the paragraph appears to be a heading (style name
        or ID starts with “heading”)."""
        name = (paragraph.style_name or "").lower()
        sid = (paragraph.style_id or "").lower()
        return name.startswith("heading") or sid.startswith("heading")

    def _heading_number_prefix(self, numbering):
        """Return the textual prefix (e.g. “1.”, “6.2”) for a numbered
        heading, keeping counters per numbering instance."""

        level = int(numbering.level_index)

        # Ensure counters dict exists
        counters = self._heading_counters

        # If this heading starts at a deeper level without its parent ever
        # having appeared (common in some documents where the first visible
        # outline level is “1.1 …” rather than “1. …”), synthesise the
        # missing parent counter(s) so the hierarchy is preserved.
        if level > 0 and (level - 1) not in counters:
            counters[level - 1] = 1

        # Increment counter for current level
        counters[level] = counters.get(level, 0) + 1

        # Reset counters for deeper levels so they start fresh whenever we
        # move up in the hierarchy.
        for deeper in list(counters.keys()):
            if deeper > level:
                del counters[deeper]

        # Build prefix string from level 0 up to current level
        parts = [str(counters[i]) for i in range(level + 1) if i in counters]
        if not parts:
            return ""

        prefix = ".".join(parts)

        # Word usually appends a trailing dot to top-level headings only.
        if level == 0:
            prefix += "."

        return prefix

    def _find_html_path(self, element, element_type, default, warn_unrecognised=False):
        """Return the first style mapping that matches *element* or *default*.

        This is the original logic from Mammoth.  It is still used for runs,
        tables, comments, etc., so we add it back unchanged.
        """
        style = self._find_style(element, element_type)
        if style is not None:
            return style.html_path

        if warn_unrecognised and getattr(element, "style_id", None) is not None:
            self._messages.append(results.warning(
                "Unrecognised {0} style: {1} (Style ID: {2})".format(
                    element_type, element.style_name, element.style_id)
            ))

        return default

    def _find_style(self, element, element_type):
        for style in self._style_map:
            document_matcher = style.document_matcher
            if _document_matcher_matches(document_matcher, element, element_type):
                return style

    def _note_html_id(self, note):
        return self._referent_html_id(note.note_type, note.note_id)

    def _note_ref_html_id(self, note):
        return self._reference_html_id(note.note_type, note.note_id)

    def _referent_html_id(self, reference_type, reference_id):
        return self._html_id("{0}-{1}".format(reference_type, reference_id))

    def _reference_html_id(self, reference_type, reference_id):
        return self._html_id("{0}-ref-{1}".format(reference_type, reference_id))

    def _html_id(self, suffix):
        return "{0}{1}".format(self._id_prefix, suffix)


@cobble.data
class Highlight:
    color = cobble.field()


def _num_fmt_to_ol_type(num_fmt):
    """Map Word’s numFmt names to the HTML ‘type’ attribute values."""
    mapping = {
        "decimal": None,  # default
        "lowerLetter": "a",
        "upperLetter": "A",
        "lowerRoman": "i",
        "upperRoman": "I",
    }
    return mapping.get(num_fmt)


def _document_matcher_matches(matcher, element, element_type):
    if matcher.element_type in ["underline", "strikethrough", "all_caps", "small_caps", "bold", "italic", "comment_reference"]:
        return matcher.element_type == element_type
    elif matcher.element_type == "highlight":
        return (
            matcher.element_type == element_type and
            (matcher.color is None or matcher.color == element.color)
        )
    elif matcher.element_type == "break":
        return (
            matcher.element_type == element_type and
            matcher.break_type == element.break_type
        )
    else: # matcher.element_type in ["paragraph", "run"]:
        return (
            matcher.element_type == element_type and (
                matcher.style_id is None or
                matcher.style_id == element.style_id
            ) and (
                matcher.style_name is None or
                element.style_name is not None and (matcher.style_name.matches(element.style_name))
            ) and (
                element_type != "paragraph" or
                matcher.numbering is None or
                _numbering_matches(matcher.numbering, element.numbering)
            )
        )


def _numbering_matches(matcher_numbering, element_numbering):
    """Return True when the list level and ordered/unordered flag match.
    Other properties (num_fmt, counters, …) are ignored so that existing
    style-map rules stay compatible after we added new fields."""

    if matcher_numbering is None or element_numbering is None:
        return matcher_numbering is element_numbering

    return (
        matcher_numbering.level_index == element_numbering.level_index and
        matcher_numbering.is_ordered == element_numbering.is_ordered
    )


def _comment_author_label(comment):
    return comment.author_initials or ""


_up_arrow = "↑"
