import collections

Chart = collections.namedtuple("Chart", ["title", "kind", "categories", "series"])
Series = collections.namedtuple("Series", ["name", "values"])

# Chart kinds we know how to render. Everything else falls back to "bar".
KIND_BAR = "bar"
KIND_PIE = "pie"


def read_chart(chart_space_element):
    chart = chart_space_element.find_child_or_null("c:chart")
    plot_area = chart.find_child_or_null("c:plotArea")

    title = _read_title(chart.find_child("c:title"))

    kind = KIND_BAR
    plot = None
    for child in plot_area.children:
        name = getattr(child, "name", None)
        if name in ("c:pieChart", "c:pie3DChart", "c:doughnutChart"):
            kind, plot = KIND_PIE, child
            break
        if name in (
            "c:barChart",
            "c:bar3DChart",
            "c:lineChart",
            "c:line3DChart",
            "c:areaChart",
            "c:area3DChart",
        ):
            kind, plot = KIND_BAR, child
            # Keep looking in case a pie shares the plot area, but a bar/line is
            # a perfectly good default.

    if plot is None:
        return Chart(title=title, kind=kind, categories=[], series=[])

    series = [_read_series(ser) for ser in plot.find_children("c:ser")]

    # Categories are usually identical across series; take the first non-empty.
    categories = []
    for ser in plot.find_children("c:ser"):
        categories = _read_points(ser.find_child_or_null("c:cat"))
        if categories:
            break

    return Chart(
        title=title,
        kind=kind,
        categories=categories,
        series=[Series(name=name, values=values) for name, values in series],
    )


def _read_series(ser_element):
    name = _read_title(ser_element.find_child("c:tx"))
    values = _read_points(ser_element.find_child_or_null("c:val"))
    return name, values


def _read_points(container):
    for ref_name, cache_name in (
        ("c:numRef", "c:numCache"),
        ("c:strRef", "c:strCache"),
    ):
        ref = container.find_child(ref_name)
        if ref is not None:
            return _read_cache(
                ref.find_child_or_null(cache_name), numeric=(ref_name == "c:numRef")
            )

    multi_level = container.find_child("c:multiLvlStrRef")
    if multi_level is not None:
        level = multi_level.find_child_or_null("c:multiLvlStrCache").find_child("c:lvl")
        if level is not None:
            return _read_pts(level, numeric=False)

    # Literal (non-referenced) caches.
    for cache_name, numeric in (("c:numLit", True), ("c:strLit", False)):
        cache = container.find_child(cache_name)
        if cache is not None:
            return _read_pts(cache, numeric=numeric)

    return []


def _read_cache(cache_element, numeric):
    return _read_pts(cache_element, numeric=numeric)


def _read_pts(cache_element, numeric):
    point_count_element = cache_element.find_child_or_null("c:ptCount")
    point_count = _to_int(point_count_element.attributes.get("val"))

    by_index = {}
    max_index = -1
    for pt in cache_element.find_children("c:pt"):
        index = _to_int(pt.attributes.get("idx"))
        if index is None:
            continue
        raw = _inner_text(pt.find_child_or_null("c:v"))
        by_index[index] = _to_float(raw) if numeric else raw
        max_index = max(max_index, index)

    length = point_count if point_count is not None else max_index + 1
    default = None if numeric else ""
    return [by_index.get(i, default) for i in range(length)]


def _read_title(title_element):
    if title_element is None:
        return None

    rich = _find_descendant(title_element, "c:rich")
    if rich is not None:
        text = "".join(_inner_text(t) for t in _find_descendants(rich, "a:t")).strip()
        if text:
            return text

    str_ref = _find_descendant(title_element, "c:strRef")
    if str_ref is not None:
        cache = str_ref.find_child_or_null("c:strCache")
        values = [
            _inner_text(pt.find_child_or_null("c:v"))
            for pt in cache.find_children("c:pt")
        ]
        text = " ".join(v for v in values if v).strip()
        if text:
            return text

    value = _find_descendant(title_element, "c:v")
    if value is not None:
        return _inner_text(value).strip() or None

    return None


def _find_descendant(element, name):
    for child in getattr(element, "children", []):
        if getattr(child, "name", None) == name:
            return child
        found = _find_descendant(child, name)
        if found is not None:
            return found
    return None


def _find_descendants(element, name):
    found = []
    for child in getattr(element, "children", []):
        if getattr(child, "name", None) == name:
            found.append(child)
        found.extend(_find_descendants(child, name))
    return found


def _inner_text(element):
    if element is None:
        return ""
    if getattr(element, "node_type", None) == 3:  # text node
        return element.value
    return "".join(_inner_text(child) for child in getattr(element, "children", []))


def _to_int(value):
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _to_float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None
