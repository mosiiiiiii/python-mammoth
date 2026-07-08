from .docx import charts_xml


# A colour-blind-friendly-ish categorical palette.
_PALETTE = [
    "#4e79a7", "#f28e2b", "#e15759", "#76b7b2", "#59a14f",
    "#edc948", "#b07aa1", "#ff9da7", "#9c755f", "#bab0ac",
]


def render_svg(chart):
    if chart.kind == charts_xml.KIND_PIE:
        return _render_pie(chart)
    return _render_bar(chart)


def text_summary(chart):
    kind = {charts_xml.KIND_PIE: "Pie chart", charts_xml.KIND_BAR: "Bar chart"}.get(
        chart.kind, "Chart",
    )
    parts = [kind] if not chart.title else ["{0}: {1}".format(kind, chart.title)]

    for series in chart.series:
        # Pie slices are shares of a whole: percentages read better than raw values.
        total = None
        if chart.kind == charts_xml.KIND_PIE:
            total = sum(value for value in series.values if value is not None) or None
        pairs = []
        for index, value in enumerate(series.values):
            if value is None:
                continue
            label = _label(chart.categories[index]) if index < len(chart.categories) else str(index + 1)
            if total is not None:
                pairs.append("{0}={1:.2%}".format(label, value / total))
            else:
                pairs.append("{0}={1}".format(label, _fmt(value, precise=True)))
        if pairs:
            prefix = "{0}: ".format(series.name) if series.name else ""
            parts.append(prefix + ", ".join(pairs))

    return ". ".join(parts) + "."


# -- rendering -------------------------------------------------------------


def _render_bar(chart):
    width, height = 640, 400
    margin_left, margin_right = 56, 16
    margin_top = 44 if chart.title else 20
    series = [s for s in chart.series if any(v is not None for v in s.values)]
    has_legend = len(series) > 1
    margin_bottom = 64 + (20 if has_legend else 0)

    plot_left = margin_left
    plot_right = width - margin_right
    plot_top = margin_top
    plot_bottom = height - margin_bottom
    plot_width = plot_right - plot_left
    plot_height = plot_bottom - plot_top

    categories = chart.categories or [str(i + 1) for i in range(_max_len(series))]
    n_cat = max(len(categories), 1)

    values = [v for s in series for v in s.values if v is not None]
    data_max = max(values) if values else 1
    data_min = min(values) if values else 0
    y_max = _nice_ceil(max(data_max, 0))
    y_min = min(data_min, 0)
    if y_max == y_min:
        y_max = y_min + 1

    def y_of(value):
        return plot_bottom - (value - y_min) / (y_max - y_min) * plot_height

    svg = _Svg(width, height)
    if chart.title:
        svg.title(width / 2, 26, chart.title)

    # y gridlines + labels
    for tick in _ticks(y_min, y_max, 5):
        y = y_of(tick)
        svg.line(plot_left, y, plot_right, y, "#e6e6e6")
        svg.text(plot_left - 8, y + 4, _fmt(tick), anchor="end", size=11, fill="#555")
    svg.line(plot_left, y_of(0), plot_right, y_of(0), "#999")

    group_width = plot_width / n_cat
    n_series = max(len(series), 1)
    bar_gap = group_width * 0.2
    bar_width = (group_width - bar_gap) / n_series

    for cat_index in range(n_cat):
        group_x = plot_left + cat_index * group_width
        for series_index, s in enumerate(series):
            value = s.values[cat_index] if cat_index < len(s.values) else None
            if value is None:
                continue
            x = group_x + bar_gap / 2 + series_index * bar_width
            top = y_of(max(value, 0))
            bottom = y_of(min(value, 0))
            svg.rect(x, top, max(bar_width - 2, 1), max(bottom - top, 1),
                     _PALETTE[series_index % len(_PALETTE)])
        label = _label(categories[cat_index]) if cat_index < len(categories) else ""
        svg.text(group_x + group_width / 2, plot_bottom + 16, _ellipsize(label, 14),
                 anchor="middle", size=11, fill="#333")

    if has_legend:
        _legend(svg, series, width, height - margin_bottom + 34)

    return svg.render()


def _render_pie(chart):
    width, height = 640, 360
    margin_top = 44 if chart.title else 16
    cx, cy = 170, margin_top + (height - margin_top) / 2
    radius = min(cx - 16, (height - margin_top) / 2 - 12)

    series = chart.series[0] if chart.series else charts_xml.Series("", [])
    labels = chart.categories or [str(i + 1) for i in range(len(series.values))]
    pairs = [
        (_label(labels[i]) if i < len(labels) else str(i + 1), v)
        for i, v in enumerate(series.values)
        if v is not None and v > 0
    ]
    total = sum(v for _, v in pairs) or 1

    svg = _Svg(width, height)
    if chart.title:
        svg.title(width / 2, 26, chart.title)

    angle = -90.0  # start at 12 o'clock
    legend_items = []
    for index, (label, value) in enumerate(pairs):
        sweep = value / total * 360
        colour = _PALETTE[index % len(_PALETTE)]
        svg.slice(cx, cy, radius, angle, angle + sweep, colour)
        angle += sweep
        legend_items.append((
            "{0} ({1:.0%})".format(_ellipsize(label, 22), value / total),
            colour,
        ))

    legend_x = cx + radius + 32
    legend_y = margin_top + 8
    for text, colour in legend_items:
        svg.rect(legend_x, legend_y, 12, 12, colour)
        svg.text(legend_x + 18, legend_y + 11, text, anchor="start", size=12, fill="#333")
        legend_y += 22

    return svg.render()


def _legend(svg, series, width, y):
    # Lay out horizontally, centred.
    entries = [(s.name or "Serie {0}".format(i + 1), _PALETTE[i % len(_PALETTE)])
               for i, s in enumerate(series)]
    item_width = min(180, max(80, width / max(len(entries), 1)))
    total_width = item_width * len(entries)
    x = (width - total_width) / 2
    for text, colour in entries:
        svg.rect(x, y - 10, 12, 12, colour)
        svg.text(x + 18, y, _ellipsize(text, 20), anchor="start", size=12, fill="#333")
        x += item_width


# -- tiny SVG builder ------------------------------------------------------


class _Svg(object):
    def __init__(self, width, height):
        self._width = width
        self._height = height
        self._parts = []

    def line(self, x1, y1, x2, y2, stroke):
        self._parts.append(
            '<line x1="{0:.1f}" y1="{1:.1f}" x2="{2:.1f}" y2="{3:.1f}" stroke="{4}" stroke-width="1"/>'
            .format(x1, y1, x2, y2, stroke))

    def rect(self, x, y, w, h, fill):
        self._parts.append(
            '<rect x="{0:.1f}" y="{1:.1f}" width="{2:.1f}" height="{3:.1f}" fill="{4}"/>'
            .format(x, y, w, h, fill))

    def text(self, x, y, value, anchor="start", size=12, fill="#000"):
        self._parts.append(
            '<text x="{0:.1f}" y="{1:.1f}" text-anchor="{2}" font-family="sans-serif" '
            'font-size="{3}" fill="{4}">{5}</text>'
            .format(x, y, anchor, size, fill, _escape(value)))

    def title(self, x, y, value):
        self._parts.append(
            '<text x="{0:.1f}" y="{1:.1f}" text-anchor="middle" font-family="sans-serif" '
            'font-size="15" font-weight="bold" fill="#222">{2}</text>'
            .format(x, y, _escape(_ellipsize(value, 64))))

    def slice(self, cx, cy, r, start_deg, end_deg, fill):
        import math
        if end_deg - start_deg >= 359.999:
            # Full circle: <path> arcs can't draw 360 in one go.
            self._parts.append(
                '<circle cx="{0:.1f}" cy="{1:.1f}" r="{2:.1f}" fill="{3}" '
                'stroke="#fff" stroke-width="1"/>'.format(cx, cy, r, fill))
            return
        start = math.radians(start_deg)
        end = math.radians(end_deg)
        x1, y1 = cx + r * math.cos(start), cy + r * math.sin(start)
        x2, y2 = cx + r * math.cos(end), cy + r * math.sin(end)
        large_arc = 1 if (end_deg - start_deg) > 180 else 0
        self._parts.append(
            '<path d="M{0:.1f},{1:.1f} L{2:.2f},{3:.2f} A{4:.1f},{4:.1f} 0 {5} 1 {6:.2f},{7:.2f} Z" '
            'fill="{8}" stroke="#fff" stroke-width="1"/>'
            .format(cx, cy, x1, y1, r, large_arc, x2, y2, fill))

    def render(self):
        return (
            '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {0} {1}" '
            'width="{0}" height="{1}" role="img">'
            '<rect width="{0}" height="{1}" fill="#ffffff"/>{2}</svg>'
        ).format(self._width, self._height, "".join(self._parts))


# -- helpers ---------------------------------------------------------------


def _max_len(series):
    return max((len(s.values) for s in series), default=0)


def _ticks(low, high, count):
    step = (high - low) / count
    return [low + step * i for i in range(count + 1)]


def _nice_ceil(value):
    if value <= 0:
        return 1
    import math
    magnitude = 10 ** math.floor(math.log10(value))
    for factor in (1, 1.5, 2, 2.5, 5, 10):
        candidate = factor * magnitude
        if candidate >= value:
            return candidate
    return 10 * magnitude


def _fmt(value, precise=False):
    if value is None:
        return ""
    if precise:
        return ("%.4g" % value)
    if value == int(value):
        return str(int(value))
    return ("%.2f" % value).rstrip("0").rstrip(".")


def _label(value):
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return "" if value is None else str(value)


def _ellipsize(text, max_length):
    text = "" if text is None else str(text)
    if len(text) <= max_length:
        return text
    return text[: max_length - 1] + "…"


def _escape(text):
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )
