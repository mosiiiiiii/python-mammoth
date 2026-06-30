import os
import re
import shutil
import subprocess
import tempfile
import weakref
import zipfile

# Caches keyed by the source package object: {part_path: png_bytes_or_None}.
# A weak key means the cache entry disappears once the document is done with,
# so processing many documents in a pipeline does not leak memory.
_cache = weakref.WeakKeyDictionary()
_media_cache = weakref.WeakKeyDictionary()

_SOFFICE_CANDIDATES = [
    "soffice",
    "/usr/bin/soffice",
    "/usr/local/bin/soffice",
    "/opt/homebrew/bin/soffice",
]

_RENDER_DPI = 300

# EMU per twip (1 inch = 914400 EMU = 1440 twips).
_EMU_PER_TWIP = 635
# Fallback extent (6" x 3.5") for charts whose size we cannot determine.
_DEFAULT_EXTENT = (5486400, 3200400)


def render_chart_png(zip_file, chart_part_path):
    try:
        rendered = _cache.get(zip_file)
        if rendered is None:
            rendered = _render_all_charts(zip_file)
            _cache[zip_file] = rendered
        return rendered.get(chart_part_path)
    except Exception:
        return None


def render_media_to_png(zip_file, media_path):
    try:
        rendered = _media_cache.get(zip_file)
        if rendered is None:
            rendered = _render_all_media(zip_file)
            _media_cache[zip_file] = rendered
        return rendered.get(media_path)
    except Exception:
        return None


def available():
    return _find_soffice() is not None and shutil.which("pdftoppm") is not None


def _render_all_media(zip_file):
    soffice = _find_soffice()
    if soffice is None or shutil.which("pdftoppm") is None:
        return {}

    try:
        media_paths = [
            name for name in zip_file.names() if name.lower().endswith((".emf", ".wmf"))
        ]
        if not media_paths:
            return {}
    except Exception:
        return {}

    work_dir = tempfile.mkdtemp(prefix="mammoth-media-")
    try:
        index_to_media = {}
        input_paths = []
        for index, media_path in enumerate(media_paths):
            ext = os.path.splitext(media_path)[1].lower()
            local = os.path.join(work_dir, "media_{0}{1}".format(index, ext))
            with open(local, "wb") as f:
                f.write(zip_file.read_bytes(media_path))
            input_paths.append(local)
            index_to_media[index] = media_path

        _run_soffice_to_pdf(soffice, input_paths, work_dir)

        result = {}
        for index, media_path in index_to_media.items():
            pdf_path = os.path.join(work_dir, "media_{0}.pdf".format(index))
            result[media_path] = (
                _pdf_to_cropped_png(pdf_path) if os.path.exists(pdf_path) else None
            )
        return result
    except Exception:
        return dict((path, None) for path in media_paths)
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


def _render_all_charts(zip_file):
    soffice = _find_soffice()
    if soffice is None or shutil.which("pdftoppm") is None:
        return {}

    try:
        chart_paths = _find_chart_parts(zip_file)
        if not chart_paths:
            return {}
        original_bytes = _read_package_bytes(zip_file)
    except Exception:
        return {}
    # Map each chart part to the relationship id it is referenced by in the main
    # document, so the mini-document can reuse the original relationships (and
    # therefore the original theme/colours/styles).
    chart_path_to_rid = _read_chart_relationship_ids(original_bytes)
    # The drawing's own extent (size) in the document: rendering at the real
    # size keeps the chart's proportions and relative text size faithful.
    rid_to_extent = _read_chart_extents(original_bytes.get("word/document.xml"))

    work_dir = tempfile.mkdtemp(prefix="mammoth-charts-")
    try:
        # One mini-docx per chart, named so we can map the output back.
        docx_paths = []
        index_to_chart = {}
        for index, chart_path in enumerate(chart_paths):
            relationship_id = chart_path_to_rid.get(chart_path)
            if relationship_id is None:
                # Chart not referenced from the main document (e.g. header):
                # skip and let the caller fall back to the SVG reconstruction.
                continue
            extent = rid_to_extent.get(relationship_id, _DEFAULT_EXTENT)
            name = "chart_{0}.docx".format(index)
            out_path = os.path.join(work_dir, name)
            with open(out_path, "wb") as f:
                f.write(
                    _build_single_chart_docx(original_bytes, relationship_id, extent)
                )
            docx_paths.append(out_path)
            index_to_chart[index] = chart_path

        _run_soffice_to_pdf(soffice, docx_paths, work_dir)

        result = {}
        for index, chart_path in index_to_chart.items():
            pdf_path = os.path.join(work_dir, "chart_{0}.pdf".format(index))
            result[chart_path] = (
                _pdf_to_cropped_png(pdf_path) if os.path.exists(pdf_path) else None
            )
        return result
    except Exception:
        return dict((path, None) for path in chart_paths)
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


def _find_chart_parts(zip_file):
    pattern = re.compile(r"^word/charts/chart\d+\.xml$")
    names = [name for name in zip_file.names() if pattern.match(name)]
    return sorted(names, key=lambda name: int(re.search(r"(\d+)", name).group(1)))


def _read_package_bytes(zip_file):
    return dict((name, zip_file.read_bytes(name)) for name in zip_file.names())


def _read_chart_relationship_ids(original_entries):
    rels = original_entries.get("word/_rels/document.xml.rels")
    if rels is None:
        return {}
    text = rels.decode("utf-8", "replace")
    mapping = {}
    for match in re.finditer(r"<Relationship\b[^>]*>", text):
        tag = match.group(0)
        if "/chart" not in tag:
            continue
        rid = re.search(r'Id="([^"]+)"', tag)
        target = re.search(r'Target="([^"]+)"', tag)
        if rid and target and "charts/chart" in target.group(1):
            normalized = "word/" + target.group(1).lstrip("/").replace("../", "")
            mapping[normalized] = rid.group(1)
    return mapping


def _read_chart_extents(document_xml):
    if document_xml is None:
        return {}
    text = document_xml.decode("utf-8", "replace")
    extents = {}
    # For each chart reference, find the nearest preceding wp:extent.
    extent_re = re.compile(r'<wp:extent\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"')
    for chart_match in re.finditer(r'<c:chart\b[^>]*\br:id="(rId\d+)"', text):
        rid = chart_match.group(1)
        preceding = text[: chart_match.start()]
        last_extent = None
        for last_extent in extent_re.finditer(preceding):
            pass
        if last_extent is not None:
            extents[rid] = (int(last_extent.group(1)), int(last_extent.group(2)))
    return extents


def _build_single_chart_docx(original_entries, relationship_id, extent):
    cx, cy = extent
    # Generous headroom guarantees the chart stays on a single page regardless of
    # line/paragraph overhead; the surrounding whitespace is trimmed by autocrop.
    page_width = cx // _EMU_PER_TWIP + 200
    page_height = cy // _EMU_PER_TWIP + 600
    document_xml = _MINI_DOCUMENT_XML.format(
        relationship_id=relationship_id,
        cx=cx,
        cy=cy,
        page_width=page_width,
        page_height=page_height,
    )

    buffer = _BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as out:
        for name, data in original_entries.items():
            if name == "word/document.xml":
                out.writestr(name, document_xml)
            else:
                out.writestr(name, data)
    return buffer.getvalue()


def _run_soffice_to_pdf(soffice, docx_paths, out_dir):
    profile = os.path.join(out_dir, "lo-profile")
    subprocess.run(
        [
            soffice,
            "--headless",
            "--invisible",
            "--nodefault",
            "--norestore",
            "--nolockcheck",
            "--nologo",
            "-env:UserInstallation=file://{0}".format(profile),
            "--convert-to",
            "pdf",
            "--outdir",
            out_dir,
        ]
        + docx_paths,
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        timeout=300,
    )


def _pdf_to_cropped_png(pdf_path):
    out_prefix = pdf_path[:-4]
    subprocess.run(
        [
            "pdftoppm",
            "-png",
            "-r",
            str(_RENDER_DPI),
            "-f",
            "1",
            "-l",
            "1",
            pdf_path,
            out_prefix,
        ],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        timeout=120,
    )
    # pdftoppm appends "-1" (or "-01") to the prefix.
    candidates = [
        out_prefix + "-1.png",
        out_prefix + "-01.png",
        out_prefix + "-1.ppm",
    ]
    png_path = next((c for c in candidates if os.path.exists(c)), None)
    if png_path is None:
        return None

    with open(png_path, "rb") as f:
        raw = f.read()
    return _autocrop(raw)


def _autocrop(png_bytes):
    try:
        from PIL import Image, ImageChops
    except ImportError:
        return png_bytes

    import io

    image = Image.open(io.BytesIO(png_bytes)).convert("RGB")
    background = Image.new("RGB", image.size, (255, 255, 255))
    diff = ImageChops.difference(image, background)
    bbox = diff.getbbox()
    if bbox:
        padding = 8
        left = max(bbox[0] - padding, 0)
        top = max(bbox[1] - padding, 0)
        right = min(bbox[2] + padding, image.width)
        bottom = min(bbox[3] + padding, image.height)
        image = image.crop((left, top, right, bottom))

    out = io.BytesIO()
    image.save(out, format="PNG")
    return out.getvalue()


def _find_soffice():
    for candidate in _SOFFICE_CANDIDATES:
        resolved = (
            shutil.which(candidate)
            if os.path.basename(candidate) == candidate
            else candidate
        )
        if resolved and os.path.exists(resolved):
            return resolved
    return None


try:
    from io import BytesIO as _BytesIO
except ImportError:  # pragma: no cover
    from StringIO import StringIO as _BytesIO


_MINI_DOCUMENT_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    "<w:document "
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">'
    "<w:body>"
    '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    "<w:r><w:drawing>"
    '<wp:inline distT="0" distB="0" distL="0" distR="0">'
    '<wp:extent cx="{cx}" cy="{cy}"/>'
    '<wp:docPr id="1" name="chart"/>'
    "<wp:cNvGraphicFramePr/>"
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
    '<c:chart r:id="{relationship_id}"/>'
    "</a:graphicData></a:graphic>"
    "</wp:inline>"
    "</w:drawing></w:r></w:p>"
    "<w:sectPr>"
    '<w:pgSz w:w="{page_width}" w:h="{page_height}"/>'
    '<w:pgMar w:top="60" w:right="60" w:bottom="60" w:left="60" '
    'w:header="0" w:footer="0" w:gutter="0"/>'
    "</w:sectPr>"
    "</w:body></w:document>"
)
