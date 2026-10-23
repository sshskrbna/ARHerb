"""Turning spreadsheet rows into labels in a Word document.

A label template is an ordinary .docx: its body is the design of ONE label and
contains placeholders like {species} or {qr}. Every row of the spreadsheet gets
its own copy of that body, filled in, and the copies are laid out in a grid.
"""
import re
from copy import deepcopy
from datetime import datetime
from io import BytesIO

import pandas as pd
import qrcode
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.text.paragraph import Paragraph
from docx.text.run import Run


FIELDS = ["family", "familyrus", "species", "speciesrus", "region", "point", "habitats",
          "date", "leg.", "det.", "num", "N", "E", "inat"]

# Darwin Core names that are understood when the default column name is absent
DWC_ALIASES = {
    "family": ["family"],
    "species": ["scientificName"],
    "speciesrus": ["vernacularName"],
    "region": ["stateProvince"],
    "point": ["locality"],
    "habitats": ["habitat"],
    "date": ["eventDate"],
    "leg.": ["recordedBy"],
    "det.": ["identifiedBy"],
    "num": ["catalogNumber", "recordNumber"],
    "N": ["decimalLatitude"],
    "E": ["decimalLongitude"],
}

QR_TARGETS = {
    "google": "Google Maps",
    "osm": "OpenStreetMap",
    "geo": "Map app (geo:)",
}

PLACEHOLDERS = {
    "herbarium": "Herbarium name entered in the app",
    "family": "Family", "familyrus": "Family in Russian (vascular plants)",
    "species": "Species", "speciesrus": "Species in Russian (vascular plants)",
    "region": "Region", "point": "Locality", "habitats": "Habitat",
    "location": "Region, locality and habitat joined with commas",
    "date": "Collection date", "leg": "Collector", "det": "Identified by", "num": "Specimen number",
    "lat": "Latitude", "lon": "Longitude", "coords": "Coordinates, e.g. 56.0153° N, 37.8471° E",
    "inat": "iNaturalist observation link",
    "qr": "QR code with the collection site", "inat_qr": "QR code with the iNaturalist observation",
}
IMAGE_KEYS = {"qr", "inat_qr"}
QR_WIDTH = Cm(1.4)
PLACEHOLDER = re.compile(r"\{([^{}]+)\}")


def filter_nan(value):
    return value if pd.notna(value) else ''


def format_num(value):
    # pandas reads an integer column with gaps as float: 12 -> 12.0
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def parse_coordinate(value):
    if value == '' or value is None:
        return None
    try:
        return float(str(value).strip().replace(',', '.'))
    except ValueError:
        return None


def remove_time_from_date(date_value):
    if isinstance(date_value, datetime):
        return date_value.date().strftime("%Y-%m-%d")
    elif isinstance(date_value, str):
        try:
            return datetime.strptime(date_value, "%Y-%m-%d %H:%M:%S").strftime("%Y-%m-%d")
        except ValueError:
            return date_value
    return date_value


def format_coordinates(lat, lon):
    def part(value, positive, negative):
        number = f"{abs(value):.5f}".rstrip("0").rstrip(".")
        return f"{number}° {positive if value >= 0 else negative}"
    return f"{part(lat, 'N', 'S')}, {part(lon, 'E', 'W')}"


def map_url(lat, lon, target):
    if target == "osm":
        return f"https://www.openstreetmap.org/?mlat={lat}&mlon={lon}#map=15/{lat}/{lon}"
    if target == "geo":
        return f"geo:{lat},{lon}"
    return f"https://www.google.com/maps?q={lat},{lon}"


def inat_url(value):
    text = format_num(value).strip()
    if "inaturalist.org/observations/" in text:
        return text
    if text.isdigit():
        return f"https://www.inaturalist.org/observations/{text}"
    return ""


def make_qr(data):
    buffer = BytesIO()
    qrcode.make(data).save(buffer, format="PNG")
    buffer.seek(0)
    return buffer


def resolve_columns(columns, custom_map=None):
    """Map each field to the spreadsheet column that holds it (or None)."""
    custom_map = custom_map or {}
    columns = [str(c) for c in columns]
    resolved = {}
    for field in FIELDS:
        if custom_map.get(field):
            resolved[field] = custom_map[field] if custom_map[field] in columns else None
            continue
        candidates = [field] + DWC_ALIASES.get(field, [])
        resolved[field] = next((c for c in candidates if c in columns), None)
    return resolved


def row_values(row, columns, herbarium, herbarium_type, qr_target=None, include_inat=False):
    """Text and images for one label. Every spreadsheet column is available too, by its own name."""
    get = lambda field: filter_nan(row.get(columns[field], '')) if columns.get(field) else ''
    values = {str(name): format_num(remove_time_from_date(filter_nan(value))) for name, value in row.items()}

    vascular = herbarium_type == "Vascular plant"
    region, point, habitats = get('region'), get('point'), get('habitats')
    values.update({
        "herbarium": herbarium,
        "family": get('family'),
        "familyrus": get('familyrus') if vascular else '',
        "species": get('species'),
        "speciesrus": get('speciesrus') if vascular else '',
        "region": region, "point": point, "habitats": habitats,
        "location": ", ".join(str(p) for p in (region, point, habitats) if str(p).strip()),
        "date": remove_time_from_date(get('date')),
        "leg": get('leg.'), "det": get('det.'),
        "num": format_num(get('num')),
        "lat": "", "lon": "", "coords": "", "inat": "",
        "qr": None, "inat_qr": None,
    })
    values = {k: (v if k in IMAGE_KEYS else str(v)) for k, v in values.items()}

    lat, lon = parse_coordinate(get('N')), parse_coordinate(get('E'))
    if lat is not None and lon is not None:
        values.update(lat=f"{lat}", lon=f"{lon}", coords=format_coordinates(lat, lon))
        if qr_target:
            values["qr"] = make_qr(map_url(lat, lon, qr_target))
    link = inat_url(get('inat'))
    values["inat"] = link
    if include_inat and link:
        values["inat_qr"] = make_qr(link)
    return values


def default_template():
    """The built-in label design. Users can save it and edit it in Word."""
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(section, side, Cm(1.5))
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(11)
    style.paragraph_format.space_after = Pt(0)

    body = doc.add_paragraph()
    body.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    body.paragraph_format.space_after = Pt(6)
    body.add_run("{herbarium}").bold = True
    for key in ("family", "familyrus", "species", "speciesrus"):
        doc.add_paragraph().add_run("{%s}" % key).italic = True

    small = []
    for i, text in enumerate(["{location}", "{coords}", "{date}\tleg.: {leg}", "№ {num}\tdet.: {det}"]):
        paragraph = doc.add_paragraph()
        if i == 0:
            paragraph.paragraph_format.space_before = Pt(6)
        paragraph.paragraph_format.tab_stops.add_tab_stop(Cm(8.4), WD_TAB_ALIGNMENT.RIGHT)
        small.append(paragraph.add_run(text))
    for run in small:
        run.font.size = Pt(10)

    qr = doc.add_paragraph()
    qr.paragraph_format.space_before = Pt(4)
    qr.add_run("{qr}  {inat_qr}")
    return doc


def _new_run_like(run):
    r = OxmlElement("w:r")
    if run._r.rPr is not None:
        r.append(deepcopy(run._r.rPr))
    return r


def fill_paragraph(paragraph, values):
    """Replace placeholders in place, keeping the formatting of the run each one starts in.

    Word often splits "{species}" over several runs, so matching works on the joined text.
    Matches are handled from the end, so inserted pictures never shift the ones still to do.
    Returns True when the paragraph held a placeholder.
    """
    text = "".join(r.text for r in paragraph.runs)
    matches = [m for m in PLACEHOLDER.finditer(text) if m.group(1).strip() in values]
    for match in reversed(matches):
        runs = paragraph.runs
        start, end = match.span()
        offsets, position = [], 0
        for run in runs:
            offsets.append(position)
            position += len(run.text)
        first = next(i for i, run in enumerate(runs) if offsets[i] <= start < offsets[i] + len(run.text))
        last = next(i for i, run in enumerate(runs) if offsets[i] < end <= offsets[i] + len(run.text))
        head = runs[first].text[:start - offsets[first]]
        tail = runs[last].text[end - offsets[last]:]
        for run in runs[first + 1:last]:
            run.text = ""

        key = match.group(1).strip()
        value = values[key]
        if key not in IMAGE_KEYS:
            if first == last:
                runs[first].text = head + value + tail
            else:
                runs[first].text = head + value
                runs[last].text = tail
            continue

        runs[first].text = head
        if first != last:
            runs[last].text = tail
        anchor = runs[first]._r
        if value is not None:
            picture = _new_run_like(runs[first])
            anchor.addnext(picture)
            Run(picture, paragraph).add_picture(value, width=QR_WIDTH)
            value.seek(0)
            anchor = picture
        if first == last and tail:
            rest = _new_run_like(runs[first])
            anchor.addnext(rest)
            Run(rest, paragraph).text = tail
    return bool(matches)


def is_blank(paragraph):
    has_picture = paragraph._p.xpath(".//w:drawing")
    return not paragraph.text.strip() and not has_picture


def fill_block(element, values, parent):
    """Fill every paragraph inside a copied piece of the template (including nested tables)."""
    for p in list(element.iter(qn("w:p"))):
        paragraph = Paragraph(p, parent)
        if fill_paragraph(paragraph, values) and is_blank(paragraph):
            # A line made only of empty placeholders (e.g. no Russian name) disappears
            container = p.getparent()
            if container.tag == qn("w:tc") and len(container.findall(qn("w:p"))) == 1:
                continue  # a table cell must keep one paragraph
            container.remove(p)


def _set_cut_lines(table, width):
    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    tbl_w.set(qn("w:type"), "dxa")
    tbl_w.set(qn("w:w"), str(int(width.twips if hasattr(width, "twips") else width / 635)))
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{side}")
        border.set(qn("w:val"), "dashed")
        border.set(qn("w:sz"), "4")
        border.set(qn("w:color"), "A0A0A0")
        borders.append(border)
    # Word is strict about the order of table properties: borders go right after the width
    tbl_w.addnext(borders)


def _keep_row_on_one_page(row):
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))


def render(template, rows_values, per_row=2):
    """Build the output document: a grid of filled-in copies of the template body."""
    doc = template
    body = doc.element.body
    label = [deepcopy(el) for el in body if el.tag != qn("w:sectPr")]
    for el in list(body):
        if el.tag != qn("w:sectPr"):
            body.remove(el)

    section = doc.sections[-1]
    width = section.page_width - section.left_margin - section.right_margin
    rows = (len(rows_values) + per_row - 1) // per_row
    table = doc.add_table(rows=rows, cols=per_row)
    table.autofit = False
    _set_cut_lines(table, width)
    for column in table.columns:
        column.width = int(width / per_row)
    for row in table.rows:
        _keep_row_on_one_page(row)
        for cell in row.cells:
            cell.width = int(width / per_row)

    for index, values in enumerate(rows_values):
        cell = table.rows[index // per_row].cells[index % per_row]
        tc = cell._tc
        for p in tc.findall(qn("w:p")):
            tc.remove(p)
        for el in label:
            tc.append(deepcopy(el))
        fill_block(tc, values, cell)
        if tc[-1].tag != qn("w:p"):
            tc.append(OxmlElement("w:p"))  # Word requires a cell to end with a paragraph

    # Copied pictures share drawing ids; Word may call the file damaged unless they are unique
    for number, doc_pr in enumerate(body.iter(qn("wp:docPr")), start=1):
        doc_pr.set("id", str(number))
    return doc


def placeholders_in(template):
    text = "\n".join(Paragraph(p, None).text for p in template.element.body.iter(qn("w:p")))
    return {m.group(1).strip() for m in PLACEHOLDER.finditer(text)}
