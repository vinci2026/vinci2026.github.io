"""Build the static schedule page from the adjacent VINCI2026_Schedule.xlsx file."""

from html import escape
from pathlib import Path
import re
import xml.etree.ElementTree as ET
from zipfile import ZipFile


ROOT = Path(__file__).resolve().parent
WORKBOOK = ROOT.parent / "VINCI2026_Schedule.xlsx"
NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def sheets():
    with ZipFile(WORKBOOK) as archive:
        strings = ["".join(node.itertext()) for node in ET.fromstring(
            archive.read("xl/sharedStrings.xml")
        ).findall("x:si", NS)]
        result = []
        for number in (1, 2):
            root = ET.fromstring(archive.read(f"xl/worksheets/sheet{number}.xml"))
            rows = []
            for row in root.findall(".//x:sheetData/x:row", NS):
                cells = {}
                for cell in row.findall("x:c", NS):
                    value = cell.find("x:v", NS)
                    if value is None:
                        continue
                    key = re.match(r"[A-Z]+", cell.attrib["r"]).group()
                    cells[key] = strings[int(value.text)] if cell.attrib.get("t") == "s" else value.text
                rows.append(cells)
            result.append(rows)
        return result


def time(value):
    minutes = round(float(value) * 24 * 60)
    return f"{minutes // 60:02d}:{minutes % 60:02d}"


VENUES = {
    "城北文澜大酒店大厅": "Chengbei Wenlan Hotel lobby",
    "文锦厅": "Wenjin Hall",
    "文锦厅 A": "Wenjin Hall A",
    "文锦厅 B": "Wenjin Hall B",
    "文锦厅外廊": "Outside Wenjin Hall",
    "各会议厅外廊": "Outside the meeting halls",
    "酒店餐厅": "Hotel restaurant",
    "酒店大门": "Hotel entrance",
    "天目里": "OōEli Tianmuli",
    "文澜": "Wenlan",
}

CHAIRS = {
    "马昱欣，\n南方科技大学": "Yuxin Ma, Southern University of Science and Technology",
    "曾伟，\n香港科技大学（广州）": "Wei Zeng, HKUST (Guangzhou)",
    "许丕文，\n香港科技大学（广州）": "James She, HKUST (Guangzhou)",
    "Zi-Wi WU, \nHong Kong University of Science and Technology\n ": "Zi-Wei Wu, Hong Kong University of Science and Technology",
}

ACTIVITIES = {
    "会场--天目里": "Transfer to OōEli Tianmuli",
    "Conference Dinner+表演": "Conference Dinner and Performance",
    "颁奖（5*3）-- Track Chair or ？ --文案联系James": "Awards Ceremony",
    "明年主办介绍--Tony Huang": "Next VINCI Host Introduction — Tony Huang",
    "--闭幕致辞-- James": "Closing Remarks — James She",
}


def activity(raw):
    raw = raw.strip()
    if raw in ACTIVITIES:
        return ACTIVITIES[raw]
    if raw.startswith("Opening Keynote"):
        return "Opening Keynote: Beyond Generation: Supporting Human Creativity and Learning with AI — Hongbo Fu"
    if raw.startswith("Keynote 2"):
        return "Keynote 2: Living Archives: From Digital Preservation to Artistic Reinterpretation — Juergen Hagler"
    if raw.startswith("Keynote 3"):
        return "Keynote 3: Explorations on Novel Natural Human-Computer Interaction Techniques — Feng Tian"
    if raw.startswith("<Leonardo @ VINCI Panel"):
        return "Leonardo @ VINCI Panel Discussion"
    return " ".join(raw.split())


def chair(raw):
    if not raw:
        return ""
    return CHAIRS.get(raw, " ".join(raw.split()))


def overview(rows):
    days = [
        ("2026.11.1", "Sunday, 1 November", "nov-1"),
        ("2026.11.2", "Monday, 2 November", "nov-2"),
        ("2026.11.3", "Tuesday, 3 November", "nov-3"),
        ("2026.11.4", "Wednesday, 4 November", "nov-4"),
    ]
    result = []
    for date, label, anchor in days:
        result.append(f'<section class="schedule-day" id="{anchor}"><h2>{label}</h2><div class="schedule-grid">')
        current_start = current_end = None
        started = False
        for row in rows[2:]:
            if row.get("A", "").startswith(date):
                started = True
            elif started and row.get("A", "").startswith("2026."):
                break
            if not started or not row.get("D"):
                continue
            current_start = row.get("B", current_start)
            current_end = row.get("C", current_end)
            if not current_start or not current_end:
                continue
            title = activity(row["D"])
            venue = VENUES.get(row.get("G", ""), row.get("G", ""))
            if date == "2026.11.2":
                if title == "Coffee Break":
                    venue = "Wenjin Hall"
                elif title in ("Lunch", "Reception"):
                    venue = "1F Puyue All-Day Restaurant"
                    if title == "Reception":
                        title = "Dinner"
            speaker = chair(row.get("F", ""))
            extra = f'<span class="schedule-chair">Chair: {escape(speaker)}</span>' if speaker else ""
            if title == "Conference Opening":
                extra += ('<span class="schedule-chair">Welcome remarks: Yun Chen (Hangzhou City University). '
                          'VINCI overview: Kang Zhang, The Hong Kong University of Science and Technology (Guangzhou). '
                          'Paper-track overview: Yuxin Ma, James She, Xiaojuan Chen and Wei Zeng.</span>')
            result.append(
                '<div class="schedule-entry">'
                f'<time>{time(current_start)}–{time(current_end)}</time>'
                f'<div><strong>{escape(title)}</strong>{extra}</div>'
                f'<span class="schedule-venue">{escape(venue)}</span></div>'
            )
        result.append('</div></section>')
    return "\n".join(result)


def details(rows):
    parts = []
    open_day = open_session = False
    i = 4
    while i < len(rows):
        value = rows[i].get("A", "").strip()
        if not value:
            i += 1
            continue
        if re.match(r"^(Monday|Tuesday|Wednesday) \d", value):
            if open_session:
                parts.append('</ul></section>')
                open_session = False
            if open_day:
                parts.append('</div></section>')
            parts.append(f'<section class="papers-day"><h3>{escape(value)}</h3><div class="papers-grid">')
            open_day = True
            i += 1
            continue
        if re.match(r"^\d{2}:\d{2}–", value):
            if open_session:
                parts.append('</ul></section>')
            parts.append('<section class="paper-session">')
            parts.append(f'<p class="paper-session-meta">{escape(value.replace(chr(160), " "))}</p>')
            open_session = True
            i += 1
            parts.append(f'<h4>{escape(rows[i].get("A", ""))}</h4><ul>')
            i += 1
            continue
        if open_session and i + 1 < len(rows):
            match = re.search(r"\s*\((full|short|poster)\)$", value)
            kind = match.group(1) if match else ""
            title = value[:match.start()].strip() if match else value
            authors = rows[i + 1].get("A", "")
            parts.append(f'<li><strong>{escape(title)}</strong> <span class="paper-type">{escape(kind)}</span>'
                         f'<span class="paper-authors">{escape(authors)}</span></li>')
            i += 2
            continue
        i += 1
    if open_session:
        parts.append('</ul></section>')
    if open_day:
        parts.append('</div></section>')
    return "\n".join(parts)


def main():
    overview_rows, detail_rows = sheets()
    template = (ROOT / "keynotes.html").read_text(encoding="utf-8")
    head = template[:template.index("    <main")]
    foot = template[template.index("    <footer"):]
    head = head.replace("Keynotes - VINCI 2026", "Schedule - VINCI 2026")
    head = head.replace("Keynote speakers and talks at VINCI 2026", "Conference schedule and paper sessions for VINCI 2026")
    head = head[:head.index("<style>")] + """<style>
.schedule-content { max-width:1100px; padding:2rem 1rem 3rem; }
.schedule-content h1 { font-size:2rem; font-weight:700; border-bottom:2px solid #00008b; padding-bottom:1rem; }
.schedule-intro { margin:1rem 0 1.5rem; color:#4b5563; }
.schedule-days { display:flex; flex-wrap:wrap; gap:.5rem; margin-bottom:2rem; }
.schedule-days a { border:1px solid #beced0; padding:.45rem .8rem; color:#1e555a; text-decoration:none; }
.schedule-days a:hover, .schedule-days a:focus { background:#e8f1f2; }
.schedule-day { margin:2rem 0 3rem; scroll-margin-top:1rem; }
.schedule-day h2, .schedule-content > h2 { font-size:1.5rem; color:#1e555a; margin-bottom:1rem; }
.schedule-grid { border-top:2px solid #1e555a; }
.schedule-entry { display:grid; grid-template-columns:140px minmax(0,1fr) 180px; gap:1rem; align-items:start; border-bottom:1px solid #dce5e6; padding:.85rem .5rem; }
.schedule-entry:nth-child(even) { background:#f7faf9; }
.schedule-entry time { font-weight:700; white-space:nowrap; }
.schedule-entry strong { display:block; line-height:1.45; }
.schedule-chair, .schedule-venue { display:block; color:#526267; font-size:.92rem; margin-top:.25rem; }
.schedule-venue { margin-top:0; }
.papers-day { margin-top:2rem; }
.papers-day h3 { font-size:1.25rem; color:#1e555a; border-bottom:1px solid #bbcdd0; padding-bottom:.5rem; }
.papers-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:1.5rem; }
.paper-session { padding:.5rem 0 1rem; min-width:0; }
.paper-session-meta { color:#1e555a; font-weight:700; margin-bottom:.3rem; }
.paper-session h4 { font-size:1.1rem; font-weight:700; margin-bottom:.8rem; }
.paper-session ul { padding-left:1.2rem; }
.paper-session li { padding:.35rem 0; overflow-wrap:anywhere; }
.paper-authors { display:block; color:#596368; font-size:.9rem; }
.paper-type { color:#1e555a; font-size:.8rem; text-transform:uppercase; white-space:nowrap; }
@media(max-width:767.98px) { .schedule-entry { grid-template-columns:100px minmax(0,1fr); gap:.4rem .8rem; } .schedule-venue { grid-column:2; } .papers-grid { grid-template-columns:1fr; gap:0; } }
@media(max-width:420px) { .schedule-entry { grid-template-columns:1fr; } .schedule-venue { grid-column:1; } }
</style>
</head>

<body>
""" + head[head.index('    <section class="container-fluid"'):]
    head = head.replace('custom-nav-link active-page" href="keynotes.html"', 'custom-nav-link" href="keynotes.html"')
    head = head.replace('dropdown-item active-page" href="keynotes.html"', 'dropdown-item" href="keynotes.html"')
    head = head.replace('custom-nav-link" href="schedule.html"', 'custom-nav-link active-page" href="schedule.html"')
    head = head.replace('dropdown-item" href="schedule.html"', 'dropdown-item active-page" href="schedule.html"')
    body = f'''    <main class="container schedule-content">
        <h1>Schedule</h1>
        <p class="schedule-intro">VINCI 2026 · Hangzhou, China · 1–4 November 2026. This program is subject to change.</p>
        <nav class="schedule-days" aria-label="Schedule dates">
            <a href="#nov-1">1 Nov · Sunday</a><a href="#nov-2">2 Nov · Monday</a>
            <a href="#nov-3">3 Nov · Tuesday</a><a href="#nov-4">4 Nov · Wednesday</a>
            <a href="#paper-sessions">Paper sessions</a>
        </nav>
{overview(overview_rows)}
        <h2 id="paper-sessions">Paper Sessions</h2>
        <p>Full papers: 13 minutes. Short papers and posters: 8 minutes. All durations include Q&amp;A.</p>
{details(detail_rows)}
    </main>

'''
    (ROOT / "schedule.html").write_text(head + body + foot, encoding="utf-8")


if __name__ == "__main__":
    main()
