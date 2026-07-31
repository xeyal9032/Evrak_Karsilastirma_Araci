# -*- coding: utf-8 -*-
"""HTML and PDF summary reports (single-file HTML with client-side filters)."""
import html
import json
import os
from datetime import datetime


def _esc(v):
    return html.escape("" if v is None else str(v))


def _entry_cells(entry, source, note, status):
    if entry is None:
        return {
            "source": source, "line": "", "umsatz": "", "sh": "", "konto": "",
            "gegenkonto": "", "datum": "", "belegfeld1": "", "text": "",
            "note": note or "", "status": status,
        }
    return {
        "source": source,
        "line": entry.get("line", ""),
        "umsatz": entry.get("umsatz", ""),
        "sh": entry.get("sh", ""),
        "konto": entry.get("konto", ""),
        "gegenkonto": entry.get("gegenkonto", ""),
        "datum": entry.get("datum", ""),
        "belegfeld1": entry.get("belegfeld1", ""),
        "text": entry.get("text", ""),
        "note": note or "",
        "status": status,
    }


def groups_to_rows(groups, f1_label, f2_label):
    """Flatten compare groups into display rows for HTML/PDF."""
    rows = []
    for g in groups:
        kind = g["kind"]
        if kind == "MATCH":
            rows.append(_entry_cells(g["e1"], f1_label, "", "MATCH"))
            rows.append(_entry_cells(g["e2"], f2_label, "", "MATCH"))
        elif kind == "MISMATCH":
            note = ""
            e1, e2 = g["e1"], g["e2"]
            if e1 and e2:
                from karsilastir_motor import diff_note
                note = diff_note(e1, e2)
            rows.append(_entry_cells(e1, f1_label, "", "MISMATCH"))
            rows.append(_entry_cells(e2, f2_label, note, "MISMATCH"))
        elif kind == "ONLY1":
            rows.append(_entry_cells(g["e1"], f1_label, "", "ONLY1"))
            rows.append(_entry_cells(None, f2_label, "MISSING", "ONLY1"))
        elif kind == "ONLY2":
            rows.append(_entry_cells(None, f1_label, "MISSING", "ONLY2"))
            rows.append(_entry_cells(g["e2"], f2_label, "", "ONLY2"))
    return rows


def write_html_report(out_path, *, f1_path, f2_path, f1_label, f2_label,
                      match_count, mismatch_count, only1_count, only2_count,
                      groups, diffs=None, title="Evrak Karsilastirma"):
    rows = groups_to_rows(groups, f1_label, f2_label)
    payload = json.dumps(rows, ensure_ascii=False)
    generated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    top_diffs = diffs[:10] if diffs else []

    diff_html = ""
    for item in top_diffs:
        k, b1, b2, d, c1, c2 = item
        diff_html += (
            f"<tr><td>{_esc(k)}</td><td>{b1:,.2f}</td><td>{b2:,.2f}</td>"
            f"<td>{d:,.2f}</td><td>{c1}</td><td>{c2}</td></tr>"
        )

    doc = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>{_esc(title)}</title>
<style>
:root {{ --bg:#f6f7f9; --card:#fff; --ink:#1a1d23; --muted:#5c6570; --line:#d8dde3;
  --match:#fff3a0; --mis:#ffd08a; --only:#ffb3b3; }}
* {{ box-sizing:border-box; }}
body {{ margin:0; font:14px/1.45 "Segoe UI", system-ui, sans-serif; color:var(--ink); background:var(--bg); }}
header {{ padding:28px 24px 12px; }}
h1 {{ margin:0 0 6px; font-size:1.55rem; }}
.meta {{ color:var(--muted); }}
.wrap {{ padding:0 24px 40px; max-width:1200px; margin:0 auto; }}
.cards {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(140px,1fr)); gap:12px; margin:16px 0 20px; }}
.card {{ background:var(--card); border:1px solid var(--line); border-radius:10px; padding:14px; }}
.card b {{ display:block; font-size:1.4rem; margin-top:4px; }}
.filters {{ display:flex; flex-wrap:wrap; gap:8px; margin:12px 0 16px; align-items:center; }}
.filters button {{ border:1px solid var(--line); background:#fff; border-radius:8px; padding:8px 12px; cursor:pointer; }}
.filters button.active {{ background:#1a1d23; color:#fff; border-color:#1a1d23; }}
.filters input {{ flex:1; min-width:180px; padding:8px 10px; border:1px solid var(--line); border-radius:8px; }}
table {{ width:100%; border-collapse:collapse; background:var(--card); border:1px solid var(--line); }}
th, td {{ padding:8px 10px; border-bottom:1px solid var(--line); text-align:left; vertical-align:top; }}
th {{ position:sticky; top:0; background:#eef1f5; font-size:12px; }}
tr.MATCH td {{ background:var(--match); }}
tr.MISMATCH td {{ background:var(--mis); }}
tr.ONLY1 td, tr.ONLY2 td {{ background:var(--only); }}
.small {{ overflow:auto; max-height:70vh; border-radius:10px; border:1px solid var(--line); }}
</style>
</head>
<body>
<header class="wrap">
  <h1>{_esc(title)}</h1>
  <div class="meta">Generated { _esc(generated) }</div>
  <div class="meta">File 1: {_esc(f1_path)}</div>
  <div class="meta">File 2: {_esc(f2_path)}</div>
</header>
<main class="wrap">
  <div class="cards">
    <div class="card">Match (yellow)<b id="c-match">{match_count}</b></div>
    <div class="card">Mismatch (orange)<b id="c-mis">{mismatch_count}</b></div>
    <div class="card">Only file1 (red)<b id="c-o1">{only1_count}</b></div>
    <div class="card">Only file2 (red)<b id="c-o2">{only2_count}</b></div>
  </div>
  <h2>Account differences</h2>
  <div class="small" style="max-height:240px">
  <table>
    <thead><tr><th>Konto</th><th>Bal 1</th><th>Bal 2</th><th>Diff</th><th>N1</th><th>N2</th></tr></thead>
    <tbody>{diff_html or "<tr><td colspan='6'>No significant diffs</td></tr>"}</tbody>
  </table>
  </div>
  <h2>Rows</h2>
  <div class="filters">
    <button type="button" class="active" data-filter="ALL">All</button>
    <button type="button" data-filter="MATCH">Match</button>
    <button type="button" data-filter="MISMATCH">Orange</button>
    <button type="button" data-filter="ONLY1">Only 1</button>
    <button type="button" data-filter="ONLY2">Only 2</button>
    <button type="button" data-filter="RED">Red (1+2)</button>
    <input id="q" type="search" placeholder="Filter text / Konto / Belegfeld1..."/>
  </div>
  <div class="small">
  <table>
    <thead>
      <tr>
        <th>Source</th><th>Line</th><th>Amount</th><th>S/H</th><th>Konto</th>
        <th>Gegen</th><th>Date</th><th>Belegfeld1</th><th>Text</th><th>Note</th><th>Status</th>
      </tr>
    </thead>
    <tbody id="tbody"></tbody>
  </table>
  </div>
</main>
<script>
const ROWS = {payload};
let filter = "ALL";
const tbody = document.getElementById("tbody");
const q = document.getElementById("q");
function visible(row) {{
  const st = row.status;
  if (filter === "RED") {{ if (st !== "ONLY1" && st !== "ONLY2") return false; }}
  else if (filter !== "ALL" && st !== filter) return false;
  const needle = (q.value || "").trim().toLowerCase();
  if (!needle) return true;
  return Object.values(row).join(" ").toLowerCase().includes(needle);
}}
function render() {{
  const frag = document.createDocumentFragment();
  for (const row of ROWS) {{
    if (!visible(row)) continue;
    const tr = document.createElement("tr");
    tr.className = row.status;
    for (const k of ["source","line","umsatz","sh","konto","gegenkonto","datum","belegfeld1","text","note","status"]) {{
      const td = document.createElement("td");
      td.textContent = row[k] == null ? "" : String(row[k]);
      tr.appendChild(td);
    }}
    frag.appendChild(tr);
  }}
  tbody.innerHTML = "";
  tbody.appendChild(frag);
}}
document.querySelectorAll(".filters button[data-filter]").forEach(btn => {{
  btn.addEventListener("click", () => {{
    document.querySelectorAll(".filters button").forEach(b => b.classList.remove("active"));
    btn.classList.add("active");
    filter = btn.dataset.filter;
    render();
  }});
}});
q.addEventListener("input", render);
render();
</script>
</body>
</html>
"""
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(doc)
    return out_path


def write_pdf_report(out_path, *, f1_path, f2_path, f1_label, f2_label,
                     match_count, mismatch_count, only1_count, only2_count,
                     groups, diffs=None, title="Evrak Karsilastirma"):
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import cm
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    except ImportError as ex:
        raise RuntimeError(
            "PDF export requires reportlab. Install with: pip install reportlab"
        ) from ex

    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    doc = SimpleDocTemplate(out_path, pagesize=A4, leftMargin=1.5 * cm, rightMargin=1.5 * cm)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Small", parent=styles["Normal"], fontSize=9, leading=12))
    story = []
    story.append(Paragraph(html.escape(title), styles["Title"]))
    story.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles["Small"]))
    story.append(Paragraph(f"File 1 ({html.escape(f1_label)}): {html.escape(f1_path)}", styles["Small"]))
    story.append(Paragraph(f"File 2 ({html.escape(f2_label)}): {html.escape(f2_path)}", styles["Small"]))
    story.append(Spacer(1, 0.4 * cm))

    summary = [
        ["Status", "Count"],
        ["MATCH (yellow)", str(match_count)],
        ["MISMATCH (orange)", str(mismatch_count)],
        ["ONLY1 (red)", str(only1_count)],
        ["ONLY2 (red)", str(only2_count)],
    ]
    t = Table(summary, colWidths=[10 * cm, 4 * cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#FFFF99")),
        ("BACKGROUND", (0, 2), (-1, 2), colors.HexColor("#FFC000")),
        ("BACKGROUND", (0, 3), (-1, 4), colors.HexColor("#FF6B6B")),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.5 * cm))
    story.append(Paragraph("Top account differences", styles["Heading2"]))

    diff_rows = [["Konto", "Bal1", "Bal2", "Diff"]]
    for item in (diffs or [])[:15]:
        k, b1, b2, d, c1, c2 = item
        diff_rows.append([str(k), f"{b1:,.2f}", f"{b2:,.2f}", f"{d:,.2f}"])
    if len(diff_rows) == 1:
        diff_rows.append(["—", "—", "—", "0"])
    td = Table(diff_rows, colWidths=[3 * cm, 3.5 * cm, 3.5 * cm, 3.5 * cm])
    td.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
    ]))
    story.append(td)
    story.append(Spacer(1, 0.5 * cm))
    story.append(Paragraph("Critical rows (red / orange sample)", styles["Heading2"]))

    sample = [["Status", "Source", "Konto", "Amount", "Belegfeld1", "Text"]]
    for row in groups_to_rows(groups, f1_label, f2_label):
        if row["status"] not in ("ONLY1", "ONLY2", "MISMATCH"):
            continue
        sample.append([
            row["status"],
            str(row["source"])[:18],
            str(row["konto"])[:10],
            str(row["umsatz"])[:12],
            str(row["belegfeld1"])[:18],
            str(row["text"])[:40],
        ])
        if len(sample) > 40:
            break
    if len(sample) == 1:
        sample.append(["—", "—", "—", "—", "—", "No critical rows"])
    ts = Table(sample, colWidths=[2.2 * cm, 2.5 * cm, 2 * cm, 2.2 * cm, 3 * cm, 5 * cm])
    ts.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
    ]))
    story.append(ts)
    doc.build(story)
    return out_path
