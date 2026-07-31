# -*- coding: utf-8 -*-
"""HTML and PDF summary reports (single-file HTML with client-side filters)."""
import html
import json
import os
from datetime import datetime

import i18n


def _esc(v):
    return html.escape("" if v is None else str(v))


def _t(key, lang, **kwargs):
    return i18n.t(key, lang=lang, **kwargs)


def _norm_lang(lang):
    lang = (lang or i18n.detect_system_lang() or "tr").lower()
    return lang if lang in i18n.SUPPORTED else "tr"


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


def groups_to_rows(groups, f1_label, f2_label, lang=None):
    """Flatten compare groups into display rows for HTML/PDF."""
    missing = _t("html_missing", _norm_lang(lang))
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
            rows.append(_entry_cells(None, f2_label, missing, "ONLY1"))
        elif kind == "ONLY2":
            rows.append(_entry_cells(None, f1_label, missing, "ONLY2"))
            rows.append(_entry_cells(g["e2"], f2_label, "", "ONLY2"))
    return rows


def write_html_report(out_path, *, f1_path, f2_path, f1_label, f2_label,
                      match_count, mismatch_count, only1_count, only2_count,
                      groups, diffs=None, title=None, lang=None):
    lang = _norm_lang(lang)
    title = title or _t("html_title", lang)
    rows = groups_to_rows(groups, f1_label, f2_label, lang=lang)
    # Escape < so a Buchungstext containing </script> cannot break out of the
    # inline JSON script block (JSON itself allows raw <).
    payload = json.dumps(rows, ensure_ascii=False).replace("<", "\\u003c").replace(">", "\\u003e")
    generated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    top_diffs = diffs[:10] if diffs else []

    diff_html = ""
    for item in top_diffs:
        k, b1, b2, d, c1, c2 = item
        diff_html += (
            f"<tr><td>{_esc(k)}</td><td>{b1:,.2f}</td><td>{b2:,.2f}</td>"
            f"<td>{d:,.2f}</td><td>{c1}</td><td>{c2}</td></tr>"
        )
    if not diff_html:
        diff_html = f"<tr><td colspan='6'>{_esc(_t('html_no_diffs', lang))}</td></tr>"

    doc = f"""<!DOCTYPE html>
<html lang="{_esc(lang)}">
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
  <div class="meta">{_esc(_t("html_generated", lang, ts=generated))}</div>
  <div class="meta">{_esc(_t("html_file1", lang, path=f1_path))}</div>
  <div class="meta">{_esc(_t("html_file2", lang, path=f2_path))}</div>
</header>
<main class="wrap">
  <div class="cards">
    <div class="card">{_esc(_t("html_card_match", lang))}<b id="c-match">{match_count}</b></div>
    <div class="card">{_esc(_t("html_card_mismatch", lang))}<b id="c-mis">{mismatch_count}</b></div>
    <div class="card">{_esc(_t("html_card_only1", lang))}<b id="c-o1">{only1_count}</b></div>
    <div class="card">{_esc(_t("html_card_only2", lang))}<b id="c-o2">{only2_count}</b></div>
  </div>
  <h2>{_esc(_t("html_account_diffs", lang))}</h2>
  <div class="small" style="max-height:240px">
  <table>
    <thead><tr>
      <th>{_esc(_t("html_th_konto", lang))}</th>
      <th>{_esc(_t("html_th_bal1", lang))}</th>
      <th>{_esc(_t("html_th_bal2", lang))}</th>
      <th>{_esc(_t("html_th_diff", lang))}</th>
      <th>{_esc(_t("html_th_n1", lang))}</th>
      <th>{_esc(_t("html_th_n2", lang))}</th>
    </tr></thead>
    <tbody>{diff_html}</tbody>
  </table>
  </div>
  <h2>{_esc(_t("html_rows", lang))}</h2>
  <div class="filters">
    <button type="button" class="active" data-filter="ALL">{_esc(_t("html_filter_all", lang))}</button>
    <button type="button" data-filter="MATCH">{_esc(_t("html_filter_match", lang))}</button>
    <button type="button" data-filter="MISMATCH">{_esc(_t("html_filter_mismatch", lang))}</button>
    <button type="button" data-filter="ONLY1">{_esc(_t("html_filter_only1", lang))}</button>
    <button type="button" data-filter="ONLY2">{_esc(_t("html_filter_only2", lang))}</button>
    <button type="button" data-filter="RED">{_esc(_t("html_filter_red", lang))}</button>
    <input id="q" type="search" placeholder="{_esc(_t("html_filter_placeholder", lang))}"/>
  </div>
  <div class="small">
  <table>
    <thead>
      <tr>
        <th>{_esc(_t("html_col_source", lang))}</th>
        <th>{_esc(_t("html_col_line", lang))}</th>
        <th>{_esc(_t("html_col_amount", lang))}</th>
        <th>{_esc(_t("html_col_sh", lang))}</th>
        <th>{_esc(_t("html_col_konto", lang))}</th>
        <th>{_esc(_t("html_col_gegen", lang))}</th>
        <th>{_esc(_t("html_col_date", lang))}</th>
        <th>{_esc(_t("html_col_beleg", lang))}</th>
        <th>{_esc(_t("html_col_text", lang))}</th>
        <th>{_esc(_t("html_col_note", lang))}</th>
        <th>{_esc(_t("html_col_status", lang))}</th>
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


def _register_pdf_fonts():
    """Register a Unicode TTF so TR/RU/DE glyphs render (Helvetica cannot)."""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    regular_name = "EvrakSans"
    bold_name = "EvrakSans-Bold"
    if regular_name in pdfmetrics.getRegisteredFontNames():
        return regular_name, bold_name

    windir = os.environ.get("WINDIR", r"C:\Windows")
    regular_candidates = [
        os.path.join(windir, "Fonts", "arialuni.ttf"),
        os.path.join(windir, "Fonts", "ARIALUNI.TTF"),
        os.path.join(windir, "Fonts", "segoeui.ttf"),
        os.path.join(windir, "Fonts", "arial.ttf"),
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
    ]
    bold_candidates = [
        os.path.join(windir, "Fonts", "segoeuib.ttf"),
        os.path.join(windir, "Fonts", "arialbd.ttf"),
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf",
    ]

    regular = next((p for p in regular_candidates if os.path.isfile(p)), None)
    if not regular:
        return "Helvetica", "Helvetica-Bold"

    pdfmetrics.registerFont(TTFont(regular_name, regular))
    bold = next((p for p in bold_candidates if os.path.isfile(p)), None)
    if bold:
        pdfmetrics.registerFont(TTFont(bold_name, bold))
    else:
        bold_name = regular_name
    return regular_name, bold_name


def _pdf_cell(text, style):
    from reportlab.platypus import Paragraph
    return Paragraph(_esc(text), style)


def write_pdf_report(out_path, *, f1_path, f2_path, f1_label, f2_label,
                     match_count, mismatch_count, only1_count, only2_count,
                     groups, diffs=None, title=None, lang=None):
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.lib.units import cm
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    except ImportError as ex:
        raise RuntimeError(
            "PDF export requires reportlab. Install with: pip install reportlab"
        ) from ex

    lang = _norm_lang(lang)
    title = title or _t("html_title", lang)
    generated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    font, font_bold = _register_pdf_fonts()

    style_title = ParagraphStyle(
        "EvrakTitle", fontName=font_bold, fontSize=16, leading=20, spaceAfter=8,
    )
    style_meta = ParagraphStyle(
        "EvrakMeta", fontName=font, fontSize=9, leading=12, textColor=colors.HexColor("#5c6570"),
    )
    style_h2 = ParagraphStyle(
        "EvrakH2", fontName=font_bold, fontSize=12, leading=15, spaceBefore=10, spaceAfter=6,
    )
    style_cell = ParagraphStyle(
        "EvrakCell", fontName=font, fontSize=7, leading=9,
    )
    style_cell_bold = ParagraphStyle(
        "EvrakCellBold", fontName=font_bold, fontSize=8, leading=10, textColor=colors.white,
    )

    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    doc = SimpleDocTemplate(
        out_path, pagesize=A4,
        leftMargin=1.4 * cm, rightMargin=1.4 * cm,
        topMargin=1.4 * cm, bottomMargin=1.4 * cm,
    )
    story = []
    story.append(Paragraph(_esc(title), style_title))
    story.append(Paragraph(_esc(_t("html_generated", lang, ts=generated)), style_meta))
    story.append(Paragraph(_esc(_t("html_file1", lang, path=f1_path)), style_meta))
    story.append(Paragraph(_esc(_t("html_file2", lang, path=f2_path)), style_meta))
    story.append(Spacer(1, 0.35 * cm))

    summary = [
        [_pdf_cell(_t("html_col_status", lang), style_cell_bold),
         _pdf_cell(_t("html_count", lang), style_cell_bold)],
        [_pdf_cell(_t("html_card_match", lang), style_cell),
         _pdf_cell(str(match_count), style_cell)],
        [_pdf_cell(_t("html_card_mismatch", lang), style_cell),
         _pdf_cell(str(mismatch_count), style_cell)],
        [_pdf_cell(_t("html_card_only1", lang), style_cell),
         _pdf_cell(str(only1_count), style_cell)],
        [_pdf_cell(_t("html_card_only2", lang), style_cell),
         _pdf_cell(str(only2_count), style_cell)],
    ]
    t = Table(summary, colWidths=[12 * cm, 3.5 * cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
        ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#FFF3A0")),
        ("BACKGROUND", (0, 2), (-1, 2), colors.HexColor("#FFD08A")),
        ("BACKGROUND", (0, 3), (-1, 4), colors.HexColor("#FFB3B3")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#d8dde3")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(t)

    story.append(Paragraph(_esc(_t("html_account_diffs", lang)), style_h2))
    diff_rows = [[
        _pdf_cell(_t("html_th_konto", lang), style_cell_bold),
        _pdf_cell(_t("html_th_bal1", lang), style_cell_bold),
        _pdf_cell(_t("html_th_bal2", lang), style_cell_bold),
        _pdf_cell(_t("html_th_diff", lang), style_cell_bold),
    ]]
    for item in (diffs or [])[:15]:
        k, b1, b2, d, c1, c2 = item
        diff_rows.append([
            _pdf_cell(str(k), style_cell),
            _pdf_cell(f"{b1:,.2f}", style_cell),
            _pdf_cell(f"{b2:,.2f}", style_cell),
            _pdf_cell(f"{d:,.2f}", style_cell),
        ])
    if len(diff_rows) == 1:
        diff_rows.append([
            _pdf_cell(_t("html_no_diffs", lang), style_cell),
            _pdf_cell("-", style_cell),
            _pdf_cell("-", style_cell),
            _pdf_cell("-", style_cell),
        ])
    td = Table(diff_rows, colWidths=[3.2 * cm, 4 * cm, 4 * cm, 4 * cm])
    td.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#d8dde3")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f6f7f9")]),
    ]))
    story.append(td)

    story.append(Paragraph(_esc(_t("pdf_critical_rows", lang)), style_h2))
    sample = [[
        _pdf_cell(_t("html_col_status", lang), style_cell_bold),
        _pdf_cell(_t("html_col_source", lang), style_cell_bold),
        _pdf_cell(_t("html_col_konto", lang), style_cell_bold),
        _pdf_cell(_t("html_col_amount", lang), style_cell_bold),
        _pdf_cell(_t("html_col_beleg", lang), style_cell_bold),
        _pdf_cell(_t("html_col_text", lang), style_cell_bold),
    ]]
    for row in groups_to_rows(groups, f1_label, f2_label, lang=lang):
        if row["status"] not in ("ONLY1", "ONLY2", "MISMATCH"):
            continue
        # Skip empty placeholder half-rows (no konto/amount/text)
        if not row["konto"] and not row["umsatz"] and not row["text"] and row["note"] in (
            _t("html_missing", lang), "MISSING", ""
        ):
            continue
        sample.append([
            _pdf_cell(row["status"], style_cell),
            _pdf_cell(str(row["source"])[:28], style_cell),
            _pdf_cell(str(row["konto"])[:12], style_cell),
            _pdf_cell(str(row["umsatz"])[:14], style_cell),
            _pdf_cell(str(row["belegfeld1"])[:22], style_cell),
            _pdf_cell(str(row["text"])[:70], style_cell),
        ])
        if len(sample) > 45:
            break
    if len(sample) == 1:
        sample.append([
            _pdf_cell("-", style_cell),
            _pdf_cell(_t("html_no_diffs", lang), style_cell),
            _pdf_cell("-", style_cell),
            _pdf_cell("-", style_cell),
            _pdf_cell("-", style_cell),
            _pdf_cell("-", style_cell),
        ])
    ts = Table(sample, colWidths=[2.0 * cm, 3.0 * cm, 1.8 * cm, 2.0 * cm, 2.8 * cm, 5.0 * cm])
    ts.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#d8dde3")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#fff5f5")]),
    ]))
    story.append(ts)
    doc.build(story)
    return out_path
