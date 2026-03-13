import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import numbers as xl_numbers
import math
import re
import io
import openai
import zipfile
import json
from lxml import etree
from fpdf import FPDF

# ──────────────────────────────────────────────
# 1. Page Setting
# ──────────────────────────────────────────────
st.set_page_config(page_title="IS 2010 Scoring System", layout="wide")
st.title("IS 2010 Excel Guided Lab Scoring System")

# ──────────────────────────────────────────────
# 2. OpenAI Client
# ──────────────────────────────────────────────
try:
    client = openai.OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except Exception:
    client = None
    st.error("⚠️ OpenAI API Key not found. Please check your .streamlit/secrets.toml")

# ──────────────────────────────────────────────
# 3. Constants
# ──────────────────────────────────────────────
STANDARD_COLORS = {
    "Dark Red":   "#C00000",
    "Red":        "#FF0000",
    "Orange":     "#FFC000",
    "Yellow":     "#FFFF00",
    "Light Green":"#92D050",
    "Green":      "#00B050",
    "Light Blue": "#00B0F0",
    "Blue":       "#0070C0",
    "Dark Blue":  "#002060",
    "Purple":     "#7030A0",
}
SCORING_MODES = ["value_only", "format_only", "both"]
SCORING_LABELS = {"value_only": "Value Only", "format_only": "Format Only", "both": "Both (Value + Format)"}

# ──────────────────────────────────────────────
# 4. Helper: Display formatting
# ──────────────────────────────────────────────
def format_ans(f, v):
    f_str = str(f).strip() if f is not None else ""
    v_str = str(v).strip() if v is not None else ""
    if f_str.startswith('='):
        return f"{f_str} ({v_str})"
    return v_str if v_str else "Empty"

# ──────────────────────────────────────────────
# 5. Core Grading: Value Comparison
# ──────────────────────────────────────────────
def check_value_equivalence(prof_f, stud_f, prof_v, stud_v):
    try:
        if prof_v is None and stud_v is None:
            return False
        v_p = float(prof_v) if prof_v is not None else 0.0
        v_s = float(stud_v) if stud_v is not None else 0.0
        values_match = math.isclose(v_p, v_s, rel_tol=1e-9, abs_tol=1e-9)
    except (ValueError, TypeError):
        values_match = str(prof_v).strip().upper() == str(stud_v).strip().upper()

    if values_match:
        p_is_formula = str(prof_f).startswith('=') if prof_f else False
        s_is_formula = str(stud_f).startswith('=') if stud_f else False
        if p_is_formula and s_is_formula:
            return str(prof_f).replace(" ", "").upper() == str(stud_f).replace(" ", "").upper()
        return True
    return False

# ──────────────────────────────────────────────
# 6. Core Grading: Format Comparison
# ──────────────────────────────────────────────
def check_format_equivalence(prof_cell, stud_cell):
    """
    Compares cell formatting: number format, font (bold/color/size), border, fill color.
    Returns (is_match: bool, issues: list[str])
    """
    issues = []

    # --- Number Format ---
    p_nf = prof_cell.number_format or "General"
    s_nf = stud_cell.number_format or "General"
    if p_nf.strip() != s_nf.strip():
        issues.append(f"Number format mismatch: yours='{s_nf}' expected='{p_nf}'")

    # --- Font ---
    p_font = prof_cell.font
    s_font = stud_cell.font
    if p_font and s_font:
        if bool(p_font.bold) != bool(s_font.bold):
            issues.append(f"Bold mismatch: yours={'Bold' if s_font.bold else 'Normal'} expected={'Bold' if p_font.bold else 'Normal'}")
        p_fcolor = str(p_font.color.rgb) if p_font.color and p_font.color.type == 'rgb' else "000000"
        s_fcolor = str(s_font.color.rgb) if s_font.color and s_font.color.type == 'rgb' else "000000"
        if p_fcolor[-6:] != s_fcolor[-6:]:
            issues.append(f"Font color mismatch: yours=#{s_fcolor[-6:]} expected=#{p_fcolor[-6:]}")
        p_fsize = p_font.size or 11
        s_fsize = s_font.size or 11
        if abs(float(p_fsize) - float(s_fsize)) > 0.5:
            issues.append(f"Font size mismatch: yours={s_fsize} expected={p_fsize}")

    # --- Border (checks all four sides) ---
    p_border = prof_cell.border
    s_border = stud_cell.border
    if p_border and s_border:
        for side in ['left', 'right', 'top', 'bottom']:
            p_bs = getattr(p_border, side)
            s_bs = getattr(s_border, side)
            p_style = p_bs.border_style if p_bs else None
            s_style = s_bs.border_style if s_bs else None
            if p_style != s_style:
                issues.append(f"Border {side} mismatch: yours='{s_style}' expected='{p_style}'")

    # Fill Color is intentionally excluded —
    # cell background color is used by the professor to mark graded cells,
    # not as a formatting requirement for students.

    return (len(issues) == 0), issues

# ──────────────────────────────────────────────
# 7. Sparkline Check (unchanged)
# ──────────────────────────────────────────────
def check_sparkline_advanced(p_cache, s_cache, cell_coord):
    def extract_xml_info(xml_cache, target_cell):
        try:
            clean_cell = target_cell.replace('$', '').upper()
            for xml_path, xml_content in xml_cache.items():
                tree = etree.fromstring(xml_content.encode())
                NS = {
                    "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
                    "xm":  "http://schemas.microsoft.com/office/excel/2006/main",
                }
                for sp in tree.findall('.//x14:sparkline', NS):
                    sq_elem = sp.find('xm:sqref', NS)
                    f_elem  = sp.find('xm:f',     NS)
                    if sq_elem is None or f_elem is None:
                        continue
                    norm_sq = sq_elem.text.replace('$', '').upper()
                    if norm_sq == clean_cell:
                        grp = sp.getparent().getparent()
                        return {
                            "range":       f_elem.text.strip(),
                            "type":        grp.get("type", "line"),
                            "markers":     grp.get("markers", "0") == "1",
                            "high_point":  grp.get("highPoint", "0") == "1",
                            "low_point":   grp.get("lowPoint", "0") == "1",
                            "first_point": grp.get("firstPoint", "0") == "1",
                            "last_point":  grp.get("lastPoint", "0") == "1",
                            "negative":    grp.get("negative", "0") == "1",
                        }
            return None
        except Exception as e:
            st.warning(f"[Sparkline Error] {e}")
            return None

    p_info = extract_xml_info(p_cache, cell_coord)
    s_info = extract_xml_info(s_cache, cell_coord)
    if not p_info: return "Skip", None, None, None
    if not s_info: return False, "Missing", "Sparkline Object", "Sparkline is missing."

    p_src = p_info['range'].replace("$", "").replace(" ", "").upper()
    s_src = s_info['range'].replace("$", "").replace(" ", "").upper()
    if p_src != s_src:
        return False, f"Range: {s_info['range']}", f"Range: {p_info['range']}", "Data range error"
    if p_info['type'] != s_info['type']:
        return False, f"Type: {s_info['type']}", f"Type: {p_info['type']}", "Sparkline type error"
    marker_keys = ['markers', 'high_point', 'low_point', 'first_point', 'last_point', 'negative']
    if any(p_info[k] != s_info[k] for k in marker_keys):
        return False, "Marker settings differ", "See professor's file", "Marker configuration error"
    return True, None, None, None

# ──────────────────────────────────────────────
# 8. Unified Grading Dispatcher
# ──────────────────────────────────────────────
def grade_cell(mode, p_wb_f, p_wb_v, s_wb_f, s_wb_v, p_cache, s_cache, sn, c):
    """
    Returns: (is_correct: bool, stud_ans_str, prof_ans_str, issues: list[str])
    """
    # Check sparkline first (always)
    is_sl, s_ans, p_ans, msg = check_sparkline_advanced(p_cache, s_cache, c)
    if is_sl != "Skip":
        return (bool(is_sl), f"[Sparkline] {s_ans}", f"[Sparkline] {p_ans}", [msg] if msg else [])

    pf = p_wb_f[sn][c].value
    sf = s_wb_f[sn][c].value
    pv = p_wb_v[sn][c].value
    sv = s_wb_v[sn][c].value

    p_cell_fmt = p_wb_f[sn][c]  # openpyxl cell object with formatting
    s_cell_fmt = s_wb_f[sn][c]

    if mode == "value_only":
        ok = check_value_equivalence(pf, sf, pv, sv)
        issues = [] if ok else [f"Value mismatch: yours='{format_ans(sf, sv)}' expected='{format_ans(pf, pv)}'"]
        return ok, format_ans(sf, sv), format_ans(pf, pv), issues

    elif mode == "format_only":
        ok, issues = check_format_equivalence(p_cell_fmt, s_cell_fmt)
        return ok, format_ans(sf, sv), format_ans(pf, pv), issues

    elif mode == "both":
        v_ok = check_value_equivalence(pf, sf, pv, sv)
        f_ok, f_issues = check_format_equivalence(p_cell_fmt, s_cell_fmt)
        issues = []
        if not v_ok:
            issues.append(f"Value mismatch: yours='{format_ans(sf, sv)}' expected='{format_ans(pf, pv)}'")
        issues.extend(f_issues)
        return (v_ok and f_ok), format_ans(sf, sv), format_ans(pf, pv), issues

    return False, "N/A", "N/A", ["Unknown mode"]

# ──────────────────────────────────────────────
# 9. AI Feedback
# ──────────────────────────────────────────────
def get_ai_feedback(prof_ans_str, stud_ans_str, issues=None, rubric=None, custom_msg=None):
    if not client: return "AI feedback disabled."
    if custom_msg: return custom_msg

    issue_text = "\n".join(issues) if issues else "General error"
    rubric_text = json.dumps(rubric, ensure_ascii=False, indent=2) if rubric else "No rubric available."

    prompt = f"""You are a professional Excel instructor giving diagnostic feedback.

Student's Answer: {stud_ans_str}
Professor's Correct Answer: {prof_ans_str}
Issues Detected: {issue_text}
Rubric Context: {rubric_text}

Write concise diagnostic feedback (under 150 words) explaining:
1. What went wrong
2. Which rubric criterion was missed (if applicable)
3. How to fix it"""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a professional Excel instructor."},
                {"role": "user",   "content": prompt}
            ]
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"AI Error: {e}"

# ──────────────────────────────────────────────
# 10. AI Rubric Auto-Generation
# ──────────────────────────────────────────────
def extract_sparkline_info_for_cell(p_cache, cell_coord):
    """
    Re-uses the XML parsing logic to extract sparkline metadata for a given cell.
    Returns a dict if sparkline found, None otherwise.
    """
    try:
        clean_cell = cell_coord.replace('$', '').upper()
        for xml_path, xml_content in p_cache.items():
            tree = etree.fromstring(xml_content.encode())
            NS = {
                "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
                "xm":  "http://schemas.microsoft.com/office/excel/2006/main",
            }
            for sp in tree.findall('.//x14:sparkline', NS):
                sq_elem = sp.find('xm:sqref', NS)
                f_elem  = sp.find('xm:f',     NS)
                if sq_elem is None or f_elem is None:
                    continue
                if sq_elem.text.replace('$', '').upper() == clean_cell:
                    grp = sp.getparent().getparent()
                    return {
                        "type":        grp.get("type", "line"),
                        "range":       f_elem.text.strip(),
                        "markers":     grp.get("markers",    "0") == "1",
                        "high_point":  grp.get("highPoint",  "0") == "1",
                        "low_point":   grp.get("lowPoint",   "0") == "1",
                        "first_point": grp.get("firstPoint", "0") == "1",
                        "last_point":  grp.get("lastPoint",  "0") == "1",
                        "negative":    grp.get("negative",   "0") == "1",
                    }
    except Exception:
        pass
    return None


def generate_rubric(check_map, p_wb_f, p_wb_v, p_cache=None):
    """
    Analyzes professor's answer cells (including sparklines) and generates a rubric via GPT.
    Returns: (rubric_map, task_summary)
      rubric_map: { (sheet, cell): rubric_item_dict }
    """
    if not client:
        return {}, ""

    cell_data = []
    for sn, cells in check_map.items():
        for c in cells:
            try:
                f = p_wb_f[sn][c].value
                v = p_wb_v[sn][c].value

                # Check if this cell has a sparkline (value/formula will be None)
                sl_info = extract_sparkline_info_for_cell(p_cache, c) if p_cache else None

                if sl_info:
                    # Sparkline cell — pass rich metadata so GPT understands it
                    active_markers = [k for k in ["markers", "high_point", "low_point",
                                                   "first_point", "last_point", "negative"]
                                      if sl_info.get(k)]
                    cell_data.append({
                        "sheet":   sn,
                        "cell":    c,
                        "type":    "sparkline",
                        "formula": f"Sparkline ({sl_info['type']})",
                        "value":   f"Range: {sl_info['range']}",
                        "sparkline_type":    sl_info["type"],
                        "sparkline_range":   sl_info["range"],
                        "active_markers":    active_markers if active_markers else "none",
                    })
                else:
                    # Normal formula/value cell
                    cell_data.append({
                        "sheet":   sn,
                        "cell":    c,
                        "type":    "formula" if str(f).startswith("=") else "value",
                        "formula": str(f),
                        "value":   str(v),
                    })
            except Exception:
                pass

    if not cell_data:
        return {}, ""

    prompt = f"""You are an Excel course instructor. Analyze the following answer key cells and generate a grading rubric.
Each cell entry includes a "type" field: "formula", "value", or "sparkline".
For sparkline cells, use the sparkline_type and sparkline_range fields to understand what is being tested.

Answer Key Data:
{json.dumps(cell_data, ensure_ascii=False, indent=2)}

For each cell, identify:
1. What Excel skill is being tested (e.g., VLOOKUP, SUM, Sparkline - Line Chart, Sparkline - Column with Markers, etc.)
2. A rubric criterion name (short, max 6 words)
3. Suggested point value (suggest reasonable values that total ~100)
4. What common mistakes to watch for

Return ONLY a valid JSON object in this exact format, no extra text:
{{
  "rubric_items": [
    {{
      "sheet": "Sheet1",
      "cell": "B5",
      "skill": "VLOOKUP Function",
      "criterion": "Correct VLOOKUP usage",
      "points": 10,
      "common_mistakes": "Wrong lookup range or missing absolute reference"
    }},
    {{
      "sheet": "Sheet1",
      "cell": "G5",
      "skill": "Sparkline - Line Chart",
      "criterion": "Correct sparkline data range",
      "points": 10,
      "common_mistakes": "Wrong data range or incorrect sparkline type selected"
    }}
  ],
  "task_summary": "Brief description of what this lab is testing overall"
}}"""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a professional Excel instructor. Return only valid JSON."},
                {"role": "user",   "content": prompt}
            ]
        )
        raw = response.choices[0].message.content.strip()
        raw = re.sub(r'^```json\s*', '', raw)
        raw = re.sub(r'\s*```$',     '', raw)
        rubric_data = json.loads(raw)

        rubric_map = {}
        for item in rubric_data.get("rubric_items", []):
            key = (item["sheet"], item["cell"])
            rubric_map[key] = item

        return rubric_map, rubric_data.get("task_summary", "")
    except Exception as e:
        st.warning(f"Rubric generation failed: {e}")
        return {}, ""

# ──────────────────────────────────────────────
# 11. PDF Report
# ──────────────────────────────────────────────
# Unicode → latin-1 safe conversion
UNICODE_REPLACEMENTS = {
    '\u2022': '-',   # bullet •
    '\u2013': '-',   # en dash
    '\u2014': '-',   # em dash
    '\u2018': "'",   # left single quote
    '\u2019': "'",   # right single quote
    '\u201c': '"',   # left double quote
    '\u201d': '"',   # right double quote
    '\u2026': '...', # ellipsis
    '\u00b7': '-',   # middle dot
    '\u2192': '->',  # arrow →
    '\u2260': '!=',  # not equal ≠
    '\u2264': '<=',  # less or equal ≤
    '\u2265': '>=',  # greater or equal ≥
}

def safe_text(text):
    """Convert any string to latin-1 safe text for FPDF."""
    if not text:
        return ""
    for char, replacement in UNICODE_REPLACEMENTS.items():
        text = text.replace(char, replacement)
    return text.encode('latin-1', 'replace').decode('latin-1')


def create_pdf_report(uid, score, errors, rubric_map=None):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(200, 15, txt="IS 2010 Lab Diagnostic Report", ln=True, align='C')
    pdf.set_font("Arial", '', 12)
    pdf.cell(200, 10, txt=f"Student ID: {uid} | Final Score: {score}", ln=True, align='C')
    pdf.line(10, 35, 200, 35)
    pdf.ln(10)

    if not errors:
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(200, 10, txt="Congratulations! Perfect Score.", ln=True, align='C')
    else:
        for i, err in enumerate(errors):
            pdf.set_font("Arial", 'B', 11)
            pdf.set_text_color(0, 0, 0)
            mode_label = err.get('Mode', 'Value Only')
            pdf.cell(0, 10, txt=safe_text(f"{i+1}. {err['Sheet']} / Cell {err['Cell']}  [{mode_label}]"), ln=True)

            pdf.set_font("Arial", '', 10)
            pdf.set_text_color(200, 0, 0)
            pdf.multi_cell(0, 7, txt=safe_text(f"Your Answer: {err['Stud_Ans']}"))
            pdf.set_text_color(0, 100, 0)
            pdf.multi_cell(0, 7, txt=safe_text(f"Expected Answer: {err['Prof_Ans']}"))

            # Issues list  (• replaced with - via safe_text)
            if err.get('Issues'):
                pdf.set_text_color(150, 75, 0)
                for issue in err['Issues']:
                    pdf.multi_cell(0, 6, txt=safe_text(f"  - {issue}"))

            # Rubric reference
            rubric_item = (rubric_map or {}).get((err['Sheet'], err['Cell']))
            if rubric_item:
                pdf.set_text_color(0, 0, 150)
                pdf.set_font("Arial", 'B', 10)
                rubric_str = f"[Rubric] {rubric_item['criterion']} ({rubric_item['points']}pts) -- {rubric_item.get('common_mistakes','')}"
                pdf.multi_cell(0, 7, txt=safe_text(rubric_str))

            # AI feedback
            rubric_ctx = rubric_item if rubric_item else None
            feedback = get_ai_feedback(err['Prof_Ans'], err['Stud_Ans'], issues=err.get('Issues'), rubric=rubric_ctx)
            pdf.set_fill_color(245, 245, 245)
            pdf.set_text_color(0, 50, 100)
            pdf.set_font("Arial", 'I', 10)
            pdf.multi_cell(0, 8, txt=safe_text(f"AI Feedback: {feedback or 'No feedback.'}"), border=1, fill=True)
            pdf.ln(5)

    return pdf.output(dest='S').encode('latin-1', 'ignore')

# ──────────────────────────────────────────────
# 12. Session State Init
# ──────────────────────────────────────────────
if 'grading_done' not in st.session_state:
    st.session_state.update({
        'grading_done': False,
        'summary_df': pd.DataFrame(),
        'all_wrongs_list': [],
        'total_questions': 0,
        'rubric_map': {},
        'rubric_summary': '',
        'color_mode_map': {},
    })

# ──────────────────────────────────────────────
# 13. UI: Color Guide
# ──────────────────────────────────────────────
with st.expander("Professor's Color Guide", expanded=True):
    cols = st.columns(10)
    st.warning(
        "※ **Crucial:** Use these 10 standard colors as **CELL SHADING (Fill Color)** to indicate graded cells. "
        "For each color, select the scoring mode below."
    )
    for idx, (name, hex_val) in enumerate(STANDARD_COLORS.items()):
        cols[idx].markdown(f"<div style='background-color:{hex_val}; height:20px; border-radius:3px;'></div>", unsafe_allow_html=True)
        cols[idx].caption(name)

# ──────────────────────────────────────────────
# 14. UI: Guidelines
# ──────────────────────────────────────────────
with st.expander("📘 Guided Lab Guidelines & Cautions", expanded=False):
    st.markdown("""
    #### 1. File Naming Convention (UNID Required)
    * **Student Files:** Filename must include your UID (e.g., **U1234567**).

    #### 2. Maintain Sheet Structure
    * Do not rename sheets or insert/delete rows or columns.

    #### 3. Precision in Formulas and Values
    * The system performs strict string comparison for formulas and numeric comparison for values.

    #### 4. Decimal Precision & Rounding Rules
    * Match the instructor's rounding method exactly (e.g., `=ROUND(calc, 2)`).

    #### 5. File Integrity
    * Only **.xlsx** supported. No password protection, no macros.
    """)

st.divider()

# ──────────────────────────────────────────────
# 15. File Upload
# ──────────────────────────────────────────────
c1, c2 = st.columns(2)
prof_file     = c1.file_uploader("1. Upload Professor's File", type=['xlsx'])
student_files = c2.file_uploader("2. Upload Student File(s)", type=['xlsx'], accept_multiple_files=True)

# ──────────────────────────────────────────────
# 16. Color → Scoring Mode Mapping UI
# ──────────────────────────────────────────────
if prof_file and student_files:
    st.subheader("🎨 Step 1: Assign Scoring Mode to Each Answer Color")
    st.caption(
        "The professor's file is scanned for colored cells. "
        "For each color found, assign its scoring mode. "
        "**All active colors are graded in a single run — no rerun needed.**"
    )

    # Read professor's file once to detect which standard colors are used
    # Must use read_only=False to reliably access fill formatting
    p_bytes_peek = prof_file.read()
    prof_file.seek(0)
    p_wb_peek = load_workbook(io.BytesIO(p_bytes_peek), read_only=False)
    used_colors = set()
    for sn in p_wb_peek.sheetnames:
        for row in p_wb_peek[sn].iter_rows():
            for cell in row:
                if cell.fill and cell.fill.fill_type == 'solid':
                    rgb = str(cell.fill.start_color.rgb)[-6:].upper()
                    used_colors.add(rgb)
    p_wb_peek.close()

    # Only show colors that actually appear in the professor's file
    active_standard = [(name, hx) for name, hx in STANDARD_COLORS.items()
                       if hx.lstrip('#').upper() in used_colors]

    color_mode_map = {}   # { hex_upper: mode_string }
    active_colors  = []

    if not active_standard:
        st.warning("⚠️ No standard colored cells found in the professor's file. Please fill answer cells with one of the 10 standard colors.")
    else:
        # Render one column per detected color — compact table-like row
        n = len(active_standard)
        grid = st.columns(n)
        for idx, (name, hex_val) in enumerate(active_standard):
            hex_upper = hex_val.lstrip('#').upper()
            with grid[idx]:
                # Color swatch
                st.markdown(
                    f"<div style='background:{hex_val};height:18px;border-radius:4px;margin-bottom:6px;'></div>",
                    unsafe_allow_html=True
                )
                mode = st.selectbox(
                    label=name,
                    options=SCORING_MODES,
                    format_func=lambda x: SCORING_LABELS[x],
                    key=f"mode_{name}",
                )
                color_mode_map[hex_upper] = [mode]   # store as list for uniform downstream handling
                active_colors.append(name)

        # Summary preview so professor can verify at a glance
        summary_md = "  |  ".join(
            f"<span style='color:{STANDARD_COLORS[n]};font-weight:bold'>■</span> {n} → {SCORING_LABELS[color_mode_map[STANDARD_COLORS[n].lstrip('#').upper()][0]]}"
            for n in active_colors
        )
        st.markdown(f"**Active mapping:** {summary_md}", unsafe_allow_html=True)

    # ──────────────────────────────────────────
    # 17. AI Rubric Generation (optional, before grading)
    # ──────────────────────────────────────────
    st.divider()
    st.subheader("🤖 Step 2: AI Rubric Auto-Generation (Optional)")
    if st.button("✨ Auto-Generate Rubric from Professor's File", disabled=(not active_colors)):
        with st.spinner("Analyzing formulas and generating rubric..."):
            p_bytes_r = prof_file.read()
            prof_file.seek(0)
            p_wb_f_r = load_workbook(io.BytesIO(p_bytes_r), data_only=False, read_only=False)
            p_wb_v_r = load_workbook(io.BytesIO(p_bytes_r), data_only=True,  read_only=False)

            # Extract XML cache for sparkline detection
            with zipfile.ZipFile(io.BytesIO(p_bytes_r)) as z:
                p_cache_r = {
                    f: z.read(f).decode('utf-8', errors='replace')
                    for f in z.namelist() if 'xl/worksheets/sheet' in f
                }

            # Build check_map using all colored cells
            temp_check_map = {}
            for sn in p_wb_f_r.sheetnames:
                cells = [
                    cell.coordinate
                    for row in p_wb_f_r[sn].iter_rows()
                    for cell in row
                    if cell.fill and cell.fill.fill_type == 'solid'
                    and str(cell.fill.start_color.rgb)[-6:].upper() in color_mode_map
                ]
                if cells:
                    temp_check_map[sn] = cells

            # Pass p_cache_r so sparkline cells are properly detected
            rubric_map, rubric_summary = generate_rubric(temp_check_map, p_wb_f_r, p_wb_v_r, p_cache=p_cache_r)
            st.session_state['rubric_map'] = rubric_map
            st.session_state['rubric_summary'] = rubric_summary

    if st.session_state.get('rubric_summary'):
        st.success(f"📝 Task Summary: {st.session_state['rubric_summary']}")
        if st.session_state.get('rubric_map'):
            rubric_df = pd.DataFrame([
                {
                    "Sheet": k[0], "Cell": k[1],
                    "Skill": v["skill"],
                    "Criterion": v["criterion"],
                    "Points": v["points"],
                    "Common Mistakes": v.get("common_mistakes", "")
                }
                for k, v in st.session_state['rubric_map'].items()
            ])
            st.dataframe(rubric_df, use_container_width=True, hide_index=True)

    # ──────────────────────────────────────────
    # 18. Grading Button
    # ──────────────────────────────────────────
    st.divider()
    st.subheader("📊 Step 3: Start Grading")

    # --- Pre-check Settings ---
    with st.expander("⚙️ Pre-check Settings (File Name & Sheet Name Validation)", expanded=True):
        pc_col1, pc_col2 = st.columns(2)
        with pc_col1:
            st.markdown("**📄 File Name Format**")
            filename_format = st.text_input(
                "Expected filename format (use [UID] as placeholder for student ID)",
                value="[UID]_Lab_3.xlsx",
                help="e.g. '[UID]_Lab_3.xlsx' → matches 'u1234567_Lab_3.xlsx'\nLeave blank to skip file name check."
            )
        with pc_col2:
            st.markdown("**📉 Penalty Points**")
            penalty_filename = st.number_input("File name mismatch penalty (pts)", min_value=0, max_value=20, value=5, step=1)
            penalty_sheet    = st.number_input("Sheet name mismatch penalty (pts, per missing sheet)", min_value=0, max_value=20, value=5, step=1)

        st.caption("Penalties are subtracted from the final score and shown in the Summary table.")

    if st.button("🚀 Start Grading Process", use_container_width=True, disabled=(not active_colors)):
        p_bytes = prof_file.read()

        with zipfile.ZipFile(io.BytesIO(p_bytes)) as z:
            p_cache = {
                f: z.read(f).decode('utf-8', errors='replace')
                for f in z.namelist() if 'xl/worksheets/sheet' in f
            }

        p_wb_f = load_workbook(io.BytesIO(p_bytes), data_only=False, read_only=False)
        p_wb_v = load_workbook(io.BytesIO(p_bytes), data_only=True,  read_only=False)

        # Build check_map: { sheet_name: [(cell_coord, [mode, ...]), ...] }
        check_map = {}
        for sn in p_wb_f.sheetnames:
            entries = []
            for row in p_wb_f[sn].iter_rows():
                for cell in row:
                    if cell.fill and cell.fill.fill_type == 'solid':
                        rgb = str(cell.fill.start_color.rgb)[-6:].upper()
                        if rgb in color_mode_map:
                            entries.append((cell.coordinate, color_mode_map[rgb]))  # modes is a list
            if entries:
                check_map[sn] = entries

        total_qs = sum(len(v) for v in check_map.values())
        summary_results = []
        all_wrongs = []
        uid_re = re.compile(r'[uU]\d{7}')
        progress_bar = st.progress(0)

        # Extract professor's sheet names for sheet name validation
        prof_sheetnames = set(p_wb_f.sheetnames)

        # Build filename regex from user-defined format
        # e.g. "[UID]_Lab_3.xlsx" → r"[uU]\d{7}_Lab_3\.xlsx"
        def build_filename_regex(fmt):
            if not fmt:
                return None
            escaped = re.escape(fmt)
            escaped = escaped.replace(r'\[UID\]', r'[uU]\d{7}')
            return re.compile(escaped, re.IGNORECASE)

        filename_regex = build_filename_regex(filename_format)

        for i, s_file in enumerate(student_files):
            with st.status(f"Grading {s_file.name}...", expanded=True) as status:
                s_bytes = s_file.read()
                with zipfile.ZipFile(io.BytesIO(s_bytes)) as z:
                    s_cache = {
                        f: z.read(f).decode('utf-8', errors='replace')
                        for f in z.namelist() if 'xl/worksheets/sheet' in f
                    }

                s_wb_f = load_workbook(io.BytesIO(s_bytes), data_only=False, read_only=False)
                s_wb_v = load_workbook(io.BytesIO(s_bytes), data_only=True,  read_only=False)
                uid = uid_re.search(s_file.name).group() if uid_re.search(s_file.name) else s_file.name

                # ── Pre-check: filename & sheet names ──────────────
                precheck_warnings = []
                penalty = 0

                # 1) File name — must match professor-defined format
                if filename_regex:
                    if not filename_regex.search(s_file.name):
                        msg = f"File name mismatch: expected format '{filename_format}', got '{s_file.name}' (-{penalty_filename}pts)"
                        precheck_warnings.append(msg)
                        penalty += penalty_filename
                        st.warning(f"⚠️ {msg}")
                    else:
                        st.success(f"✅ File name OK: {s_file.name}")

                # 2) Sheet names — must exactly match professor's sheet names
                stud_sheetnames = set(s_wb_f.sheetnames)
                missing_sheets  = prof_sheetnames - stud_sheetnames
                if missing_sheets:
                    for ms in sorted(missing_sheets):
                        msg = f"Sheet '{ms}' missing or renamed (-{penalty_sheet}pts)"
                        precheck_warnings.append(msg)
                        penalty += penalty_sheet
                        st.warning(f"⚠️ {msg}")
                else:
                    st.success("✅ All sheet names match")
                # ───────────────────────────────────────────────────

                correct = 0
                for sn, cell_entries in check_map.items():
                    if sn not in s_wb_f.sheetnames:
                        continue
                    for (c, modes) in cell_entries:
                        # Run each selected mode independently, collect all issues
                        cell_correct = True
                        cell_issues = []
                        stud_ans_str = "N/A"
                        prof_ans_str = "N/A"

                        for mode in modes:
                            is_correct, stud_ans, prof_ans, issues = grade_cell(
                                mode, p_wb_f, p_wb_v, s_wb_f, s_wb_v, p_cache, s_cache, sn, c
                            )
                            stud_ans_str = stud_ans  # same cell, same display value
                            prof_ans_str = prof_ans
                            if not is_correct:
                                cell_correct = False
                                # Tag each issue with which mode caught it
                                for iss in issues:
                                    cell_issues.append(f"[{SCORING_LABELS[mode]}] {iss}")

                        if cell_correct:
                            correct += 1
                        else:
                            all_wrongs.append({
                                "UnID":     uid,
                                "Sheet":    sn,
                                "Cell":     c,
                                "Mode":     " + ".join([SCORING_LABELS[m] for m in modes]),
                                "Stud_Ans": stud_ans_str,
                                "Prof_Ans": prof_ans_str,
                                "Issues":   cell_issues,
                            })

                status.update(label=f"✅ {uid} Done!", state="complete", expanded=False)

            final_score = max(correct - penalty, 0)
            summary_results.append({
                "UnID":       uid,
                "Score":      f"{correct}/{total_qs}",
                "Raw":        correct,
                "Penalty":    f"-{penalty}pts" if penalty > 0 else "—",
                "Warnings":   " | ".join(precheck_warnings) if precheck_warnings else "✅ OK",
            })
            progress_bar.progress((i + 1) / len(student_files))

        st.session_state.update({
            'summary_df': pd.DataFrame(summary_results).sort_values("Raw", ascending=False),
            'all_wrongs_list': all_wrongs,
            'grading_done': True,
            'total_questions': total_qs,
            'color_mode_map': color_mode_map,
        })
        st.rerun()

# ──────────────────────────────────────────────
# 19. Results Display
# ──────────────────────────────────────────────
if st.session_state['grading_done']:
    df_all_errors = pd.DataFrame(st.session_state['all_wrongs_list'])

    st.divider()
    col_chart, col_table = st.columns([1, 1])
    with col_chart:
        st.subheader("📊 Class Error Analysis")
        if not df_all_errors.empty:
            st.bar_chart(df_all_errors['Cell'].value_counts())
        else:
            st.success("No errors detected! 🎉")
    with col_table:
        st.subheader("📋 Score Summary")
        st.dataframe(
            st.session_state['summary_df'].drop(columns=["Raw"], errors='ignore'),
            use_container_width=True,
            hide_index=True
        )

    # Error detail by mode
    if not df_all_errors.empty:
        st.divider()
        st.subheader("🔍 Error Detail by Scoring Mode")
        if 'Mode' in df_all_errors.columns:
            # Mode column stores labels (e.g. "Value Only"), group by unique values
            for mode_label in df_all_errors['Mode'].unique():
                mode_df = df_all_errors[df_all_errors['Mode'] == mode_label]
                if not mode_df.empty:
                    with st.expander(f"{mode_label} Errors ({len(mode_df)})"):
                        display_df = mode_df[['UnID', 'Sheet', 'Cell', 'Stud_Ans', 'Prof_Ans']].copy()
                        st.dataframe(display_df, use_container_width=True, hide_index=True)
        else:
            st.info("No mode information available.")

    st.divider()
    st.subheader("📥 Download Reports")
    col_dl1, col_dl2 = st.columns(2)

    with col_dl1:
        xlsx_report = io.BytesIO()
        with pd.ExcelWriter(xlsx_report, engine='openpyxl') as writer:
            st.session_state['summary_df'].drop(columns=['Raw'], errors='ignore').to_excel(writer, index=False, sheet_name="Summary")
            if not df_all_errors.empty:
                export_df = df_all_errors.copy()
                if 'Issues' in export_df.columns:
                    export_df['Issues'] = export_df['Issues'].apply(lambda x: "; ".join(x) if isinstance(x, list) else x)
                export_df.to_excel(writer, index=False, sheet_name="All_Errors")
            if st.session_state.get('rubric_map'):
                rubric_export = pd.DataFrame([
                    {"Sheet": k[0], "Cell": k[1], **v}
                    for k, v in st.session_state['rubric_map'].items()
                ])
                rubric_export.to_excel(writer, index=False, sheet_name="Rubric")
        st.download_button(
            "📊 Download Excel Summary",
            xlsx_report.getvalue(),
            "IS2010_Results.xlsx",
            use_container_width=True
        )

    with col_dl2:
        if st.button("📄 Step 1: Generate AI Reports", use_container_width=True):
            zip_buffer = io.BytesIO()
            rubric_map = st.session_state.get('rubric_map', {})
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zf:
                for uid in st.session_state['summary_df']['UnID'].tolist():
                    with st.spinner(f"Generating AI Feedback for {uid}..."):
                        s_errs  = [e for e in st.session_state['all_wrongs_list'] if e['UnID'] == uid]
                        s_score = st.session_state['summary_df'][
                            st.session_state['summary_df']['UnID'] == uid
                        ]['Score'].values[0]
                        pdf_data = create_pdf_report(uid, s_score, s_errs, rubric_map=rubric_map)
                        zf.writestr(f"Report_{uid}.pdf", pdf_data)
            st.session_state['zip_data'] = zip_buffer.getvalue()
            st.success("✅ All reports generated!")

        if 'zip_data' in st.session_state:
            st.download_button(
                label="📥 Download All PDF Reports (ZIP)",
                data=st.session_state['zip_data'],
                file_name="Student_Reports.zip",
                mime="application/zip",
                use_container_width=True
            )
else:
    st.info("Upload files and configure scoring modes to begin.")
