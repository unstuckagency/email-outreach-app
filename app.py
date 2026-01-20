import re
from io import BytesIO
import pandas as pd
import streamlit as st

PLACEHOLDER_RE = re.compile(r"\{\{\s*([^\}]+?)\s*\}\}")

def norm_key(s: str) -> str:
    # case-insensitive, ignore spaces + underscores
    return re.sub(r"[ _]+", "", str(s).strip().lower())

def extract_placeholders(template: str) -> list[str]:
    return PLACEHOLDER_RE.findall(template)

def build_header_map(df: pd.DataFrame) -> dict[str, str]:
    # {normalized_header: original_header} (first wins)
    m = {}
    for col in df.columns:
        k = norm_key(col)
        if k not in m:
            m[k] = col
    return m

def validate_mappings(all_templates: list[str], header_map: dict[str, str]):
    # Validate placeholders across ALL templates passed in
    all_placeholders = []
    for t in all_templates:
        all_placeholders.extend(extract_placeholders(t))

    missing = []
    mapping = {}
    for ph in all_placeholders:
        key = norm_key(ph)
        if key in header_map:
            mapping[ph] = header_map[key]
        else:
            missing.append(ph)

    # de-dupe missing while preserving order
    seen = set()
    missing_unique = []
    for m in missing:
        if m not in seen:
            missing_unique.append(m)
            seen.add(m)

    return mapping, missing_unique

def merge_row(template: str, row: pd.Series, mapping: dict[str, str], blank_fill: str) -> str:
    def repl(match: re.Match) -> str:
        raw = match.group(1)
        col = mapping.get(raw)
        if not col:
            return ""
        val = row.get(col, "")
        if pd.isna(val) or str(val).strip() == "":
            return blank_fill
        return str(val)
    return PLACEHOLDER_RE.sub(repl, template)

def find_email_column(df: pd.DataFrame) -> str | None:
    # match any "Email" variant (case/spaces/underscores)
    for col in df.columns:
        if norm_key(col) == "email":
            return col
    return None

def template_editor(title: str, session_key: str, min_templates: int = 1, help_text: str | None = None):
    """
    Renders a dynamic list of template text areas stored in st.session_state[session_key].
    Returns a list of non-empty templates (trimmed).
    """
    st.subheader(title)
    if help_text:
        st.caption(help_text)

    if session_key not in st.session_state:
        st.session_state[session_key] = [""] * max(1, min_templates)

    cols = st.columns([1, 1, 2])
    with cols[0]:
        if st.button(f"Add {title.lower()}", key=f"add_{session_key}"):
            st.session_state[session_key].append("")
    with cols[1]:
        if st.button("Remove last", key=f"rm_{session_key}") and len(st.session_state[session_key]) > min_templates:
            st.session_state[session_key].pop()

    for i in range(len(st.session_state[session_key])):
        label = f"{title} {chr(65+i)}"
        st.session_state[session_key][i] = st.text_area(
            label,
            value=st.session_state[session_key][i],
            height=120,
            placeholder="Use {{placeholders}} like {{first_name}}",
            key=f"{session_key}_{i}",
        )

    return [t.strip() for t in st.session_state[session_key] if t.strip()]

# ---------------- UI ----------------

st.set_page_config(page_title="Outreach Merge Tool", layout="centered")
st.title("Outreach Merge Tool")

with st.expander("Access", expanded=True):
    pw = st.text_input("Team password", type="password")
    expected = st.secrets.get("APP_PASSWORD", "")
    if expected and pw != expected:
        st.warning("Enter the team password to use the tool.")
        st.stop()

uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
blank_fill = st.text_input("Blank cell replacement", value="[MISSING]")
st.caption("If a cell is blank/empty, it becomes the value above (use empty string if you prefer).")

subject_templates = template_editor(
    "Subject template",
    session_key="subject_templates",
    min_templates=1,
    help_text="One or more subject lines. Rotates A → B → A… across rows."
)

email_templates = template_editor(
    "Email copy template",
    session_key="email_templates",
    min_templates=1,
    help_text="One or more main email bodies. Rotates A → B → A… across rows."
)

chaser_templates = template_editor(
    "Chaser copy template",
    session_key="chaser_templates",
    min_templates=0,
    help_text="Optional follow-up copy. If multiple are provided, rotates A → B → A…"
)

run = st.button(
    "Generate output XLSX",
    type="primary",
    disabled=(uploaded is None or len(subject_templates) == 0 or len(email_templates) == 0)
)

if run:
    logs = []

    # Read Excel
    try:
        df = pd.read_excel(uploaded, dtype=object)
    except Exception as e:
        st.error(f"Could not read Excel: {e}")
        st.stop()

    header_map = build_header_map(df)

    # Validate placeholders across subject + email + chaser
    all_templates = subject_templates + email_templates + chaser_templates
    mapping, missing_placeholders = validate_mappings(all_templates, header_map)

    # Hard stop if any placeholder doesn't map
    if missing_placeholders:
        st.error("Some placeholders do not match any Excel column header (case-insensitive; ignores spaces/underscores).")
        for ph in missing_placeholders:
            logs.append(f"UNMAPPED PLACEHOLDER: {{{{{ph}}}}}")
        st.code("\n".join(logs))
        st.stop()

    email_col = find_email_column(df)

    out_email_address = []
    out_subject = []
    out_email_copy = []
    out_email_sent = []
    out_chaser_copy = []
    out_chaser_sent = []
    out_status = []

    # Generate (preserve row order)
    for i in range(len(df)):
        row = df.iloc[i]

        subj_t = subject_templates[i % len(subject_templates)]
        body_t = email_templates[i % len(email_templates)]
        chaser_t = chaser_templates[i % len(chaser_templates)] if len(chaser_templates) > 0 else ""

        subject_line = merge_row(subj_t, row, mapping, blank_fill)
        email_copy = merge_row(body_t, row, mapping, blank_fill)
        chaser_copy = merge_row(chaser_t, row, mapping, blank_fill) if chaser_t else ""

        if email_col:
            v = row.get(email_col, "")
            out_email_address.append("" if pd.isna(v) else str(v))
        else:
            out_email_address.append("")

        out_subject.append(subject_line)
        out_email_copy.append(email_copy)
        out_email_sent.append("☐")     # tickbox placeholder
        out_chaser_copy.append(chaser_copy)
        out_chaser_sent.append("☐")    # tickbox placeholder
        out_status.append("")

    out_df = pd.DataFrame({
        "Email address": out_email_address,
        "Subject line": out_subject,
        "Email Copy": out_email_copy,
        "Email Sent?": out_email_sent,
        "Chaser copy": out_chaser_copy,
        "Chaser sent?": out_chaser_sent,
        "Status": out_status,
    })

    # Write XLSX to memory (WITH real Excel checkboxes)
    buffer = BytesIO()
    
    # Use XlsxWriter (openpyxl can't reliably add form checkboxes)
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        out_df.to_excel(writer, index=False, sheet_name="Outreach")
    
        workbook  = writer.book
        worksheet = writer.sheets["Outreach"]
    
        # Optional: nicer widths
        for col_idx, col_name in enumerate(out_df.columns):
            sample = out_df[col_name].astype(str).head(50)
            max_len = max([len(col_name)] + [len(x) for x in sample])
            worksheet.set_column(col_idx, col_idx, min(max(12, max_len + 2), 60))
    
        # ---- Insert REAL checkboxes ----
        # Pandas wrote headers on Excel row 0, data starts at row 1.
        email_sent_col = out_df.columns.get_loc("Email Sent?")
        chaser_sent_col = out_df.columns.get_loc("Chaser sent?")
    
        # Make the checkbox columns a bit narrower
        worksheet.set_column(email_sent_col, email_sent_col, 12)
        worksheet.set_column(chaser_sent_col, chaser_sent_col, 12)
    
        # Insert an unchecked checkbox for each data row
        for r in range(1, len(out_df) + 1):  # 1..n (since row 0 is headers)
            worksheet.insert_checkbox(r, email_sent_col, False)
            worksheet.insert_checkbox(r, chaser_sent_col, False)
    
    buffer.seek(0)

