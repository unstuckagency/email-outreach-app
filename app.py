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

def validate_mappings(templates: list[str], header_map: dict[str, str]):
    # Validate placeholders across ALL templates
    all_placeholders = []
    for t in templates:
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
            return ""  # should not happen if validation passed
        val = row.get(col, "")
        if pd.isna(val) or str(val).strip() == "":
            return blank_fill
        return str(val)
    return PLACEHOLDER_RE.sub(repl, template)

def find_email_column(df: pd.DataFrame) -> str | None:
    # preserve Email column if present (any casing/spaces/underscores)
    for col in df.columns:
        if norm_key(col) == "email":
            return col
    return None

# ---------------- UI ----------------

st.set_page_config(page_title="Outreach Merge Tool", layout="centered")
st.title("Outreach Merge Tool")

# Simple access gate (optional but recommended)
with st.expander("Access", expanded=True):
    pw = st.text_input("Team password", type="password")
    expected = st.secrets.get("APP_PASSWORD", "")
    if expected and pw != expected:
        st.warning("Enter the team password to use the tool.")
        st.stop()

uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
blank_fill = st.text_input("Blank cell replacement", value="[MISSING]")
st.caption("If a cell is blank/empty, it becomes the value above (use empty string if you prefer).")

st.subheader("Email templates (copy rotation)")

# Store templates in session state so the UI can add/remove dynamically
if "templates" not in st.session_state:
    st.session_state.templates = [""]  # start with one template

cols = st.columns([1, 1, 2])
with cols[0]:
    if st.button("Add template"):
        st.session_state.templates.append("")
with cols[1]:
    if st.button("Remove last") and len(st.session_state.templates) > 1:
        st.session_state.templates.pop()

# Render each template input
for i in range(len(st.session_state.templates)):
    st.session_state.templates[i] = st.text_area(
        f"Template {chr(65+i)}",
        value=st.session_state.templates[i],
        height=140,
        placeholder="Hi {{first_name}}, ...",
        key=f"tmpl_{i}"
    )

templates = [t for t in st.session_state.templates if t.strip()]

run = st.button(
    "Generate output CSV",
    type="primary",
    disabled=(uploaded is None or len(templates) == 0)
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
    mapping, missing_placeholders = validate_mappings(templates, header_map)

    # Hard stop if any placeholder doesn't map (requirement #5)
    if missing_placeholders:
        st.error("Some placeholders do not match any Excel column (case-insensitive; ignores spaces/underscores).")
        for ph in missing_placeholders:
            logs.append(f"UNMAPPED PLACEHOLDER: {{{{{ph}}}}}")
        st.code("\n".join(logs))
        st.stop()

    email_col = find_email_column(df)

    out_email = []
    out_copy = []

    # Generate (preserve row order) with rotation
    for i in range(len(df)):
        row = df.iloc[i]
        tmpl = templates[i % len(templates)]  # A->B->C->... rotation
        copy = merge_row(tmpl, row, mapping, blank_fill)
        out_copy.append(copy)

        if email_col:
            v = row.get(email_col, "")
            out_email.append("" if pd.isna(v) else str(v))

    # Output: only Email (if exists) + Email Copy
    if email_col:
        out_df = pd.DataFrame({"Email": out_email, "Email Copy": out_copy})
    else:
        out_df = pd.DataFrame({"Email Copy": out_copy})

    csv_bytes = out_df.to_csv(index=False).encode("utf-8")

    st.success(f"Done. Generated {len(out_df)} rows using {len(templates)} template(s) in rotation.")
    st.download_button(
        label="Download output.csv",
        data=csv_bytes,
        file_name="output.csv",
        mime="text/csv",
    )
