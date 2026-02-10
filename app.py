import streamlit as st
import pandas as pd
from docx import Document
import re
import io
import zipfile

# --- HELPER FUNCTIONS ---
def clean_filename(name):
    """Removes invalid characters."""
    return re.sub(r'[\\/*?:"<>|]', "", name)

def get_start_time(text):
    """Parses '8:30' from '8:30 A 10:00AM' -> (8, 30)."""
    if not isinstance(text, str): return None, None
    # Look for H:MM pattern
    match = re.search(r'(\d{1,2})[:.](\d{2})', text)
    if match:
        return int(match.group(1)), int(match.group(2))
    return None, None

def normalize_level(text):
    """Turns 'NIVEL 01' -> '1' for matching."""
    if not isinstance(text, str): return str(text)
    # Remove 'NIVEL' or 'LEVEL' and leading zeros
    clean = re.sub(r'^(LEVEL|NIVEL)\s*', '', text.strip(), flags=re.IGNORECASE)
    clean = re.sub(r'^0+', '', clean)
    return clean.upper()

# --- STREAMLIT UI ---
st.set_page_config(page_title="Generator: ProgMyJFeb2026", page_icon="📅")

st.title("📅 Document Generator (Level_Hour Mode)")
st.markdown("Upload your new lists. Filenames will be: `LEVEL 01_830A1000AM.docx`")

# 1. FILE UPLOADERS
col1, col2 = st.columns(2)
with col1:
    course_file = st.file_uploader("1. Upload Course CSV", type=["csv"])
with col2:
    links_file = st.file_uploader("2. Upload Links CSV", type=["csv"])

template_file = st.file_uploader("3. Upload Template (.docx)", type=["docx"])

# 2. SETTINGS
st.divider()
c1, c2 = st.columns(2)
with c1:
    date_text = st.text_input("Start Date", "24 de febrero de 2026")
with c2:
    days_text = st.text_input("Days Text", "TUESDAY TO FRIDAY") # Change if needed

# 3. GENERATE BUTTON
if st.button("🚀 Generate Files", type="primary"):
    if not course_file or not links_file or not template_file:
        st.error("Please upload all 3 files.")
    else:
        try:
            # Load Data
            courses_df = pd.read_csv(course_file, encoding='latin1')
            links_df = pd.read_csv(links_file, encoding='latin1')
            
            # Clean Headers
            courses_df.columns = [str(c).upper().strip() for c in courses_df.columns]
            links_df.columns = [str(c).upper().strip() for c in links_df.columns]
            
            zip_buffer = io.BytesIO()
            files_created = 0
            
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                progress_bar = st.progress(0)
                total_rows = len(courses_df)

                for index, row in courses_df.iterrows():
                    # Extract Data
                    level_raw = str(row.get('NIVEL', '')).strip()   # "NIVEL 01"
                    schedule_raw = str(row.get('HORARIO', '')).strip() # "8:30 A 10:00AM"
                    id_raw = str(row.get('ID', '')).replace('.0', '').strip()
                    
                    course_h, course_m = get_start_time(schedule_raw)
                    course_lvl_code = normalize_level(level_raw) # "1"

                    # FIND LINK
                    found_link = "LINK_NOT_FOUND"
                    if course_h is not None:
                        for _, link_row in links_df.iterrows():
                            link_h, link_m = get_start_time(str(link_row.get('HORA', '')))
                            link_lvl_code = normalize_level(str(link_row.get('LEVEL', '')))
                            
                            # Match: Same Hour, Same Minute, Same Level
                            if link_h == course_h and link_m == course_m and link_lvl_code == course_lvl_code:
                                found_link = str(link_row.get('LINK', 'MISSING_LINK'))
                                break
                    
                    # CREATE DOC
                    try:
                        template_file.seek(0)
                        doc = Document(template_file)
                        
                        for p in doc.paragraphs:
                            if "24 de" in p.text and "2025" in p.text:
                                p.text = re.sub(r'24 de \w+ de 2025', date_text, p.text, flags=re.IGNORECASE)
                            
                            if "{{LEVEL}}" in p.text: p.text = p.text.replace("{{LEVEL}}", level_raw)
                            if "{{ID}}" in p.text: p.text = p.text.replace("{{ID}}", id_raw)
                            if "{{WA_LINK}}" in p.text: p.text = p.text.replace("{{WA_LINK}}", found_link)
                            if "{{SCHEDULE}}" in p.text: 
                                p.text = p.text.replace("{{SCHEDULE}}", f"{days_text} / {schedule_raw}")

                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        
                        # FILENAME FORMAT: "LEVEL 01_830A1000AM.docx"
                        # 1. Clean Schedule (remove spaces, colons)
                        schedule_clean = schedule_raw.replace(":", "").replace(" ", "").replace("/", "")
                        # 2. Build Name
                        fname_str = f"{level_raw}_{schedule_clean}.docx"
                        fname = clean_filename(fname_str)
                        
                        zip_file.writestr(fname, doc_io.getvalue())
                        files_created += 1
                        
                    except Exception as e:
                        st.warning(f"Error row {index}: {e}")

                    progress_bar.progress((index + 1) / total_rows)

            st.success(f"✅ Generated {files_created} documents.")
            st.download_button(
                "📥 Download Zip",
                data=zip_buffer.getvalue(),
                file_name="MyJ_Feb2026_Docs.zip",
                mime="application/zip"
            )

        except Exception as e:
            st.error(f"Error: {e}")
