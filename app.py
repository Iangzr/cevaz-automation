import streamlit as st
import pandas as pd
from docx import Document
import re
import io
import zipfile

# --- HELPER FUNCTIONS ---
def clean_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "", name)

def get_start_time(text):
    """Parses '6:15' -> (6, 15)."""
    if not isinstance(text, str): return None, None
    match = re.search(r'(\d{1,2})[:.](\d{2})', text)
    if match:
        return int(match.group(1)), int(match.group(2))
    return None, None

def normalize_level(text):
    """Turns 'LEVEL 02B' -> '2B' for strict matching."""
    if not isinstance(text, str): return str(text)
    clean = re.sub(r'^(LEVEL|NIVEL)\s*', '', text.strip(), flags=re.IGNORECASE)
    clean = re.sub(r'^0+', '', clean)
    return clean.upper()

# --- STREAMLIT UI ---
st.set_page_config(page_title="CEVAZ Generator", page_icon="📄")

st.title("📄 GENERADOR DE INVITACIONES")
st.markdown("Upload your lists, and this tool will match links and generate the files for you.")

# 1. FILE UPLOADERS
col1, col2 = st.columns(2)
with col1:
    course_file = st.file_uploader("1. Upload Courses CSV", type=["csv"])
with col2:
    links_file = st.file_uploader("2. Upload Links CSV", type=["csv"])

template_file = st.file_uploader("3. Upload Word Template (.docx)", type=["docx"])

# 2. SETTINGS
st.divider()
st.subheader("⚙️ Settings")
c1, c2 = st.columns(2)
with c1:
    date_text = st.text_input("Start Date Text", "24 de febrero de 2026")
with c2:
    days_text = st.text_input("Days Text", "TUESDAY TO FRIDAY")

# 3. PROCESSING BUTTON
if st.button("🚀 Generate Documents", type="primary"):
    if not course_file or not links_file or not template_file:
        st.error("Please upload all 3 files (Courses, Links, and Template) to continue.")
    else:
        # --- LOGIC START ---
        try:
            # Load Data
            courses_df = pd.read_csv(course_file, encoding='latin1')
            links_df = pd.read_csv(links_file, encoding='latin1')
            
            # Normalize Columns
            courses_df.columns = [str(c).upper().strip() for c in courses_df.columns]
            links_df.columns = [str(c).upper().strip() for c in links_df.columns]
            
            # Prepare Zip Buffer
            zip_buffer = io.BytesIO()
            files_created = 0
            
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                
                # Progress Bar
                progress_bar = st.progress(0)
                total_rows = len(courses_df)

                for index, row in courses_df.iterrows():
                    # Parse Course Info
                    level_raw = str(row.get('NIVEL', '')).strip()
                    schedule_raw = str(row.get('HORARIO', '')).strip()
                    id_raw = str(row.get('ID', '')).replace('.0', '').strip()
                    
                    if not id_raw or id_raw.lower() == 'nan':
                        id_raw = f"Row{index+1}"

                    course_h, course_m = get_start_time(schedule_raw)
                    course_lvl_code = normalize_level(level_raw)

                    # FIND MATCH
                    found_link = "LINK_NOT_FOUND"
                    if course_h is not None:
                        for _, link_row in links_df.iterrows():
                            link_h, link_m = get_start_time(str(link_row.get('HORA', '')))
                            link_lvl_code = normalize_level(str(link_row.get('LEVEL', '')))
                            
                            # Strict Match
                            if link_h == course_h and link_m == course_m and link_lvl_code == course_lvl_code:
                                found_link = str(link_row.get('LINK', 'MISSING_LINK'))
                                break
                    
                    # GENERATE DOC (IN MEMORY)
                    try:
                        # Reset file pointer for template to read it fresh every time
                        template_file.seek(0)
                        doc = Document(template_file)
                        
                        for p in doc.paragraphs:
                            # Replacements
                            if "24 de" in p.text and "2025" in p.text:
                                p.text = re.sub(r'24 de \w+ de 2025', date_text, p.text, flags=re.IGNORECASE)
                            
                            if "{{LEVEL}}" in p.text: p.text = p.text.replace("{{LEVEL}}", level_raw)
                            if "{{ID}}" in p.text: p.text = p.text.replace("{{ID}}", id_raw)
                            if "{{WA_LINK}}" in p.text: p.text = p.text.replace("{{WA_LINK}}", found_link)
                            if "{{SCHEDULE}}" in p.text: 
                                p.text = p.text.replace("{{SCHEDULE}}", f"{days_text} / {schedule_raw}")

                        # Save to stream
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        
                        # Define Filename
                        schedule_safe = schedule_raw.replace(":", "").replace(" ", "").replace("/", "")
                        fname = clean_filename(f"{level_raw}_{schedule_safe}_{id_raw}.docx")
                        
                        # Add to Zip
                        zip_file.writestr(fname, doc_io.getvalue())
                        files_created += 1
                        
                    except Exception as e:
                        st.warning(f"Error on row {index}: {e}")

                    # Update Progress
                    progress_bar.progress((index + 1) / total_rows)

            # --- OUTPUT ---
            st.success(f"✅ Success! Generated {files_created} documents.")
            
            # Download Button
            st.download_button(
                label="📥 Download All Files (.zip)",
                data=zip_buffer.getvalue(),
                file_name="CEVAZ_Documents.zip",
                mime="application/zip"
            )

        except Exception as e:

            st.error(f"An error occurred: {e}")
