import streamlit as st
import pandas as pd
from docx import Document
import re
import io
import zipfile
import unicodedata

# --- HELPER FUNCTIONS ---
def clean_filename(name):
    """Removes invalid characters."""
    return re.sub(r'[\\/*?:"<>|]', "", name)

def normalize_text(text):
    """Removes accents and lowers case for comparison (jovenes == JÓVENES)."""
    if not isinstance(text, str): return str(text)
    # Normalize unicode characters (e.g., ó -> o)
    text = unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('utf-8')
    return text.lower().strip()

def get_start_time(text):
    """Parses '2:45' -> (2, 45)."""
    if not isinstance(text, str): return None, None
    match = re.search(r'(\d{1,2})[:.](\d{2})', text)
    if match:
        return int(match.group(1)), int(match.group(2))
    return None, None

def normalize_level(text):
    """Turns 'NIVEL 01' -> '1'."""
    if not isinstance(text, str): return str(text)
    clean = re.sub(r'^(LEVEL|NIVEL)\s*', '', text.strip(), flags=re.IGNORECASE)
    clean = re.sub(r'^0+', '', clean)
    return clean.upper()

# --- STREAMLIT UI ---
st.set_page_config(page_title="Generator: Multi-Mode", page_icon="⚡")

st.title("⚡ Universal Document Generator")
st.markdown("Works for **Adults** (Time Match) AND **Kids/Teens** (Category Match).")

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
    days_text = st.text_input("Days Text", "TUESDAY TO FRIDAY")

# 3. GENERATE BUTTON
if st.button("🚀 Generate Files", type="primary"):
    if not course_file or not links_file or not template_file:
        st.error("Please upload all 3 files.")
    else:
        try:
            # Load Data
            courses_df = pd.read_csv(course_file, encoding='latin1')
            links_df = pd.read_csv(links_file, encoding='latin1')
            
            # Clean Headers (Uppercased and Strip)
            courses_df.columns = [str(c).upper().strip() for c in courses_df.columns]
            links_df.columns = [str(c).upper().strip() for c in links_df.columns]
            
            # --- DETECT MODE ---
            # Check if Links file has "EDAD" (Category Mode) or "HORA" (Time Mode)
            MODE = "UNKNOWN"
            if 'EDAD' in links_df.columns:
                MODE = "CATEGORY"
                st.info("🔹 Mode Detected: **Category Matching** (Kids/Teens)")
            elif 'HORA' in links_df.columns:
                MODE = "TIME"
                st.info("🔹 Mode Detected: **Time Matching** (Adults)")
            
            # Smart Column Detection
            link_level_col = 'NIVEL' if 'NIVEL' in links_df.columns else 'LEVEL'
            
            zip_buffer = io.BytesIO()
            files_created = 0
            
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                progress_bar = st.progress(0)
                total_rows = len(courses_df)

                for index, row in courses_df.iterrows():
                    # Extract Common Data
                    level_raw = str(row.get('NIVEL', '')).strip()
                    schedule_raw = str(row.get('HORARIO', '')).strip()
                    id_raw = str(row.get('ID', '')).replace('.0', '').strip()
                    
                    # Normalize for matching
                    course_lvl_code = normalize_level(level_raw)
                    course_h, course_m = get_start_time(schedule_raw)

                    # Initialize Match
                    found_link = "LINK_NOT_FOUND"
                    category_prefix = "" # Used for filename

                    # --- MATCHING LOGIC ---
                    if MODE == "CATEGORY":
                        # Get Category from Course (e.g. "NINOS")
                        course_cat = normalize_text(str(row.get('CATEGORIA', ''))) # ninos
                        category_prefix = course_cat.upper() + "_"

                        for _, link_row in links_df.iterrows():
                            # Get Link Category (e.g. "KIDS")
                            link_cat_raw = str(link_row.get('EDAD', ''))
                            link_cat = normalize_text(link_cat_raw) # kids
                            link_lvl_code = normalize_level(str(link_row.get(link_level_col, '')))
                            
                            # 1. CHECK LEVEL
                            if link_lvl_code != course_lvl_code:
                                continue

                            # 2. CHECK CATEGORY (Smart Mapping)
                            # Map "ninos" -> "kids"
                            is_cat_match = False
                            if "nino" in course_cat and "kid" in link_cat: is_cat_match = True
                            elif "joven" in course_cat and "joven" in link_cat: is_cat_match = True
                            elif course_cat == link_cat: is_cat_match = True
                            
                            if is_cat_match:
                                found_link = str(link_row.get('LINK', 'MISSING_LINK'))
                                break

                    elif MODE == "TIME":
                        # Old Logic (Adults)
                        if course_h is not None:
                            for _, link_row in links_df.iterrows():
                                link_h, link_m = get_start_time(str(link_row.get('HORA', '')))
                                link_lvl_code = normalize_level(str(link_row.get(link_level_col, '')))
                                
                                if link_h == course_h and link_m == course_m and link_lvl_code == course_lvl_code:
                                    found_link = str(link_row.get('LINK', 'MISSING_LINK'))
                                    break
                    
                    # --- CREATE DOC ---
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
                        
                        # FILENAME: "JOVENES_LEVEL 01_245.docx"
                        schedule_safe = schedule_raw.replace(":", "").replace(" ", "").replace("/", "")
                        fname_str = f"{category_prefix}{level_raw}_{schedule_safe}.docx"
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
                file_name="Universal_Docs.zip",
                mime="application/zip"
            )

        except Exception as e:
            st.error(f"Error: {e}")
