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
    if not isinstance(text, str): return str(text)
    text = unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('utf-8')
    return text.lower().strip()

def get_start_time(text):
    if not isinstance(text, str): return None, None
    match = re.search(r'(\d{1,2})[:.](\d{2})', text)
    if match:
        return int(match.group(1)), int(match.group(2))
    return None, None

def normalize_level(text):
    if not isinstance(text, str): return str(text)
    clean = re.sub(r'^(LEVEL|NIVEL)\s*', '', text.strip(), flags=re.IGNORECASE)
    clean = re.sub(r'^0+', '', clean)
    return clean.upper()

def load_csv(file):
    try:
        file.seek(0)
        return pd.read_csv(file, encoding='utf-8')
    except UnicodeDecodeError:
        file.seek(0)
        return pd.read_csv(file, encoding='latin1')

# --- STREAMLIT UI ---
st.set_page_config(page_title="Generator: Perfect Filenames", page_icon="📝")

st.title("📝 Document Generator (Perfect Filenames)")
st.markdown("Preserves your template's formatting and creates clean filenames like: **Level 7 4:30 - 6:00 M-V**")

# 1. FILE UPLOADERS
col1, col2 = st.columns(2)
with col1:
    course_file = st.file_uploader("1. Upload Course CSV", type=["csv"])
with col2:
    links_file = st.file_uploader("2. Upload Links CSV", type=["csv"])

template_file = st.file_uploader("3. Upload Template (.docx)", type=["docx"])

# 2. SETTINGS
st.divider()
st.subheader("⚙️ Text & Filename Settings")
c1, c2, c3 = st.columns(3)
with c1:
    date_text = st.text_input("Start Date", "24 de febrero de 2026")
with c2:
    days_text = st.text_input("Days (Inside Doc)", "TUESDAY TO FRIDAY")
with c3:
    days_abbrev = st.text_input("Days Abbrev. (For Filename)", "M-V")

# 3. GENERATE BUTTON
if st.button("🚀 Generate Files", type="primary"):
    if not course_file or not links_file or not template_file:
        st.error("Please upload all 3 files.")
    else:
        try:
            courses_df = load_csv(course_file)
            links_df = load_csv(links_file)
            
            courses_df.columns = [str(c).upper().strip() for c in courses_df.columns]
            links_df.columns = [str(c).upper().strip() for c in links_df.columns]
            
            MODE = "UNKNOWN"
            if 'EDAD' in links_df.columns:
                MODE = "CATEGORY"
            elif 'HORA' in links_df.columns:
                MODE = "TIME"
            
            link_level_col = 'NIVEL' if 'NIVEL' in links_df.columns else 'LEVEL'
            
            zip_buffer = io.BytesIO()
            files_created = 0
            
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                progress_bar = st.progress(0)
                total_rows = len(courses_df)

                for index, row in courses_df.iterrows():
                    level_raw = str(row.get('NIVEL', '')).strip()
                    schedule_raw = str(row.get('HORARIO', '')).strip()
                    id_raw = str(row.get('ID', '')).replace('.0', '').strip()
                    
                    course_lvl_code = normalize_level(level_raw)
                    course_h, course_m = get_start_time(schedule_raw)

                    found_link = "LINK_NOT_FOUND"
                    category_prefix = "" 
                    type_label = "para adultos"

                    if MODE == "CATEGORY":
                        course_cat = normalize_text(str(row.get('CATEGORIA', '')))
                        category_prefix = course_cat.upper() + "_"

                        if "nino" in course_cat: type_label = "para niños"
                        elif "joven" in course_cat: type_label = "para jóvenes"

                        for _, link_row in links_df.iterrows():
                            link_cat_raw = str(link_row.get('EDAD', ''))
                            link_cat = normalize_text(link_cat_raw)
                            link_lvl_code = normalize_level(str(link_row.get(link_level_col, '')))
                            
                            if link_lvl_code != course_lvl_code: continue

                            is_cat_match = False
                            if "nino" in course_cat and "kid" in link_cat: is_cat_match = True
                            elif "joven" in course_cat and "joven" in link_cat: is_cat_match = True
                            elif course_cat == link_cat: is_cat_match = True
                            
                            if is_cat_match:
                                found_link = str(link_row.get('LINK', 'MISSING_LINK'))
                                break

                    elif MODE == "TIME":
                        if course_h is not None:
                            for _, link_row in links_df.iterrows():
                                link_h, link_m = get_start_time(str(link_row.get('HORA', '')))
                                link_lvl_code = normalize_level(str(link_row.get(link_level_col, '')))
                                
                                if link_h == course_h and link_m == course_m and link_lvl_code == course_lvl_code:
                                    found_link = str(link_row.get('LINK', 'MISSING_LINK'))
                                    break
                    
                    # --- CREATE DOC (FORMAT PRESERVING) ---
                    try:
                        template_file.seek(0)
                        doc = Document(template_file)
                        
                        replacements = {
                            "{{LEVEL}}": level_raw,
                            "{{ID}}": id_raw,
                            "{{WA_LINK}}": found_link,
                            "{{SCHEDULE}}": f"{days_text} / {schedule_raw}",
                            "{{TYPE}}": type_label
                        }

                        if type_label != "para adultos":
                            replacements["para adultos"] = type_label

                        for p in doc.paragraphs:
                            for run in p.runs:
                                for key, val in replacements.items():
                                    if key in run.text:
                                        run.text = run.text.replace(key, val)
                                
                                if "24 de" in run.text and "2025" in run.text:
                                    run.text = re.sub(r'24 de \w+ de 2025', date_text, run.text, flags=re.IGNORECASE)

                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        
                        # --- FILENAME FORMATTING MAGIC ---
                        # 1. Format Level: "LEVEL 07" -> "Level 7"
                        lvl_num_match = re.search(r'\d+', level_raw)
                        if lvl_num_match:
                            lvl_num = str(int(lvl_num_match.group())) # removes leading zero
                            lvl_clean = f"Level {lvl_num}"
                        else:
                            lvl_clean = level_raw

                        # 2. Format Schedule: "4:30 A 06:00PM" -> "4:30 - 6:00"
                        times = re.findall(r'(\d{1,2})[:.](\d{2})', schedule_raw)
                        if len(times) >= 2:
                            h1, m1 = int(times[0][0]), times[0][1]
                            h2, m2 = int(times[1][0]), times[1][1]
                            sched_clean = f"{h1}:{m1:02d} - {h2}:{m2:02d}"
                        else:
                            # Fallback if it can't find two times
                            sched_clean = schedule_raw.replace(":", "").replace("/", "-") 
                            
                        # 3. Combine: "JOVENES_Level 7 4:30 - 6:00 M-V.docx"
                        fname_str = f"{category_prefix}{lvl_clean} {sched_clean} {days_abbrev}.docx"
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
                file_name="Final_Invitations.zip",
                mime="application/zip"
            )

        except Exception as e:
            st.error(f"Error: {e}")
