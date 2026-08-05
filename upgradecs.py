# ********** Computer Science 9618 PYP Portal ***********
import io
import os
import re
import fitz  # PyMuPDF
import streamlit as st

# Word Document Libraries
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Google API Libraries
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from googleapiclient.errors import HttpError

# ==========================================
# 0. STREAMLIT PAGE CONFIG & CUSTOM STYLING
# ==========================================
st.set_page_config(
    page_title="9618 Computer Science PYP Archives", 
    page_icon="💻",
    layout="wide"
)

# Custom CSS Theme Implementation
st.markdown("""
    <style>
    /* Main window frame background */
    .stApp {
        background-color: #E6BBFC !important;
    }

    /* Sidebar background */
    [data-testid="stSidebar"] {
        background-color: #C663F8 !important;
    }
    
    /* Ensure sidebar inner elements remain transparent to show sidebar color */
    [data-testid="stSidebar"] > div:first-child {
        background-color: #C663F8 !important;
    }

    /* Parameter input bars, select boxes, text inputs, and buttons */
    div[data-baseweb="input"], 
    div[data-baseweb="select"] > div, 
    .stTextInput input, 
    .stSelectbox select,
    div[data-testid="stFileUploader"] {
        background-color: #BBFCBC !important;
        color: #FFFFFF !important;
        border-radius: 8px !important;
        border: 1px solid #620092 !important;
    }

    /* Input label text styling for readability */
    label, .stWidgetLabel p {
        color: #2D004B !important;
        font-weight: bold !important;
    }

    /* Input text color fix when typing */
    input {
        color: #070F9C !important;
    }
    
    /* Placeholders inside inputs */
    input::placeholder {
        color: #E0C2F7 !important;
    }
    </style>
""", unsafe_allow_html=True)


# ==========================================
# 1. CONFIGURATION & DRIVE FOLDER MAPPING
# ==========================================
SYLLABUS_CODE = "9618"

# Google Drive subfolder IDs
FOLDER_IDS = {
    "theory": "1BPW1HYttzzNLQ5j2HAlwEBU8qO-QuzBQ",    
    "practical": "18mm1ZI83hu8mvRivte43cedMp50-XSus", 
    "zips": "1V-p8DCoSik1_ghAY10cvdYKJ3GXfFT57"        
}

# Local folder names matching structure
LOCAL_FOLDERS = {
    "theory": "9618Theory",
    "practical": "9618Practical",
    "zips": "9618Zip"
}

# Ensure local storage directories exist
for folder_path in LOCAL_FOLDERS.values():
    os.makedirs(folder_path, exist_ok=True)


# ==========================================
# 2. HELPER FUNCTIONS: WORD DOCUMENT XML
# ==========================================
def add_page_number_to_run(run):
    """Adds a dynamic page number field to a Word document header run."""
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    
    r = run._r
    r.append(fldChar1)
    r.append(instrText)
    r.append(fldChar2)
    r.append(fldChar3)


# ==========================================
# 3. AUTOMATIC ROUTING & GOOGLE DRIVE API
# ==========================================
def determine_target_folder(filename: str) -> tuple[str, str]:
    """Determines destination folder based on Cambridge 9618 file naming conventions."""
    filename_lower = filename.lower()
    
    # 1. Zip / Source Files / Evidence Documents -> 9618Zip
    if filename_lower.endswith(".zip") or "_sf_" in filename_lower or "_evi_" in filename_lower or "_src_" in filename_lower:
        return "zips", "9618Zip (Source Files & Evidence Files)"
    
    # 2. Practical Papers (Paper 4) -> 9618Practical
    if re.search(r'_(qp|ms)_4[123]\b', filename_lower):
        return "practical", "9618Practical (Paper 4)"
    
    # 3. Theory Papers (Papers 1, 2, 3) -> 9618Theory
    if re.search(r'_(qp|ms)_(1[123]|2[123]|3[123])\b', filename_lower):
        return "theory", "9618Theory (Papers 1, 2, 3)"
    
    return None, None

def build_drive_service():
    """Authenticates and builds Google Drive API service using User OAuth 2.0 Credentials."""
    required_keys = ["refresh_token", "client_id", "client_secret"]
    missing_keys = [k for k in required_keys if k not in st.secrets]
    if missing_keys:
        st.error(f"❌ Missing Secret Key(s) in secrets.toml: {', '.join(missing_keys)}")
        return None
    try:
        creds = Credentials(
            token=None,
            refresh_token=st.secrets["refresh_token"],
            token_uri="https://oauth2.googleapis.com/token",
            client_id=st.secrets["client_id"],
            client_secret=st.secrets["client_secret"]
        )
        return build('drive', 'v3', credentials=creds)
    except Exception as e:
        st.error(f"❌ Google Drive Authentication Error: {e}")
        return None

def upload_file_to_drive(file_bytes, filename, folder_id, mime_type):
    """Uploads file to Google Drive using personal user storage quota."""
    service = build_drive_service()
    if not service:
        return None
    try:
        file_stream = io.BytesIO(file_bytes)
        file_metadata = {'name': filename, 'parents': [folder_id]}
        media = MediaIoBaseUpload(file_stream, mimetype=mime_type, resumable=True)
        uploaded_file = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
        return uploaded_file
    except Exception as error:
        st.error(f"❌ Drive API Upload Failed: {error}")
        return None

def sync_drive_folder_to_local(folder_key: str) -> tuple[int, str]:
    """Downloads files missing from local storage from the specified Google Drive subfolder."""
    service = build_drive_service()
    if not service:
        return 0, "Failed to authenticate."
    drive_folder_id = FOLDER_IDS[folder_key]
    local_path = LOCAL_FOLDERS[folder_key]
    try:
        query = f"'{drive_folder_id}' in parents and trashed = false"
        results = service.files().list(q=query, fields="files(id, name, mimeType)").execute()
        drive_files = results.get('files', [])
        downloaded_count = 0
        for file_info in drive_files:
            file_name = file_info['name']
            file_id = file_info['id']
            local_file_path = os.path.join(local_path, file_name)
            if not os.path.exists(local_file_path):
                request = service.files().get_media(fileId=file_id)
                with open(local_file_path, "wb") as f:
                    downloader = MediaIoBaseDownload(f, request)
                    done = False
                    while not done:
                        status, done = downloader.next_chunk()
                downloaded_count += 1
        return downloaded_count, f"Synced {downloaded_count} new file(s) for `{folder_key}`."
    except Exception as e:
        return 0, f"Sync error: {e}"

def perform_bulk_sync():
    """Syncs all three subfolders (Theory, Practical, Zip)."""
    total_synced = 0
    messages = []
    for f_key in ["theory", "practical", "zips"]:
        count, msg = sync_drive_folder_to_local(f_key)
        total_synced += count
        messages.append(msg)
    return total_synced, messages


# ==========================================
# 4. SEARCH ENGINE & PREVIEW HELPER FUNCTIONS
# ==========================================
def render_pdf_page_preview(filepath: str, page_num: int):
    """Renders a PDF page to PNG format for online preview."""
    try:
        doc = fitz.open(filepath)
        page = doc.load_page(page_num)
        pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
        img_bytes = pix.tobytes("png")
        doc.close()
        return img_bytes
    except Exception as e:
        st.error(f"Unable to render page preview: {e}")
        return None

def search_pdfs(keyword_list, folder_path, allowed_variants, match_mode="ALL"):
    """Scans local PDFs for specified target terms."""
    results = []
    if not os.path.exists(folder_path):
        return results
    cleaned_keywords = [k.strip().lower() for k in keyword_list if k.strip()]
    if not cleaned_keywords:
        return results

    for file in os.listdir(folder_path):
        if file.endswith(".pdf"):
            base_name = os.path.splitext(file)[0]
            if "_ci_" in file or "_sf_" in file:
                continue
                
            is_valid_variant = any(base_name.endswith(f"_{variant}") for variant in allowed_variants)
            if not is_valid_variant:
                continue

            filepath = os.path.join(folder_path, file)
            try:
                doc = fitz.open(filepath)
                for page_num in range(len(doc)):
                    page_text = doc[page_num].get_text()
                    matched_keywords_count = 0
                    for kw in cleaned_keywords:
                        escaped_kw = re.escape(kw)
                        pattern = r'\b' + escaped_kw + r'(s|es)?\b'
                        if re.search(pattern, page_text, re.IGNORECASE) or kw in page_text.lower():
                            matched_keywords_count += 1

                    if match_mode == "ALL" and matched_keywords_count == len(cleaned_keywords):
                        results.append({"file": file, "page": page_num, "path": filepath, "type": "QP" if "_qp_" in file else "MS"})
                    elif match_mode == "ANY" and matched_keywords_count > 0:
                        results.append({"file": file, "page": page_num, "path": filepath, "type": "QP" if "_qp_" in file else "MS"})
                doc.close()
            except Exception:
                continue
    return results


# ==========================================
# 5. APP STATE INITIALIZATION & AUTO-SYNC
# ==========================================
if 'handout_basket' not in st.session_state:
    st.session_state.handout_basket = []
if 'theory_results' not in st.session_state:
    st.session_state.theory_results = []
if 'practical_results' not in st.session_state:
    st.session_state.practical_results = []

if 'has_auto_synced' not in st.session_state:
    st.session_state.has_auto_synced = True
    with st.spinner("🚀 Waking up portal & auto-syncing files from Google Drive..."):
        perform_bulk_sync()


# ==========================================
# 6. STREAMLIT UI LAYOUT
# ==========================================
st.title("GCE A/AS LEVEL COMPUTER SCIENCE")
st.subheader("💻 9618 Computer Science PYP Resource Library")

# --- SIDEBAR CONTROLS ---
with st.sidebar:
    st.header("🔄 Google Drive Sync")
    if st.button("🔄 Sync Google Drive", type="primary", use_container_width=True):
        with st.spinner("Syncing Google Drive folders..."):
            synced_count, sync_msgs = perform_bulk_sync()
            st.success(f"🎉 Sync Complete! {synced_count} new file(s) downloaded.")
            for m in sync_msgs:
                st.caption(m)

    st.markdown("---")
    st.header("Handout Basket Summary")
    sidebar_metric_placeholder = st.empty()

    st.markdown("---")
    if st.button("🗑️ Clear Entire Basket", use_container_width=True):
        st.session_state.handout_basket = []
        st.rerun()

# --- NAVIGATION TABS ---
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "**🔍 Theory Search (P1, P2, P3)**", 
    "**💻 Practical Search (P4)**", 
    "**🛒 Handout Cart**", 
    "**📦 Source Files & Evidence**", 
    "**🔒 Upload PYP & Admin**"
])

# --- TAB 1: THEORY SEARCH ---
with tab1:
    st.header("Search CS Theory Papers (9618Theory)")
    st.caption("Focus Variants: June (13, 23, 33) | Nov (11, 22, 32)")
    
    col_t1_kw, col_t1_mode = st.columns([3, 1])
    with col_t1_kw:
        keyword_t1 = st.text_input("Enter Theory Keywords (comma-separated)", placeholder="e.g., binary tree, recursion, pipeline", key="t1_kw")
    with col_t1_mode:
        match_mode_t1 = st.selectbox("Search Match Mode", options=["Match ALL (AND)", "Match ANY (OR)"], key="t1_mode")

    if st.button("Search Theory Papers", type="primary"):
        if keyword_t1.strip():
            with st.spinner("Scanning 9618Theory PDFs..."):
                keywords = [k.strip() for k in keyword_t1.split(",") if k.strip()]
                theory_variants = ["13", "23", "33", "11", "22", "32", "12", "21", "31"]
                selected_mode = "ALL" if "ALL" in match_mode_t1 else "ANY"
                st.session_state.theory_results = search_pdfs(keywords, LOCAL_FOLDERS["theory"], theory_variants, match_mode=selected_mode)
        else:
            st.warning("Please enter at least one keyword.")

    if st.session_state.theory_results:
        st.write(f"Found **{len(st.session_state.theory_results)}** matching pages:")
        for idx, item in enumerate(st.session_state.theory_results):
            doc_kind = "📝 Question Paper" if item["type"] == "QP" else "🔑 Marking Scheme"
            with st.expander(f"📄 {item['file']} | {doc_kind} | Page {item['page'] + 1}"):
                c1, c2 = st.columns([3, 1])
                with c1:
                    preview_img = render_pdf_page_preview(item["path"], item["page"])
                    if preview_img:
                        st.image(preview_img, caption=f"Preview Page {item['page'] + 1}", use_container_width=True)
                with c2:
                    if st.button("➕ Add to Basket", key=f"add_t1_{idx}"):
                        st.session_state.handout_basket.append(item)
                        st.toast("Added to basket!")
                    if os.path.exists(item["path"]):
                        with open(item["path"], "rb") as pdf_file:
                            st.download_button(label="📥 Download Full PDF", data=pdf_file, file_name=item["file"], mime="application/pdf", key=f"dl_t1_{idx}")

# --- TAB 2: PRACTICAL SEARCH ---
with tab2:
    st.header("Search CS Practical Papers (9618Practical)")
    st.caption("Focus Variants: June (43) | Nov (42)")
    
    col_t2_kw, col_t2_mode = st.columns([3, 1])
    with col_t2_kw:
        keyword_t2 = st.text_input("Enter Practical Keywords (comma-separated)", placeholder="e.g., class, stack, queue, bubble sort", key="t2_kw")
    with col_t2_mode:
        match_mode_t2 = st.selectbox("Search Match Mode", options=["Match ALL (AND)", "Match ANY (OR)"], key="t2_mode")

    if st.button("Search Practical Papers", type="primary"):
        if keyword_t2.strip():
            with st.spinner("Scanning 9618Practical PDFs..."):
                keywords = [k.strip() for k in keyword_t2.split(",") if k.strip()]
                practical_variants = ["43", "42", "41"]
                selected_mode = "ALL" if "ALL" in match_mode_t2 else "ANY"
                st.session_state.practical_results = search_pdfs(keywords, LOCAL_FOLDERS["practical"], practical_variants, match_mode=selected_mode)
        else:
            st.warning("Please enter at least one keyword.")

    if st.session_state.practical_results:
        st.write(f"Found **{len(st.session_state.practical_results)}** matching pages:")
        for idx, item in enumerate(st.session_state.practical_results):
            doc_kind = "📝 Question Paper" if item["type"] == "QP" else "🔑 Marking Scheme"
            with st.expander(f"📄 {item['file']} | {doc_kind} | Page {item['page'] + 1}"):
                c1, c2 = st.columns([3, 1])
                with c1:
                    preview_img = render_pdf_page_preview(item["path"], item["page"])
                    if preview_img:
                        st.image(preview_img, caption=f"Preview Page {item['page'] + 1}", use_container_width=True)
                with c2:
                    if st.button("➕ Add to Basket", key=f"add_t2_{idx}"):
                        st.session_state.handout_basket.append(item)
                        st.toast("Added to basket!")
                    if os.path.exists(item["path"]):
                        with open(item["path"], "rb") as pdf_file:
                            st.download_button(label="📥 Download Full PDF", data=pdf_file, file_name=item["file"], mime="application/pdf", key=f"dl_t2_{idx}")

# --- TAB 3: HANDOUT CART & WORD BUILDER ---
with tab3:
    st.header("Worksheet / Handout Builder")
    if st.session_state.handout_basket:
        st.subheader(f"Selected Pages: {len(st.session_state.handout_basket)}")

        for idx, item in enumerate(st.session_state.handout_basket):
            col_info, col_remove = st.columns([4, 1])
            col_info.write(f"{idx + 1}. **{item['file']}** (Page {item['page'] + 1})")
            if col_remove.button("❌ Remove", key=f"remove_basket_{idx}"):
                st.session_state.handout_basket.pop(idx)
                st.rerun()

        st.markdown("---")
        if st.button("🪄 Export Handout to Word Document", type="primary"):
            try:
                doc = Document()
                section = doc.sections[0]

                section.orientation = WD_ORIENT.PORTRAIT
                section.page_width = Inches(8.5)
                section.page_height = Inches(11.5)
                section.top_margin = Inches(0.5)
                section.bottom_margin = Inches(0.5)
                section.left_margin = Inches(0.5)
                section.right_margin = Inches(0.5)

                header = section.header
                header_paragraph = header.paragraphs[0]
                header_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                header_run = header_paragraph.add_run("Page ")
                header_run.font.name = "Calibri"
                header_run.font.size = Pt(10)
                add_page_number_to_run(header_run)

                doc.add_heading(f'{SYLLABUS_CODE} Computer Science Handout', level=1)

                for i, item in enumerate(st.session_state.handout_basket):
                    doc.add_heading(f"Source: {item['file']} (Page {item['page'] + 1})", level=2)
                    
                    pdf_doc = fitz.open(item['path'])
                    page = pdf_doc.load_page(item['page'])
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    img_data = io.BytesIO(pix.tobytes("png"))
                    
                    doc.add_picture(img_data, width=Inches(7.0), height=Inches(9.2))
                    
                    if i < len(st.session_state.handout_basket) - 1:
                        doc.add_page_break()
                    pdf_doc.close()

                target_filename = f"{SYLLABUS_CODE}_CS_Handout.docx"
                doc.save(target_filename)

                with open(target_filename, "rb") as f:
                    st.download_button(label="📥 Click to Download Word Handout", data=f, file_name=target_filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            except Exception as e:
                st.error(f"❌ Handout Export Failed: {e}")
    else:
        st.info("Your basket is empty. Add pages from Tab 1 or Tab 2.")

# --- TAB 4: SOURCE FILES & EVIDENCE DOCUMENTS ---
with tab4:
    st.header("Download Practical Source Files & Evidence (9618Zip)")
    c1, c2, c3 = st.columns(3)
    with c1:
        z_year = st.selectbox("Select Year", [str(y) for y in range(2026, 2020, -1)], key="z_year_box")
    with c2:
        z_session = st.selectbox("Select Session", ["October/November (w)", "May/June (s)"], key="z_sess_box")
        session_code = z_session.split("(")[1].replace(")", "")
    with c3:
        z_paper_label = st.selectbox("Select Paper Component", ["Paper 43 (June)", "Paper 42 (Nov)", "Paper 41"], key="z_paper_box")
        z_paper_num = z_paper_label.split(" ")[1]

    short_year = z_year[-2:]
    possible_filenames = [
        f"{SYLLABUS_CODE}_{session_code}{short_year}_sf_{z_paper_num}.zip",
        f"{SYLLABUS_CODE}_{session_code}{short_year}_sf_{z_paper_num}.txt",
        f"{SYLLABUS_CODE}_{session_code}{short_year}_evi_{z_paper_num}.docx",
        f"{SYLLABUS_CODE}_{session_code}{short_year}_evi_{z_paper_num}.pdf",
    ]

    found_file_path = None
    matched_filename = None
    for fname in possible_filenames:
        check_path = os.path.join(LOCAL_FOLDERS["zips"], fname)
        if os.path.exists(check_path):
            found_file_path = check_path
            matched_filename = fname
            break

    st.markdown("---")
    if found_file_path and matched_filename:
        st.success(f"Found File: `{matched_filename}`")
        if matched_filename.endswith(".pdf"):
            mime_type = "application/pdf"
        elif matched_filename.endswith(".docx"):
            mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        elif matched_filename.endswith(".txt"):
            mime_type = "text/plain"
        else:
            mime_type = "application/zip"

        with open(found_file_path, "rb") as f:
            st.download_button(label=f"📥 Download {matched_filename}", data=f, file_name=matched_filename, mime=mime_type, key="dl_ci_file")
    else:
        st.warning(f"No Source File / Evidence Document found for `{SYLLABUS_CODE}_{session_code}{short_year}` {z_paper_label}.")

# --- TAB 5: ADMIN & DIRECT UPLOAD PANEL ---
with tab5:
    st.header("Admin Panel & Direct File Upload")
    admin_password = st.secrets.get("ADMIN_PASSWORD")
    if not admin_password:
        st.error("🚨 `ADMIN_PASSWORD` is not configured in Streamlit Secrets.")
    else:
        pwd = st.text_input("Enter Admin Password", type="password")
        if pwd == admin_password:
            st.success("Admin Access Granted")
            st.subheader("📤 Single File Direct Upload")
            uploaded_file = st.file_uploader("Upload Past Paper (PDF) or Source File (ZIP/TXT/DOCX)", type=["pdf", "zip", "txt", "docx"], key="admin_file_uploader")

            if uploaded_file is not None:
                folder_key, folder_name = determine_target_folder(uploaded_file.name)
                if folder_key is None:
                    st.warning(f"⚠️ Filename `{uploaded_file.name}` does not match 9618 naming conventions.")
                else:
                    st.info(f"🎯 Target Destination: **{folder_name}**")
                    if st.button("🚀 Upload File"):
                        with st.spinner("Uploading file..."):
                            file_bytes = uploaded_file.read()
                            local_dest_dir = LOCAL_FOLDERS[folder_key]
                            local_save_path = os.path.join(local_dest_dir, uploaded_file.name)
                            
                            # Save locally first
                            with open(local_save_path, "wb") as f:
                                f.write(file_bytes)
                                
                            # Upload to Google Drive using OAuth user credentials
                            drive_result = upload_file_to_drive(file_bytes, uploaded_file.name, FOLDER_IDS[folder_key], uploaded_file.type)
                            if drive_result:
                                st.success(f"✅ Uploaded `{uploaded_file.name}` successfully!")

# Populate metric display in sidebar
sidebar_metric_placeholder.metric(
    label="Saved Pages in Basket", 
    value=len(st.session_state.handout_basket)
)
