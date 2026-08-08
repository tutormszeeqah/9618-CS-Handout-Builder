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

# Google API Libraries (Service Account Authentication)
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

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
    .stApp {
        background-color: #E6BBFC !important;
    }
    [data-testid="stSidebar"] {
        background-color: #C663F8 !important;
    }
    [data-testid="stSidebar"] > div:first-child {
        background-color: #C663F8 !important;
    }
    div[data-baseweb="input"], 
    div[data-baseweb="select"] > div, 
    .stTextInput input, 
    .stSelectbox select {
        background-color: #BBFCBC !important;
        color: #070F9C !important;
        border-radius: 8px !important;
        border: 1px solid #620092 !important;
    }
    label, .stWidgetLabel p {
        color: #2D004B !important;
        font-weight: bold !important;
    }
    input {
        color: #070F9C !important;
    }
    </style>
""", unsafe_allow_html=True)


# ==========================================
# 1. CONFIGURATION & DIRECTORY SETUP
# ==========================================
SYLLABUS_CODE = "9618"
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

# Map local folder names to secret keys
LOCAL_FOLDERS = {
    "theory": "9618Theory",
    "practical": "9618Practical",
    "answer_scheme": "9618AnswerScheme",
    "zips": "9618Zip"
}

# Ensure local directories exist
for folder_path in LOCAL_FOLDERS.values():
    os.makedirs(folder_path, exist_ok=True)


# ==========================================
# 2. SERVICE ACCOUNT AUTHENTICATION & SYNC
# ==========================================
def build_drive_service():
    """Authenticates using Google Service Account credentials stored in st.secrets."""
    try:
        if "gcp_service_account" in st.secrets:
            service_account_info = dict(st.secrets["gcp_service_account"])
            creds = service_account.Credentials.from_service_account_info(
                service_account_info, 
                scopes=SCOPES
            )
            return build('drive', 'v3', credentials=creds)
        else:
            st.error("❌ Missing [gcp_service_account] section in secrets.toml.")
            return None
    except Exception as e:
        st.error(f"❌ Service Account Authentication Error: {e}")
        return None

def sync_drive_folder_to_local(folder_key: str) -> tuple[int, str]:
    """Downloads missing files from Google Drive to local directories."""
    service = build_drive_service()
    if not service:
        return 0, "Failed to authenticate Service Account."
    
    folder_ids = st.secrets.get("drive_folders", {})
    drive_folder_id = folder_ids.get(folder_key)
    
    if not drive_folder_id:
        return 0, f"Missing drive_folder_id for `{folder_key}` in secrets.toml."

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
                        _, done = downloader.next_chunk()
                downloaded_count += 1

        return downloaded_count, f"Synced {downloaded_count} file(s) for `{folder_key}`."
    except Exception as e:
        return 0, f"Sync error for `{folder_key}`: {e}"

def perform_bulk_sync():
    """Syncs all 4 configured Google Drive folders."""
    total_synced = 0
    messages = []
    for f_key in ["theory", "practical", "answer_scheme", "zips"]:
        count, msg = sync_drive_folder_to_local(f_key)
        total_synced += count
        messages.append(msg)
    return total_synced, messages


# ==========================================
# 3. HELPER FUNCTIONS: WORD DOCS & PREVIEWS
# ==========================================
def add_page_number_to_run(run):
    """Inserts a dynamic Word page number field."""
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

def render_pdf_page_preview(filepath: str, page_num: int = 0):
    """Renders a single PDF page into PNG image bytes for previewing."""
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


# ==========================================
# 4. APP STATE INITIALIZATION
# ==========================================
if 'handout_basket' not in st.session_state:
    st.session_state.handout_basket = []

if 'has_auto_synced' not in st.session_state:
    st.session_state.has_auto_synced = True
    with st.spinner("🚀 Waking up portal & auto-syncing files via Service Account..."):
        perform_bulk_sync()


# ==========================================
# 5. STREAMLIT UI LAYOUT
# ==========================================
st.title("PUSAT TINGKATAN ENAM SENGKURONG")
st.subheader(f"💻 {SYLLABUS_CODE} Computer Science PYP Resource Library")

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
    st.metric(label="Saved Pages in Basket", value=len(st.session_state.handout_basket))

    if st.button("🗑️ Clear Entire Basket", use_container_width=True):
        st.session_state.handout_basket = []
        st.rerun()

# --- NAVIGATION TABS ---
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "🔍 Theory Search", 
    "🔑 Answer Scheme", 
    "📅 View Exam Papers", 
    "📝 Export Handout", 
    "📦 Source Files", 
    "⚙️ Admin Dashboard"
])


# --- TAB 1: THEORY SEARCH ---
with tab1:
    st.header("Search CS Theory & Practical Papers")
    keyword_input = st.text_input("Enter Keywords", placeholder="e.g., binary tree, recursion, pipeline", key="t1_kw")

    if st.button("Search Papers", type="primary"):
        if keyword_input.strip():
            results = []
            keywords = [k.strip().lower() for k in keyword_input.split(",") if k.strip()]
            
            for folder_key in ["theory", "practical"]:
                folder_path = LOCAL_FOLDERS[folder_key]
                for file in os.listdir(folder_path):
                    if file.endswith(".pdf"):
                        filepath = os.path.join(folder_path, file)
                        try:
                            doc = fitz.open(filepath)
                            for page_num in range(len(doc)):
                                text = doc[page_num].get_text().lower()
                                if all(kw in text for kw in keywords):
                                    results.append({"file": file, "page": page_num, "path": filepath})
                            doc.close()
                        except Exception:
                            continue

            st.write(f"Found **{len(results)}** matching page(s):")
            for idx, item in enumerate(results):
                with st.expander(f"📄 {item['file']} | Page {item['page'] + 1}"):
                    c1, c2 = st.columns([3, 1])
                    with c1:
                        preview_img = render_pdf_page_preview(item["path"], item["page"])
                        if preview_img:
                            st.image(preview_img, use_container_width=True)
                    with c2:
                        if st.button("➕ Add to Basket", key=f"add_t1_{idx}"):
                            st.session_state.handout_basket.append(item)
                            st.toast("Added to basket!")


# --- TAB 2: ANSWER SCHEME FINDER ---
with tab2:
    st.header("🔑 Answer Scheme Finder (Marking Schemes)")
    st.caption("Locate, preview, and download official Cambridge Marking Schemes.")

    col_y, col_m, col_v = st.columns(3)
    with col_y:
        as_year = st.selectbox("Select Year", [str(y) for y in range(2026, 2020, -1)], key="as_yr")
    with col_m:
        as_month = st.selectbox("Select Session", ["June (s)", "November (w)"], key="as_mth")
        month_code = "s" if "June" in as_month else "w"
    with col_v:
        as_variant = st.selectbox(
            "Select Variant Component", 
            ["11", "12", "13", "21", "22", "23", "31", "32", "33", "41", "42", "43"], 
            key="as_var"
        )

    short_year = as_year[-2:]
    expected_ms_filename = f"{SYLLABUS_CODE}_{month_code}{short_year}_ms_{as_variant}.pdf"

    st.markdown("---")
    
    # Check both the dedicated Answer Scheme folder and local mirrors
    found_ms_path = None
    for folder_path in [LOCAL_FOLDERS["answer_scheme"], LOCAL_FOLDERS["theory"], LOCAL_FOLDERS["practical"]]:
        check_path = os.path.join(folder_path, expected_ms_filename)
        if os.path.exists(check_path):
            found_ms_path = check_path
            break

    if found_ms_path:
        st.success(f"✅ Found Answer Scheme: `{expected_ms_filename}`")
        
        col_btn1, col_btn2 = st.columns([1, 1])
        with col_btn1:
            with open(found_ms_path, "rb") as f:
                st.download_button(
                    label="📥 Download Answer Scheme PDF",
                    data=f,
                    file_name=expected_ms_filename,
                    mime="application/pdf",
                    type="primary",
                    use_container_width=True
                )
        
        with st.expander("👁️ Click to Expand & Preview Answer Scheme Document"):
            try:
                doc = fitz.open(found_ms_path)
                total_pages = len(doc)
                doc.close()
                st.info(f"Document contains {total_pages} page(s). Rendering previews below:")
                
                for p in range(total_pages):
                    img_data = render_pdf_page_preview(found_ms_path, p)
                    if img_data:
                        st.image(img_data, caption=f"Page {p + 1} of {total_pages}", use_container_width=True)
            except Exception as e:
                st.error(f"Error rendering PDF document: {e}")
    else:
        st.warning(f"⚠️ Answer Scheme `{expected_ms_filename}` was not found locally. Click 'Sync Google Drive' in the sidebar to download standard archives.")


# --- TAB 3: VIEW EXAM PAPERS ---
with tab3:
    st.header("📅 Download Full Question Papers & Marking Schemes")
    c1, c2, c3 = st.columns(3)
    with c1:
        v_year = st.selectbox("Year", [str(y) for y in range(2026, 2020, -1)], key="vp_yr")
    with c2:
        v_month = st.selectbox("Session", ["June (s)", "November (w)"], key="vp_mth")
        m_code = "s" if "June" in v_month else "w"
    with c3:
        v_paper = st.selectbox("Variant", ["11", "12", "13", "21", "22", "23", "31", "32", "33", "41", "42", "43"], key="vp_var")

    short_y = v_year[-2:]
    qp_name = f"{SYLLABUS_CODE}_{m_code}{short_y}_qp_{v_paper}.pdf"
    ms_name = f"{SYLLABUS_CODE}_{m_code}{short_y}_ms_{v_paper}.pdf"

    col_q, col_m = st.columns(2)
    with col_q:
        qp_path = os.path.join(LOCAL_FOLDERS["practical"] if v_paper.startswith("4") else LOCAL_FOLDERS["theory"], qp_name)
        if os.path.exists(qp_path):
            st.success(f"Found QP: `{qp_name}`")
            with open(qp_path, "rb") as f:
                st.download_button("📥 Download Question Paper", f, file_name=qp_name, use_container_width=True)
        else:
            st.info(f"QP `{qp_name}` not found locally.")

    with col_m:
        ms_path = os.path.join(LOCAL_FOLDERS["answer_scheme"], ms_name)
        if os.path.exists(ms_path):
            st.success(f"Found MS: `{ms_name}`")
            with open(ms_path, "rb") as f:
                st.download_button("📥 Download Marking Scheme", f, file_name=ms_name, use_container_width=True)
        else:
            st.info(f"MS `{ms_name}` not found locally.")


# --- TAB 4: EXPORT HANDOUT ---
with tab4:
    st.header("📝 Generate Custom Word Worksheet")
    if st.session_state.handout_basket:
        st.write(f"Selected Pages: **{len(st.session_state.handout_basket)}**")
        
        if st.button("🪄 Export Handout to Word (.docx)", type="primary"):
            try:
                doc = Document()
                section = doc.sections[0]
                section.orientation = WD_ORIENT.PORTRAIT
                section.top_margin = Inches(0.5)
                section.bottom_margin = Inches(0.5)

                header = section.header
                header_p = header.paragraphs[0]
                header_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                header_run = header_p.add_run("Page ")
                add_page_number_to_run(header_run)

                doc.add_heading(f'PTES {SYLLABUS_CODE} Computer Science Worksheet', level=1)

                for idx, item in enumerate(st.session_state.handout_basket):
                    doc.add_heading(f"Source: {item['file']} (Page {item['page'] + 1})", level=2)
                    pdf_doc = fitz.open(item['path'])
                    page = pdf_doc.load_page(item['page'])
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    img_data = io.BytesIO(pix.tobytes("png"))
                    doc.add_picture(img_data, width=Inches(6.5))
                    if idx < len(st.session_state.handout_basket) - 1:
                        doc.add_page_break()
                    pdf_doc.close()

                target_filename = f"{SYLLABUS_CODE}_CS_Worksheet.docx"
                doc.save(target_filename)
                with open(target_filename, "rb") as f:
                    st.download_button("📥 Download Word Document", f, file_name=target_filename)
            except Exception as e:
                st.error(f"Error generating Word file: {e}")
    else:
        st.info("Basket is empty. Add pages from Tab 1.")


# --- TAB 5: SOURCE FILES ---
with tab5:
    st.header("📦 Download Practical Source Files & Evidence")
    st.caption("Access ZIP source files and supporting document archives.")
    
    sf_files = os.listdir(LOCAL_FOLDERS["zips"])
    if sf_files:
        selected_sf = st.selectbox("Select File", sf_files)
        sf_path = os.path.join(LOCAL_FOLDERS["zips"], selected_sf)
        with open(sf_path, "rb") as f:
            st.download_button(f"📥 Download {selected_sf}", f, file_name=selected_sf, type="primary")
    else:
        st.info("No source files found in local storage.")


# --- TAB 6: ADMIN DASHBOARD (DIRECT DRIVE REDIRECTS) ---
with tab6:
    st.header("⚙️ Admin Dashboard")
    st.caption("Direct shortcuts to Google Drive folder dashboards for fast file uploads and management.")

    admin_pwd = st.secrets.get("ADMIN_PASSWORD", "")
    pwd_input = st.text_input("Enter Admin Password", type="password")

    if pwd_input and pwd_input == admin_pwd:
        st.success("🔓 Authenticated as Administrator")
        st.markdown("---")
        st.subheader("📁 Google Drive Upload Dashboards")
        st.write("Click any button below to open the corresponding Google Drive folder in a new tab for direct file uploads:")

        drive_links = st.secrets.get("drive_web_links", {})

        col_a, col_b = st.columns(2)
        with col_a:
            st.link_button(
                "📘 Open Theory Drive Folder", 
                drive_links.get("theory", "https://drive.google.com"), 
                type="primary", 
                use_container_width=True
            )
            st.markdown("<br>", unsafe_allow_html=True)
            st.link_button(
                "🔑 Open Answer Scheme Drive Folder", 
                drive_links.get("answer_scheme", "https://drive.google.com"), 
                type="primary", 
                use_container_width=True
            )

        with col_b:
            st.link_button(
                "💻 Open Practical Drive Folder", 
                drive_links.get("practical", "https://drive.google.com"), 
                type="primary", 
                use_container_width=True
            )
            st.markdown("<br>", unsafe_allow_html=True)
            st.link_button(
                "📦 Open Source File Zip Drive Folder", 
                drive_links.get("zips", "https://drive.google.com"), 
                type="primary", 
                use_container_width=True
            )

    elif pwd_input:
        st.error("❌ Incorrect Admin Password.")


# ==========================================
# 6. PORTAL FOOTER
# ==========================================
st.markdown("---")
SCHOOL_NAME = "Pusat Tingkatan Enam Sengkurong (PTES)"
SCHOOL_VISION = "Nurturing Resilient Leaders & Future-Ready Citizens"
DEVELOPER_NAME = "Miss Hajah Nurul Haziqah HN / Computer Science Department"

footer_html = f"""
<div style="text-align: center; padding: 15px 0px; color: #2D004B; font-family: sans-serif;">
    <p style="margin: 0; font-size: 1.0em; font-weight: bold;">🏫 {SCHOOL_NAME}</p>
    <p style="margin: 5px 0; font-size: 0.9em; font-style: italic;">"{SCHOOL_VISION}"</p>
    <p style="margin: 5px 0 0 0; font-size: 0.85em; font-weight: 600; color: #070F9C;">💻 Developed by {DEVELOPER_NAME}</p>
</div>
"""
st.markdown(footer_html, unsafe_allow_html=True)
