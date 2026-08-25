import os
import io
import tempfile
import fitz  # PyMuPDF
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

# ==========================================
# 0. STREAMLIT PAGE CONFIG & CUSTOM STYLING
# ==========================================
st.set_page_config(
    page_title="9618 Computer Science PYP Archives", 
    page_icon="💻",
    layout="wide"
)

# Custom CSS Theme Implementation (Preserving your exact color scheme)
st.markdown("""
    <style>
    /* Main Page Background */
    .stApp {
        background-color: #E6BBFC !important;
    }
    
    /* Sidebar Background */
    [data-testid="stSidebar"], [data-testid="stSidebar"] > div:first-child {
        background-color: #C663F8 !important;
    }
    
    /* Ensure all Markdown and standard text in Sidebar is crisp and dark */
    [data-testid="stSidebar"] p, 
    [data-testid="stSidebar"] span, 
    [data-testid="stSidebar"] div,
    [data-testid="stSidebar"] h1, 
    [data-testid="stSidebar"] h2, 
    [data-testid="stSidebar"] h3 {
        color: #2D004B !important;
    }

    /* Metric Label and Value Contrast Fix in Sidebar */
    [data-testid="stMetricValue"] {
        color: #2D004B !important;
        font-weight: bold !important;
    }
    [data-testid="stMetricLabel"] p {
        color: #2D004B !important;
    }

    /* Input Fields & Select Boxes */
    div[data-baseweb="input"], 
    div[data-baseweb="select"] > div, 
    .stTextInput input, 
    .stSelectbox select {
        background-color: #BBFCBC !important;
        color: #070F9C !important;
        border-radius: 8px !important;
        border: 1px solid #620092 !important;
    }

    /* Labels */
    label, .stWidgetLabel p {
        color: #2D004B !important;
        font-weight: bold !important;
    }
    
    input {
        color: #070F9C !important;
    }

    /* Custom Button Styling */
    div.stButton > button {
        background-color: #BBFCBC !important;
        color: #070F9C !important;
        border: 1px solid #620092 !important;
        border-radius: 8px !important;
        font-weight: bold !important;
    }
    
    div.stButton > button:hover {
        background-color: #9bf09d !important;
        color: #070F9C !important;
    }
    </style>
""", unsafe_allow_html=True)

# Local folder directories mapped to folder keys
LOCAL_FOLDERS = {
    "paper1": "9618HandOuts/9618P1C1_C8",
    "paper2": "9618HandOuts/9618P2C9_C12",
    "paper3": "9618HandOuts/9618P3C13_C16",
    "paper4": "9618HandOuts/9618P4C17_C20",
    "other_notes": "9618HandOuts/Other9618Notes"
}

# Chapter options mapping per paper
PAPER_CHAPTER_MAPPING = {
    "paper1": [f"Chapter {i}" for i in range(1, 9)],
    "paper2": [f"Chapter {i}" for i in range(9, 13)],
    "paper3": [f"Chapter {i}" for i in range(13, 17)],
    "paper4": [f"Chapter {i}" for i in range(17, 21)]
}

# Ensure directories exist locally
for folder in LOCAL_FOLDERS.values():
    os.makedirs(folder, exist_ok=True)

# Initialize Session State
if "handout_basket" not in st.session_state:
    st.session_state.handout_basket = []

for p_key in ["paper1", "paper2", "paper3", "paper4"]:
    if f"{p_key}_results" not in st.session_state:
        st.session_state[f"{p_key}_results"] = []


# ==========================================
# 1. GOOGLE DRIVE SERVICES & SYNC ENGINE
# ==========================================
def get_gdrive_service():
    """Authenticates using Streamlit Service Account Secrets."""
    scopes = ["https://www.googleapis.com/auth/drive"]
    creds_dict = st.secrets["gcp_service_account"]
    credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return build("drive", "v3", credentials=credentials)

def sync_single_folder(service, folder_id: str, local_dir: str) -> tuple[int, str]:
    """Downloads missing PDF files from Google Drive to local directory."""
    if not folder_id:
        return 0, f"No Folder ID configured for directory {local_dir}"
    
    query = f"'{folder_id}' in parents and mimeType='application/pdf' and trashed=false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    drive_files = results.get("files", [])
    
    download_count = 0
    for file in drive_files:
        local_filepath = os.path.join(local_dir, file["name"])
        if not os.path.exists(local_filepath):
            request = service.files().get_media(fileId=file["id"])
            with open(local_filepath, "wb") as f:
                downloader = MediaIoBaseDownload(f, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
            download_count += 1
            
    total_files = len([f for f in os.listdir(local_dir) if f.endswith(".pdf")])
    return download_count, f"Synced {download_count} new file(s) for `{os.path.basename(local_dir)}` (Total local: {total_files})."

def perform_bulk_sync():
    """Syncs all 5 Google Drive folders to local storage."""
    try:
        service = get_gdrive_service()
        folder_secrets = st.secrets["gdrive_folders"]
        
        folder_mapping = {
            "paper1": (folder_secrets.get("paper1_id"), LOCAL_FOLDERS["paper1"]),
            "paper2": (folder_secrets.get("paper2_id"), LOCAL_FOLDERS["paper2"]),
            "paper3": (folder_secrets.get("paper3_id"), LOCAL_FOLDERS["paper3"]),
            "paper4": (folder_secrets.get("paper4_id"), LOCAL_FOLDERS["paper4"]),
            "other_notes": (folder_secrets.get("other_notes_id"), LOCAL_FOLDERS["other_notes"]),
        }
        
        total_downloaded = 0
        messages = []
        for key, (f_id, l_dir) in folder_mapping.items():
            dl_cnt, msg = sync_single_folder(service, f_id, l_dir)
            total_downloaded += dl_cnt
            messages.append(msg)
            
        return total_downloaded, messages
    except Exception as e:
        return 0, [f"Error during Drive Sync: {str(e)}"]

def upload_file_to_gdrive(uploaded_file, target_folder_key: str) -> tuple[bool, str]:
    """Uploads user selected PDF from Admin Tab to Google Drive."""
    try:
        service = get_gdrive_service()
        folder_id = st.secrets["gdrive_folders"].get(f"{target_folder_key}_id")
        
        if not folder_id:
            return False, "Folder ID not found in secrets."

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = tmp_file.name

        file_metadata = {
            "name": uploaded_file.name,
            "parents": [folder_id]
        }
        media = MediaFileUpload(tmp_path, mimetype="application/pdf")
        
        service.files().create(body=file_metadata, media_body=media, fields="id").execute()
        os.remove(tmp_path)
        
        # Trigger immediate sync to update local storage
        perform_bulk_sync()
        return True, f"Successfully uploaded `{uploaded_file.name}` to Google Drive!"
    except Exception as e:
        return False, f"Upload failed: {str(e)}"


# ==========================================
# 2. SEARCH & PDF PREVIEW GENERATOR
# ==========================================
def render_pdf_page_preview(filepath: str, page_num: int) -> bytes:
    """Renders a PDF page to PNG image bytes."""
    try:
        doc = fitz.open(filepath)
        page = doc[page_num]
        pix = page.get_pixmap(dpi=150)
        img_bytes = pix.tobytes("png")
        doc.close()
        return img_bytes
    except Exception:
        return None

def execute_chapter_search(paper_key: str, keyword_string: str, selected_chapter: str) -> list[dict]:
    """Searches paper folder and other notes for keywords."""
    results = []
    keywords = [k.strip().lower() for k in keyword_string.split(",") if k.strip()]
    target_folders = [LOCAL_FOLDERS[paper_key], LOCAL_FOLDERS["other_notes"]]

    for folder_path in target_folders:
        if not os.path.exists(folder_path):
            continue

        for file in os.listdir(folder_path):
            if not file.endswith(".pdf"):
                continue

            if selected_chapter != "All Chapters" and folder_path == LOCAL_FOLDERS[paper_key]:
                chap_num = selected_chapter.split(" ")[1]
                file_lower = file.lower()
                
                p_ch = f"ch{chap_num}"
                p_chap = f"chapter{chap_num}"
                p_chap_space = f"chapter {chap_num}"
                
                if p_ch not in file_lower and p_chap not in file_lower and p_chap_space not in file_lower:
                    continue

            filepath = os.path.join(folder_path, file)
            try:
                doc = fitz.open(filepath)
                for page_num in range(len(doc)):
                    raw_text = doc[page_num].get_text()
                    normalized_text = " ".join(raw_text.lower().split())
                    
                    if all(kw in normalized_text for kw in keywords):
                        results.append({
                            "file": file, 
                            "page": page_num, 
                            "path": filepath
                        })
                doc.close()
            except Exception:
                continue

    return results


# ==========================================
# 3. WORD HANDOUT GENERATOR
# ==========================================
def generate_docx_handout(basket_items: list[dict]) -> io.BytesIO:
    """Compiles selected PDF page snapshots into a Word Document (.docx)."""
    doc = Document()
    
    title = doc.add_heading("A-Level Computer Science (9618) Handout", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    subtitle = doc.add_paragraph("Custom Topical Reference Materials")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.runs[0].font.italic = True
    subtitle.runs[0].font.size = Pt(11)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    for idx, item in enumerate(basket_items, start=1):
        heading = doc.add_heading(f"Item {idx}: {item['file']} (Page {item['page'] + 1})", level=2)
        heading.paragraph_format.space_before = Pt(12)
        heading.paragraph_format.space_after = Pt(6)

        img_bytes = render_pdf_page_preview(item["path"], item["page"])
        if img_bytes:
            image_stream = io.BytesIO(img_bytes)
            doc.add_picture(image_stream, width=Inches(6.0))
            last_paragraph = doc.paragraphs[-1]
            last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            last_paragraph.paragraph_format.space_after = Pt(18)

    output_stream = io.BytesIO()
    doc.save(output_stream)
    output_stream.seek(0)
    return output_stream


# ==========================================
# 4. TAB CONTENT RENDERER
# ==========================================
def render_paper_tab(tab_object, paper_key: str, paper_title: str):
    with tab_object:
        st.header(f"🔍 Search {paper_title}")
        
        col_kw, col_ch = st.columns([2, 1])
        with col_kw:
            kw = st.text_input("Enter Keyword", placeholder="e.g., virtual memory, binary, stack", key=f"{paper_key}_kw")
        with col_ch:
            chapter_options = ["All Chapters"] + PAPER_CHAPTER_MAPPING[paper_key]
            selected_ch = st.selectbox("Select Chapter Filter", options=chapter_options, key=f"{paper_key}_chap")

        if st.button(f"Search {paper_title}", key=f"btn_{paper_key}"):
            if kw.strip():
                with st.spinner(f"Scanning {paper_title} & Other Notes..."):
                    st.session_state[f"{paper_key}_results"] = execute_chapter_search(paper_key, kw, selected_ch)
            else:
                st.warning("Please enter a keyword.")

        results = st.session_state[f"{paper_key}_results"]
        if results:
            st.write(f"Found **{len(results)}** matching page(s):")
            for idx, item in enumerate(results):
                with st.expander(f"📄 {item['file']} | Page {item['page'] + 1}"):
                    c1, c2 = st.columns([3, 1])
                    with c1:
                        preview_img = render_pdf_page_preview(item["path"], item["page"])
                        if preview_img:
                            st.image(preview_img, use_container_width=True)
                    with c2:
                        if st.button("➕ Add to Cart", key=f"add_{paper_key}_{idx}"):
                            st.session_state.handout_basket.append(item)
                            st.toast(f"Added Page {item['page'] + 1} to basket!")
                            st.rerun()
                        
                        st.markdown("<br>", unsafe_allow_html=True)
                        with open(item["path"], "rb") as pdf_f:
                            st.download_button(
                                label="📥 Download Target PDF",
                                data=pdf_f,
                                file_name=item["file"],
                                mime="application/pdf",
                                key=f"dl_{paper_key}_{idx}"
                            )


# ==========================================
# 5. MAIN APPLICATION & NAVIGATION
# ==========================================
st.title("💻 Cambridge 9618 Computer Science PYP Archives")
st.markdown("Search topical notes, preview past paper pages, and compile custom `.docx` handouts.")

# Sidebar
with st.sidebar:
    st.header("🔄 Google Drive Sync")
    if st.button("🔄 Sync Google Drive", use_container_width=True):
        with st.spinner("Syncing Google Drive folders..."):
            count, msgs = perform_bulk_sync()
            st.success(f"🎉 Sync Complete! {count} new file(s) downloaded.")

    st.markdown("---")
    
    st.subheader("📁 Local Storage Status")
    for key, folder_path in LOCAL_FOLDERS.items():
        if os.path.exists(folder_path):
            file_count = len([f for f in os.listdir(folder_path) if f.endswith(".pdf")])
            st.write(f"• **{key.capitalize()}**: `{file_count}` file(s)")
        else:
            st.write(f"• **{key.capitalize()}**: `Directory missing`")

    st.markdown("---")
    st.metric(label="Saved Pages in Basket", value=len(st.session_state.handout_basket))

    if st.button("🗑️ Clear Entire Basket", use_container_width=True):
        st.session_state.handout_basket = []
        st.rerun()

# Tabs
tab1, tab2, tab3, tab4, tab_cart, tab_admin = st.tabs([
    "📘 Paper 1 (Ch 1–8)",
    "📗 Paper 2 (Ch 9–12)",
    "📙 Paper 3 (Ch 13–16)",
    "📕 Paper 4 (Ch 17–20)",
    "🛒 Basket / Cart",
    "⚙️ Admin / Upload"
])

# Render Paper Search Tabs
render_paper_tab(tab1, "paper1", "Paper 1 (Chapters 1–8)")
render_paper_tab(tab2, "paper2", "Paper 2 (Chapters 9–12)")
render_paper_tab(tab3, "paper3", "Paper 3 (Chapters 13–16)")
render_paper_tab(tab4, "paper4", "Paper 4 (Chapters 17–20)")

# Render Basket Tab
with tab_cart:
    st.header("🛒 Selected Pages Basket")
    
    if not st.session_state.handout_basket:
        st.info("Your basket is empty. Search for topics in Papers 1–4 and click '➕ Add to Cart' to add pages here.")
    else:
        st.write(f"Total items in basket: **{len(st.session_state.handout_basket)}**")
        
        for b_idx, b_item in enumerate(st.session_state.handout_basket):
            with st.expander(f"Item {b_idx + 1}: {b_item['file']} (Page {b_item['page'] + 1})", expanded=False):
                col_img, col_act = st.columns([3, 1])
                with col_img:
                    img = render_pdf_page_preview(b_item["path"], b_item["page"])
                    if img:
                        st.image(img, use_container_width=True)
                with col_act:
                    if st.button("❌ Remove Item", key=f"remove_basket_{b_idx}"):
                        st.session_state.handout_basket.pop(b_idx)
                        st.rerun()

        st.markdown("---")
        
        docx_data = generate_docx_handout(st.session_state.handout_basket)
        st.download_button(
            label="🪄 Download Dynamic Word Document Handout (.docx)",
            data=docx_data,
            file_name="9618_Computer_Science_Handout.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# Render Admin / Upload Tab
with tab_admin:
    st.header("⚙️ Admin - Upload PDF Files to Google Drive")
    st.markdown("Select a PDF file from your computer and choose the destination folder on Google Drive.")
    
    col_target, col_up = st.columns([1, 2])
    
    with col_target:
        target_folder = st.selectbox(
            "Target Drive Folder",
            options=["paper1", "paper2", "paper3", "paper4", "other_notes"],
            format_func=lambda x: {
                "paper1": "Paper 1 (Ch 1-8)",
                "paper2": "Paper 2 (Ch 9-12)",
                "paper3": "Paper 3 (Ch 13-16)",
                "paper4": "Paper 4 (Ch 17-20)",
                "other_notes": "Other 9618 Notes"
            }[x]
        )
    
    with col_up:
        uploaded_pdf = st.file_uploader("Choose a PDF file to upload", type=["pdf"])
        
    if st.button("📤 Upload PDF to Google Drive"):
        if uploaded_pdf is not None:
            with st.spinner(f"Uploading `{uploaded_pdf.name}` to Google Drive..."):
                success, msg = upload_file_to_gdrive(uploaded_pdf, target_folder)
                if success:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)
        else:
            st.warning("Please select a PDF file first.")
