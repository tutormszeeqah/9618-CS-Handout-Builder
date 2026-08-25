import os
import io
import fitz  # PyMuPDF
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ==========================================
# 1. PAGE CONFIGURATION & CONSTANTS
# ==========================================
st.set_page_config(
    page_title="9618 Computer Science Topical Portal",
    page_icon="💻",
    layout="wide"
)

# Apply Custom Color Theme & Styling
CUSTOM_CSS = """
<style>
    /* Main App Background & Primary Font Styling */
    .stApp {
        background-color: #f8f9fa;
    }
    
    /* Custom Sidebar Styling */
    [data-testid="stSidebar"] {
        background-color: #2c3e50;
        color: #ffffff;
    }
    [data-testid="stSidebar"] stMarkdown, [data-testid="stSidebar"] label, [data-testid="stSidebar"] span {
        color: #ecf0f1 !important;
    }
    
    /* Headers & Accent Colors */
    h1, h2, h3 {
        color: #2c3e50;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    /* Primary Buttons Styling */
    div.stButton > button[kind="primary"] {
        background-color: #3498db;
        color: white;
        border-radius: 8px;
        border: none;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    div.stButton > button[kind="primary"]:hover {
        background-color: #2980b9;
        box-shadow: 0 4px 8px rgba(0,0,0,0.15);
    }
    
    /* Custom Metric & Status Card Container */
    .status-card {
        background-color: #ffffff;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #3498db;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        margin-bottom: 10px;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

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

# Ensure local storage directories exist on server startup
for folder in LOCAL_FOLDERS.values():
    os.makedirs(folder, exist_ok=True)

# Initialize Session State Variables
if "handout_basket" not in st.session_state:
    st.session_state.handout_basket = []

for p_key in ["paper1", "paper2", "paper3", "paper4"]:
    if f"{p_key}_results" not in st.session_state:
        st.session_state[f"{p_key}_results"] = []


# ==========================================
# 2. GOOGLE DRIVE AUTHENTICATION & SYNC
# ==========================================
def get_gdrive_service():
    """Authenticates using Streamlit Service Account Secrets."""
    scopes = ["https://www.googleapis.com/auth/drive.readonly"]
    creds_dict = st.secrets["gcp_service_account"]
    credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return build("drive", "v3", credentials=credentials)

def sync_single_folder(service, folder_id: str, local_dir: str) -> tuple[int, str]:
    """Downloads missing PDF files from a Google Drive folder to local directory."""
    if not folder_id:
        return 0, f"No Folder ID found for directory {local_dir}"
    
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
    """Runs sync across all 5 configured Google Drive folders."""
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


# ==========================================
# 3. PDF RENDERING & SEARCH ENGINE LOGIC
# ==========================================
def render_pdf_page_preview(filepath: str, page_num: int) -> bytes:
    """Renders a PDF page to PNG image bytes for web preview and document export."""
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
    """
    Performs full-text search across BOTH the specified Paper folder 
    AND the shared Other9618Notes folder with robust chapter filtering.
    """
    results = []
    keywords = [k.strip().lower() for k in keyword_string.split(",") if k.strip()]
    
    target_folders = [LOCAL_FOLDERS[paper_key], LOCAL_FOLDERS["other_notes"]]

    for folder_path in target_folders:
        if not os.path.exists(folder_path):
            continue

        for file in os.listdir(folder_path):
            if not file.endswith(".pdf"):
                continue

            # Apply chapter filter ONLY if a specific chapter is selected AND scanning the main paper folder
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
                    # Extract text and normalize spaces/newlines
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
# 4. WORD DOCUMENT EXPORT GENERATOR
# ==========================================
def generate_docx_handout(basket_items: list[dict]) -> io.BytesIO:
    """Compiles selected PDF page snapshots into a dynamic Word Document handout."""
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
# 5. UI COMPONENTS & TAB RENDERER
# ==========================================
def render_paper_tab(tab_object, paper_key: str, paper_title: str):
    """Renders search controls, filters, and page previews for Paper tabs."""
    with tab_object:
        st.header(f"🔍 Search {paper_title}")
        
        col_kw, col_ch = st.columns([2, 1])
        with col_kw:
            kw = st.text_input("Enter Keyword", placeholder="e.g., virtual, binary, stack", key=f"{paper_key}_kw")
        with col_ch:
            chapter_options = ["All Chapters"] + PAPER_CHAPTER_MAPPING[paper_key]
            selected_ch = st.selectbox("Select Chapter Filter", options=chapter_options, key=f"{paper_key}_chap")

        if st.button(f"Search {paper_title}", type="primary", key=f"btn_{paper_key}"):
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
# 6. APPLICATION MAIN LAYOUT
# ==========================================
st.title("💻 Cambridge 9618 Computer Science Portal")
st.markdown("Search topical notes, preview pages, and compile dynamic `.docx` handouts.")

# Sidebar Layout
with st.sidebar:
    st.header("🔄 Google Drive Sync")
    if st.button("🔄 Sync Google Drive", type="primary", use_container_width=True):
        with st.spinner("Syncing Google Drive folders..."):
            count, msgs = perform_bulk_sync()
            st.success(f"🎉 Sync Complete! {count} new file(s) downloaded.")
            for m in msgs:
                st.caption(m)

    st.markdown("---")
    
    # Local Storage Status Card
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

# Navigation Tabs
tab1, tab2, tab3, tab4, tab_cart = st.tabs([
    "📘 Paper 1 (Ch 1–8)",
    "📗 Paper 2 (Ch 9–12)",
    "📙 Paper 3 (Ch 13–16)",
    "📕 Paper 4 (Ch 17–20)",
    "🛒 Basket / Cart"
])

# Render Search Tabs
render_paper_tab(tab1, "paper1", "Paper 1 (Chapters 1–8)")
render_paper_tab(tab2, "paper2", "Paper 2 (Chapters 9–12)")
render_paper_tab(tab3, "paper3", "Paper 3 (Chapters 13–16)")
render_paper_tab(tab4, "paper4", "Paper 4 (Chapters 17–20)")

# Render Cart Tab
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
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary"
        )
