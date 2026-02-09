import streamlit as st
import sys
import os
import subprocess
import platform
import shutil
import time
import tempfile
import multiprocessing
import queue
import urllib.request
import zipfile
from PIL import Image

# ------------------------------------------------------------------
# CONFIGURATION (Must be first Streamlit command)
# ------------------------------------------------------------------
st.set_page_config(page_title="Ultra-Fast OCR", layout="wide")

# ------------------------------------------------------------------
# 1. SETUP & DEPENDENCIES
# ------------------------------------------------------------------

def install_python_deps():
    """Ensures required python packages are installed."""
    required = {'pytesseract': 'pytesseract', 'pdf2image': 'pdf2image'}
    missing = []
    for package, import_name in required.items():
        try:
            __import__(import_name)
        except ImportError:
            missing.append(package)
    
    if missing:
        st.warning(f"Installing missing Python packages: {', '.join(missing)}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])

# Run dependency check once
try:
    install_python_deps()
    import pytesseract
    from pdf2image import convert_from_path, pdfinfo_from_path
except ImportError:
    st.error("Failed to import dependencies. Please restart the app.")
    sys.exit(1)

def get_poppler_path_windows():
    """Returns the path to the local poppler bin folder if it exists."""
    install_dir = os.path.join(os.getcwd(), "poppler_bin")
    if os.path.exists(install_dir):
        for root, dirs, files in os.walk(install_dir):
            if "pdfinfo.exe" in files:
                return root
    return None

def download_and_install_poppler_windows():
    """Downloads portable Poppler for Windows."""
    poppler_url = "https://github.com/oschwartz10612/poppler-windows/releases/download/v24.02.0-0/Release-24.02.0-0.zip"
    existing_path = get_poppler_path_windows()
    if existing_path:
        return existing_path
            
    try:
        st.info("Downloading Poppler for Windows (this may take a moment)...")
        install_dir = os.path.join(os.getcwd(), "poppler_bin")
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as tmp_file:
            urllib.request.urlretrieve(poppler_url, tmp_file.name)
            zip_path = tmp_file.name
        
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(install_dir)
        
        os.remove(zip_path)
        return get_poppler_path_windows()

    except Exception as e:
        st.error(f"Failed to download/install Poppler: {e}")
        return None

def get_tesseract_cmd():
    if shutil.which("tesseract"): return "tesseract"
    if platform.system() == "Windows":
        potential_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        if os.path.exists(potential_path): return potential_path
    return None

def check_system_deps():
    """Checks system dependencies."""
    system_os = platform.system()
    tess_cmd = get_tesseract_cmd()
    
    # 1. Tesseract Check
    if not tess_cmd:
        if system_os == "Linux":
            # On Streamlit Cloud/Linux, we CANNOT use sudo/apt-get from python.
            # We must rely on packages.txt.
            st.error("❌ System Dependency Missing: Tesseract-OCR")
            st.markdown("""
            **Deployment Error:**
            The app cannot find `tesseract`. 
            
            1. Ensure `packages.txt` exists in your repo.
            2. It must contain exactly:
               ```
               tesseract-ocr
               poppler-utils
               libgl1
               ```
            3. Reboot the app in Streamlit Cloud dashboard.
            """)
            st.stop()
            return False
            
        elif system_os == "Darwin":
            st.error("Tesseract not found. Run: `brew install tesseract`")
            return False
        elif system_os == "Windows":
             st.warning("Tesseract not found. Attempting Winget install...")
             try:
                 subprocess.check_call(["winget", "install", "UB-Mannheim.TesseractOCR"])
                 st.success("Tesseract installed. Please restart the app.")
             except:
                 st.error("Automatic install failed. Please install Tesseract-OCR manually.")
             return False
    else:
        pytesseract.pytesseract.tesseract_cmd = tess_cmd

    # 2. Poppler Check
    poppler_path = None
    if system_os == "Windows":
        poppler_path = get_poppler_path_windows()
        if not poppler_path:
             poppler_path = download_and_install_poppler_windows()
        
        if poppler_path and poppler_path not in os.environ["PATH"]:
            os.environ["PATH"] += os.pathsep + poppler_path
    
    if not shutil.which("pdfinfo") and not poppler_path:
        st.error("❌ System Dependency Missing: Poppler (pdfinfo)")
        if system_os == "Linux":
            st.markdown("Ensure `poppler-utils` is listed in your `packages.txt` file.")
        st.stop()
        return False

    return True

# ------------------------------------------------------------------
# 2. WORKER LOGIC
# ------------------------------------------------------------------

def process_pdf_chunk(args):
    """
    Worker function to process a specific range of pages.
    """
    pdf_path, start_page, end_page, tess_cmd, poppler_path, batch_id, queue_obj = args
    
    if tess_cmd:
        pytesseract.pytesseract.tesseract_cmd = tess_cmd
        
    results = [] 
    
    try:
        # Notify UI: Start
        if queue_obj:
            queue_obj.put(('START', batch_id, f"Pages {start_page}-{end_page}"))

        images = convert_from_path(
            pdf_path, 
            first_page=start_page, 
            last_page=end_page, 
            poppler_path=poppler_path,
            thread_count=1 
        )
        
        total = len(images)
        for i, img in enumerate(images):
            text = pytesseract.image_to_string(img, lang='eng')
            results.append((start_page + i, text))
            
            # Notify UI: Progress
            if queue_obj:
                queue_obj.put(('PROGRESS', batch_id, i + 1, total))
        
        # Notify UI: Done
        if queue_obj:
            queue_obj.put(('DONE', batch_id))
            
        return (True, results)
    except Exception as e:
        if queue_obj:
            queue_obj.put(('DONE', batch_id))
        return (False, str(e))

# ------------------------------------------------------------------
# 3. STREAMLIT UI
# ------------------------------------------------------------------

def main():
    st.title("Ultra-Fast Multiprocess OCR")
    
    # Check dependencies on load
    if not check_system_deps():
        st.stop()
    
    cpu_count = multiprocessing.cpu_count()
    st.sidebar.info(f"Detected {cpu_count} CPU Cores")
    
    # --- RESULT HANDLING ---
    if "ocr_result" not in st.session_state:
        st.session_state.ocr_result = None
        st.session_state.result_filename = "ocr_output.txt"

    # --- IMMEDIATE DOWNLOAD BUTTON ---
    # We display this AT THE TOP if results exist
    if st.session_state.ocr_result:
        st.success("✅ Processing Complete! Download your file below:")
        st.download_button(
            label="⬇️ Download Extracted Text", 
            data=st.session_state.ocr_result, 
            file_name=st.session_state.result_filename,
            mime="text/plain",
            type="primary",
            key="download_top"
        )
        st.markdown("---")

    uploaded_file = st.file_uploader("Upload PDF or Image", type=['pdf', 'png', 'jpg', 'jpeg', 'tiff'])
    
    if uploaded_file:
        # Reset result if a new file is uploaded
        if "current_file" not in st.session_state or st.session_state.current_file != uploaded_file.name:
            st.session_state.ocr_result = None
            st.session_state.current_file = uploaded_file.name

        # --- PROCESS BUTTON ---
        if st.button(f"Start Processing (Uses {cpu_count} Cores)"):
            
            # Save uploaded file to temp path because pdf2image needs a real path
            with tempfile.NamedTemporaryFile(delete=False, suffix=f".{uploaded_file.name.split('.')[-1]}") as tmp:
                tmp.write(uploaded_file.getvalue())
                input_path = tmp.name

            # --- PROCESSING LOGIC ---
            st.write("### Processing Status")
            progress_container = st.container()
            status_container = st.empty()
            
            progress_bars = {}
            status_labels = {}
            
            try:
                file_ext = os.path.splitext(input_path)[1].lower()
                final_text = ""
                
                if file_ext == '.pdf':
                    # 1. Info
                    poppler_path = get_poppler_path_windows() if platform.system() == "Windows" else None
                    info = pdfinfo_from_path(input_path, poppler_path=poppler_path)
                    total_pages = info["Pages"]
                    
                    status_container.info(f"PDF has {total_pages} pages. Distributing tasks...")
                    
                    # 2. Prepare Tasks
                    chunk_size = 5
                    tasks = []
                    batch_counter = 0
                    
                    manager = multiprocessing.Manager()
                    m_queue = manager.Queue()
                    
                    tess_cmd = pytesseract.pytesseract.tesseract_cmd
                    
                    for start in range(1, total_pages + 1, chunk_size):
                        end = min(start + chunk_size - 1, total_pages)
                        batch_counter += 1
                        tasks.append((input_path, start, end, tess_cmd, poppler_path, batch_counter, m_queue))
                    
                    # 3. Execute & Monitor
                    pool = multiprocessing.Pool(processes=cpu_count)
                    async_results = []
                    
                    for task in tasks:
                        res = pool.apply_async(process_pdf_chunk, (task,))
                        async_results.append(res)
                    
                    pool.close() 
                    
                    # Polling Loop
                    active_tasks = len(tasks)
                    
                    with progress_container:
                        pass # Anchor

                    while active_tasks > 0:
                        try:
                            while True:
                                msg = m_queue.get_nowait()
                                type_, batch_id = msg[0], msg[1]
                                
                                if type_ == 'START':
                                    label_txt = msg[2]
                                    with progress_container:
                                        col1, col2 = st.columns([1, 3])
                                        status_labels[batch_id] = col1.empty()
                                        status_labels[batch_id].text(label_txt)
                                        progress_bars[batch_id] = col2.progress(0)
                                        
                                elif type_ == 'PROGRESS':
                                    curr, total = msg[2], msg[3]
                                    if batch_id in progress_bars:
                                        progress_bars[batch_id].progress(curr / total)
                                        
                                elif type_ == 'DONE':
                                    if batch_id in progress_bars:
                                        progress_bars[batch_id].empty()
                                        status_labels[batch_id].empty()
                                        del progress_bars[batch_id]
                                        del status_labels[batch_id]

                        except queue.Empty:
                            pass
                        
                        if all(r.ready() for r in async_results):
                            break
                        
                        time.sleep(0.1)

                    pool.join()
                    
                    # 4. Collect Results
                    extracted_data = []
                    for r in async_results:
                        success, data = r.get()
                        if success:
                            extracted_data.extend(data)
                        else:
                            st.error(f"Error in batch: {data}")
                            
                    extracted_data.sort(key=lambda x: x[0])
                    final_text = "\n".join([f"--- Page {p} ---\n{t}" for p, t in extracted_data])

                else:
                    # Image
                    with Image.open(input_path) as img:
                        final_text = pytesseract.image_to_string(img)

                # Store result in session state
                st.session_state.ocr_result = final_text
                st.session_state.result_filename = f"{os.path.splitext(uploaded_file.name)[0]}_ocr.txt"
                
                status_container.success("Processing Complete!")
                
            except Exception as e:
                st.error(f"An error occurred: {e}")
            finally:
                if os.path.exists(input_path):
                    os.unlink(input_path)
            
            # Force rerun to show the download button immediately at the top
            st.rerun()

if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
