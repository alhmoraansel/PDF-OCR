import streamlit_app as st
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
# 1. SETUP & DEPENDENCIES (Ported from original)
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
    """Checks and configures system dependencies."""
    system_os = platform.system()
    
    # 1. Tesseract
    tess_cmd = get_tesseract_cmd()
    if not tess_cmd:
        if system_os == "Linux":
            st.error("Tesseract not found. Run: `sudo apt-get install tesseract-ocr`")
        elif system_os == "Darwin":
            st.error("Tesseract not found. Run: `brew install tesseract`")
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

    # 2. Poppler
    poppler_path = None
    if system_os == "Windows":
        poppler_path = get_poppler_path_windows()
        if not poppler_path:
             poppler_path = download_and_install_poppler_windows()
        
        if poppler_path and poppler_path not in os.environ["PATH"]:
            os.environ["PATH"] += os.pathsep + poppler_path
    
    if not shutil.which("pdfinfo") and not poppler_path:
        st.error("Poppler (pdfinfo) not found. PDF processing will fail.")
        return False

    return True

# ------------------------------------------------------------------
# 2. WORKER LOGIC (Must be at module level for multiprocessing)
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
    st.set_page_config(page_title="Ultra-Fast OCR", layout="wide")
    
    st.title("Ultra-Fast Multiprocess OCR")
    
    # Check dependencies on load
    if not check_system_deps():
        st.stop()
    
    cpu_count = multiprocessing.cpu_count()
    st.sidebar.info(f"Detected {cpu_count} CPU Cores")
    
    uploaded_file = st.file_uploader("Upload PDF or Image", type=['pdf', 'png', 'jpg', 'jpeg', 'tiff'])
    
    if uploaded_file:
        # Save uploaded file to temp path because pdf2image needs a real path
        with tempfile.NamedTemporaryFile(delete=False, suffix=f".{uploaded_file.name.split('.')[-1]}") as tmp:
            tmp.write(uploaded_file.getvalue())
            input_path = tmp.name

        if st.button(f"Start Processing (Uses {cpu_count} Cores)"):
            
            # --- PROCESSING LOGIC ---
            st.write("### Processing Status")
            log_container = st.empty()
            progress_container = st.container()
            
            # Dictionary to track dynamic progress bars
            # Key: batch_id, Value: st.progress object
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
                    
                    log_container.info(f"PDF has {total_pages} pages. Distributing tasks...")
                    
                    # 2. Prepare Tasks
                    chunk_size = 5
                    tasks = []
                    batch_counter = 0
                    
                    # Use a Manager Queue for cross-process communication
                    manager = multiprocessing.Manager()
                    m_queue = manager.Queue()
                    
                    tess_cmd = pytesseract.pytesseract.tesseract_cmd
                    
                    for start in range(1, total_pages + 1, chunk_size):
                        end = min(start + chunk_size - 1, total_pages)
                        batch_counter += 1
                        tasks.append((input_path, start, end, tess_cmd, poppler_path, batch_counter, m_queue))
                    
                    # 3. Execute & Monitor
                    # We use apply_async so we can poll the queue in the main thread
                    pool = multiprocessing.Pool(processes=cpu_count)
                    async_results = []
                    
                    for task in tasks:
                        res = pool.apply_async(process_pdf_chunk, (task,))
                        async_results.append(res)
                    
                    pool.close() # No more tasks will be added
                    
                    # Polling Loop
                    active_tasks = len(tasks)
                    
                    # Container for dynamic bars
                    with progress_container:
                        # We create a grid or list of placeholders
                        # Since we don't know exact order, we create placeholders dynamically
                        pass

                    while active_tasks > 0:
                        try:
                            # Drain queue
                            while True:
                                msg = m_queue.get_nowait()
                                type_, batch_id = msg[0], msg[1]
                                
                                if type_ == 'START':
                                    label_txt = msg[2]
                                    with progress_container:
                                        # Create new UI elements for this batch
                                        col1, col2 = st.columns([1, 3])
                                        status_labels[batch_id] = col1.empty()
                                        status_labels[batch_id].text(label_txt)
                                        progress_bars[batch_id] = col2.progress(0)
                                        
                                elif type_ == 'PROGRESS':
                                    curr, total = msg[2], msg[3]
                                    if batch_id in progress_bars:
                                        progress_bars[batch_id].progress(curr / total)
                                        
                                elif type_ == 'DONE':
                                    # In Streamlit, we can't easily "destroy" widgets, 
                                    # but we can empty them.
                                    if batch_id in progress_bars:
                                        progress_bars[batch_id].empty()
                                        status_labels[batch_id].empty()
                                        del progress_bars[batch_id]
                                        del status_labels[batch_id]

                        except queue.Empty:
                            pass
                        
                        # Check if processes are actually done
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

                st.success("Processing Complete!")
                st.download_button("Download Extracted Text", final_text, file_name="ocr_output.txt")
                
            except Exception as e:
                st.error(f"An error occurred: {e}")
            finally:
                # Cleanup temp file
                if os.path.exists(input_path):
                    os.unlink(input_path)

if __name__ == "__main__":
    # Required for Windows multiprocessing
    multiprocessing.freeze_support()
    main()