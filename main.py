import sys
import os
import subprocess
import platform
import shutil
import threading
import tkinter as tk
from tkinter import ttk
import urllib.request
import zipfile
import tempfile
import multiprocessing
import queue  # Import standard queue for Empty exception
from tkinter import filedialog, messagebox, scrolledtext

# ------------------------------------------------------------------
# DEPENDENCY CHECK & AUTO-INSTALL
# ------------------------------------------------------------------
def install_python_deps():
    required = {'pytesseract': 'pytesseract', 'pillow': 'PIL', 'pdf2image': 'pdf2image'}
    missing = []
    for package, import_name in required.items():
        try:
            __import__(import_name)
        except ImportError:
            missing.append(package)
    
    if missing:
        print(f"Installing missing Python packages: {', '.join(missing)}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])

try:
    install_python_deps()
    import pytesseract
    from PIL import Image
    from pdf2image import convert_from_path, pdfinfo_from_path
    from pdf2image.exceptions import PDFInfoNotInstalledError
except Exception as e:
    messagebox.showerror("Dependency Error", f"Failed to install/import dependencies: {e}")
    sys.exit(1)

# ------------------------------------------------------------------
# SYSTEM CONFIGURATION & HELPERS
# ------------------------------------------------------------------

def get_poppler_path_windows():
    """Returns the path to the local poppler bin folder if it exists."""
    install_dir = os.path.join(os.getcwd(), "poppler_bin")
    if os.path.exists(install_dir):
        # Search for bin folder containing pdfinfo.exe
        for root, dirs, files in os.walk(install_dir):
            if "pdfinfo.exe" in files:
                return root
    return None

def download_and_install_poppler_windows(log_callback=print):
    """Downloads portable Poppler for Windows."""
    poppler_url = "https://github.com/oschwartz10612/poppler-windows/releases/download/v24.02.0-0/Release-24.02.0-0.zip"
    
    # Check if already installed
    existing_path = get_poppler_path_windows()
    if existing_path:
        os.environ["PATH"] += os.pathsep + existing_path
        log_callback(f"[Auto-Setup] Found existing Poppler: {existing_path}")
        return True
            
    try:
        log_callback("[Auto-Setup] Downloading Poppler for Windows (this may take a moment)...")
        install_dir = os.path.join(os.getcwd(), "poppler_bin")
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as tmp_file:
            urllib.request.urlretrieve(poppler_url, tmp_file.name)
            zip_path = tmp_file.name
        
        log_callback("[Auto-Setup] Extracting Poppler...")
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(install_dir)
        
        os.remove(zip_path)
        
        # Add to path
        bin_path = get_poppler_path_windows()
        if bin_path:
            os.environ["PATH"] += os.pathsep + bin_path
            log_callback(f"[Auto-Setup] Poppler installed: {bin_path}")
            return True
        else:
            log_callback("[Auto-Setup] Downloaded Poppler but could not find 'bin' folder.")
            return False

    except Exception as e:
        log_callback(f"[Auto-Setup] Failed to download/install Poppler: {e}")
        return False

def install_system_deps(log_callback=print):
    """Attempts to install Tesseract and Poppler."""
    system_os = platform.system()
    deps_missing = []
    
    if not shutil.which("tesseract"):
        deps_missing.append("tesseract")
    if not shutil.which("pdfinfo") and (system_os != "Windows" or not get_poppler_path_windows()):
        deps_missing.append("poppler")

    if not deps_missing:
        return True

    log_callback(f"[Auto-Setup] Missing: {', '.join(deps_missing)}. Installing...")

    try:
        if system_os == "Linux":
            if shutil.which("apt-get"):
                subprocess.check_call(["sudo", "apt-get", "update"])
                if "tesseract" in deps_missing:
                    subprocess.check_call(["sudo", "apt-get", "install", "-y", "tesseract-ocr"])
                if "poppler" in deps_missing:
                    subprocess.check_call(["sudo", "apt-get", "install", "-y", "poppler-utils"])
                return True

        elif system_os == "Darwin":  # macOS
            if shutil.which("brew"):
                if "tesseract" in deps_missing:
                    subprocess.check_call(["brew", "install", "tesseract"])
                if "poppler" in deps_missing:
                    subprocess.check_call(["brew", "install", "poppler"])
                return True

        elif system_os == "Windows":
            if "tesseract" in deps_missing and shutil.which("winget"):
                log_callback("Installing Tesseract via Winget...")
                subprocess.check_call(["winget", "install", "UB-Mannheim.TesseractOCR"])
            
            if "poppler" in deps_missing:
                download_and_install_poppler_windows(log_callback)
            
            return True
        
    except Exception as e:
        log_callback(f"Auto-install failed: {e}")
        return False

    return False

def get_tesseract_cmd():
    if shutil.which("tesseract"): return "tesseract"
    if platform.system() == "Windows":
        potential_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        if os.path.exists(potential_path): return potential_path
    return None

# ------------------------------------------------------------------
# MULTIPROCESSING WORKER
# ------------------------------------------------------------------
# Must be at top level for pickle serialization on Windows
def process_pdf_chunk(args):
    """
    Worker function to process a specific range of pages from a PDF.
    Sends updates to a queue for dynamic UI.
    """
    pdf_path, start_page, end_page, tess_cmd, poppler_path, batch_id, queue = args
    
    # Setup Tesseract inside the worker process
    if tess_cmd:
        pytesseract.pytesseract.tesseract_cmd = tess_cmd
        
    results = [] 
    
    try:
        # Notify UI: Worker Started
        if queue:
            queue.put(('START', batch_id, f"Pages {start_page}-{end_page}"))

        # Convert only the specific chunk of pages to images
        images = convert_from_path(
            pdf_path, 
            first_page=start_page, 
            last_page=end_page, 
            poppler_path=poppler_path,
            thread_count=1 
        )
        
        total = len(images)
        for i, img in enumerate(images):
            # Perform OCR
            text = pytesseract.image_to_string(img, lang='eng')
            results.append((start_page + i, text))
            
            # Notify UI: Progress Update
            if queue:
                queue.put(('PROGRESS', batch_id, i + 1, total))
        
        # Notify UI: Worker Done
        if queue:
            queue.put(('DONE', batch_id))
            
        return (True, results)
    except Exception as e:
        if queue:
            queue.put(('DONE', batch_id)) # Ensure bar is removed on error
        return (False, str(e))

# ------------------------------------------------------------------
# GUI APPLICATION
# ------------------------------------------------------------------

class OCRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Ultra-Fast OCR (Multiprocessing)")
        self.root.geometry("700x750")
        
        # Configure Tesseract
        tess_cmd = get_tesseract_cmd()
        if tess_cmd:
            pytesseract.pytesseract.tesseract_cmd = tess_cmd
        
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.cpu_count = multiprocessing.cpu_count()
        
        # State for dynamic bars
        self.active_bars = {}  # batch_id -> {frame, progress, label}
        self.current_queue = None
        self.is_processing = False

        self._build_ui()
        
        # Check deps logic
        path_check = shutil.which("pdfinfo") or get_poppler_path_windows()
        if not tess_cmd or not path_check:
            self.log("Dependencies missing. Auto-installing...")
            threading.Thread(target=self._run_install_deps).start()
        else:
            self.log(f"System ready. Detected {self.cpu_count} CPU cores.")

    def _build_ui(self):
        # Input Section
        frame_input = tk.LabelFrame(self.root, text="Input File", padx=10, pady=10)
        frame_input.pack(fill="x", padx=10, pady=5)
        tk.Entry(frame_input, textvariable=self.input_path).pack(side="left", fill="x", expand=True)
        tk.Button(frame_input, text="Browse...", command=self.browse_input).pack(side="left", padx=5)

        # Output Section
        frame_output = tk.LabelFrame(self.root, text="Output File", padx=10, pady=10)
        frame_output.pack(fill="x", padx=10, pady=5)
        tk.Entry(frame_output, textvariable=self.output_path).pack(side="left", fill="x", expand=True)
        tk.Button(frame_output, text="Browse...", command=self.browse_output).pack(side="left", padx=5)

        # Action Section
        self.btn_convert = tk.Button(self.root, text=f"Start Multiprocess OCR (Uses {self.cpu_count} Cores)", 
                                   command=self.start_conversion_thread, 
                                   bg="#2196F3", fg="white", font=("Arial", 11, "bold"))
        self.btn_convert.pack(pady=10)

        # Active Workers Section (Dynamic Bars)
        lbl_workers = tk.Label(self.root, text="Active Workers:", font=("Arial", 9, "bold"))
        lbl_workers.pack(anchor="w", padx=10)
        
        # Frame to hold dynamic progress bars
        self.bars_frame = tk.Frame(self.root)
        self.bars_frame.pack(fill="x", padx=10, pady=5)

        # Log Section
        tk.Label(self.root, text="Processing Log:").pack(anchor="w", padx=10)
        self.log_area = scrolledtext.ScrolledText(self.root, height=15)
        self.log_area.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def log(self, message):
        def _insert():
            self.log_area.insert(tk.END, message + "\n")
            self.log_area.see(tk.END)
        self.root.after(0, _insert)

    def browse_input(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF & Images", "*.pdf *.png *.jpg *.jpeg *.tiff")])
        if filename: self.input_path.set(filename)

    def browse_output(self):
        filename = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
        if filename: self.output_path.set(filename)

    def _run_install_deps(self):
        install_system_deps(log_callback=self.log)
        win_poppler = get_poppler_path_windows()
        if win_poppler and win_poppler not in os.environ["PATH"]:
             os.environ["PATH"] += os.pathsep + win_poppler
        
        tess = get_tesseract_cmd()
        if tess: pytesseract.pytesseract.tesseract_cmd = tess
        self.log("Dependency check finished.")

    def process_queue_updates(self):
        """
        Runs periodically on main thread to check for messages from workers
        and update the UI progress bars accordingly.
        """
        if not self.current_queue:
            return

        try:
            # Process all available messages
            while True:
                msg = self.current_queue.get_nowait()
                msg_type = msg[0]
                batch_id = msg[1]

                if msg_type == 'START':
                    # Create new progress row
                    label_text = msg[2]
                    
                    row_frame = tk.Frame(self.bars_frame, pady=2)
                    row_frame.pack(fill="x", expand=True)
                    
                    lbl = tk.Label(row_frame, text=label_text, width=20, anchor="w", font=("Consolas", 8))
                    lbl.pack(side="left")
                    
                    pb = ttk.Progressbar(row_frame, orient="horizontal", length=200, mode="determinate")
                    pb.pack(side="left", fill="x", expand=True, padx=5)
                    
                    count_lbl = tk.Label(row_frame, text="0%", width=8, anchor="e", font=("Consolas", 8))
                    count_lbl.pack(side="left")

                    self.active_bars[batch_id] = {
                        "frame": row_frame,
                        "pb": pb,
                        "lbl": count_lbl
                    }

                elif msg_type == 'PROGRESS':
                    # Update existing row
                    current, total = msg[2], msg[3]
                    if batch_id in self.active_bars:
                        bar_data = self.active_bars[batch_id]
                        bar_data["pb"]["maximum"] = total
                        bar_data["pb"]["value"] = current
                        bar_data["lbl"].config(text=f"{current}/{total}")

                elif msg_type == 'DONE':
                    # Remove row
                    if batch_id in self.active_bars:
                        bar_data = self.active_bars[batch_id]
                        bar_data["frame"].destroy()
                        del self.active_bars[batch_id]

        except queue.Empty:
            pass
        finally:
            if self.is_processing:
                self.root.after(100, self.process_queue_updates)

    def start_conversion_thread(self):
        input_file = self.input_path.get()
        output_file = self.output_path.get()

        if not input_file or not output_file:
            messagebox.showwarning("Error", "Please select input and output files.")
            return

        self.btn_convert.config(state="disabled", text="Processing...")
        self.log_area.delete('1.0', tk.END)
        self.is_processing = True
        
        # Clear any left-over bars
        for widget in self.bars_frame.winfo_children():
            widget.destroy()
        self.active_bars.clear()

        threading.Thread(target=self.perform_ocr_multiprocess, args=(input_file, output_file)).start()

    def perform_ocr_multiprocess(self, input_path, output_path):
        # Use Manager for shared queue
        with multiprocessing.Manager() as manager:
            self.current_queue = manager.Queue()
            
            # Start monitoring queue in main thread
            self.root.after(50, self.process_queue_updates)

            try:
                self.log(f"Analyzing {input_path}...")
                file_ext = os.path.splitext(input_path)[1].lower()
                
                extracted_data = []

                if file_ext == '.pdf':
                    # 1. Get Page Count
                    poppler_path = get_poppler_path_windows() if platform.system() == "Windows" else None
                    try:
                        info = pdfinfo_from_path(input_path, poppler_path=poppler_path)
                        total_pages = info["Pages"]
                    except Exception as e:
                        self.log("Error reading PDF. Is Poppler installed?")
                        self.log(str(e))
                        return

                    self.log(f"PDF has {total_pages} pages.")
                    
                    # 2. Create Chunks (Small batches for responsive UI)
                    chunk_size = 5
                    tasks = []
                    
                    batch_counter = 0
                    for start in range(1, total_pages + 1, chunk_size):
                        end = min(start + chunk_size - 1, total_pages)
                        tess_cmd = pytesseract.pytesseract.tesseract_cmd
                        batch_counter += 1
                        # Pass batch_id and queue to worker
                        tasks.append((input_path, start, end, tess_cmd, poppler_path, batch_counter, self.current_queue))

                    self.log(f"Queueing {len(tasks)} batches on {self.cpu_count} cores...")

                    # 3. Execute in Parallel
                    with multiprocessing.Pool(processes=self.cpu_count) as pool:
                        for result in pool.imap_unordered(process_pdf_chunk, tasks):
                            success, data = result
                            if success:
                                extracted_data.extend(data)
                                pages_done = [p for p, t in data]
                                self.log(f"Completed Pages {min(pages_done)}-{max(pages_done)}")
                            else:
                                self.log(f"Error in batch: {data}")

                    # 4. Sort results
                    extracted_data.sort(key=lambda x: x[0])
                    final_text = "\n".join([f"--- Page {p} ---\n{t}" for p, t in extracted_data])

                else:
                    self.log("Processing single image...")
                    # Simulating batch behavior for single image for consistency
                    self.current_queue.put(('START', 1, "Image Processing"))
                    with Image.open(input_path) as img:
                        final_text = pytesseract.image_to_string(img)
                    self.current_queue.put(('PROGRESS', 1, 1, 1))
                    self.current_queue.put(('DONE', 1))

                # Save
                with open(output_path, "w", encoding="utf-8") as f:
                    f.write(final_text)
                
                self.log("-" * 30)
                self.log("COMPLETED SUCCESSFULLY.")
                messagebox.showinfo("Success", "OCR Complete")

            except Exception as e:
                self.log(f"CRITICAL ERROR: {e}")
                messagebox.showerror("Error", str(e))
            finally:
                self.is_processing = False
                self.current_queue = None
                
                # Update button in main thread
                self.root.after(0, lambda: self.btn_convert.config(state="normal", text=f"Start Multiprocess OCR (Uses {self.cpu_count} Cores)"))

if __name__ == "__main__":
    multiprocessing.freeze_support()
    root = tk.Tk()
    app = OCRApp(root)
    root.mainloop()