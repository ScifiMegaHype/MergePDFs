import os, sys
import ctypes
from ctypes import wintypes
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PyPDF2 import PdfMerger, PdfReader

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def isMergePDFsRunning(mutex_name):
    # Create a named mutex
    kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
    CreateMutex = kernel32.CreateMutexW
    CreateMutex.argtypes = (wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR)
    CreateMutex.restype = wintypes.HANDLE

    _ = CreateMutex(None, False, mutex_name) # keeps mutex so Windows doesn't destroy the mutex while the app runs

    ERROR_ALREADY_EXISTS = 183
    last_error = ctypes.get_last_error()

    if last_error == ERROR_ALREADY_EXISTS:
        return True
    
    return False
    
def combine_pdfs(pdf_list, output_path):
    with PdfMerger() as merger:
        for _, pdf_path in enumerate(pdf_list):
            merger.append(pdf_path, pages=None) # will allow for page selection in future
        merger.write(output_path)

def get_pdf_page_count(path):
    """Return number of pages in PDF, or -1 on error."""
    try:
        reader = PdfReader(path)
        return len(reader.pages)
    except Exception:
        return -1

class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Merger")
        self.root.iconbitmap(resource_path(r"adobe.ico"))
        self.root.geometry("600x520")
        self.WIDTH = 600
        self.HEIGHT = 520
        self.center_window(self.WIDTH, self.HEIGHT)
        self.root.resizable(False, False)

        # Main frame for file list and move buttons
        self.frame_main = tk.LabelFrame(root, text="Files to Merge", padx=10, pady=10)
        self.frame_main.pack(fill="both", expand=True, padx=10, pady=0)

        # Frame for the left move buttons
        self.frame_right = tk.Frame(self.frame_main)
        self.frame_right.pack(side=tk.RIGHT, padx=(0, 10), fill=tk.Y)

        self.btn_up = tk.Button(self.frame_right, text="↑", width=4, command=self.move_up)
        self.btn_up.pack(anchor=tk.S, pady=5, expand=True)

        self.btn_down = tk.Button(self.frame_right, text="↓", width=4, command=self.move_down)
        self.btn_down.pack(anchor=tk.N, pady=5, expand=True)

        # Frame for the listbox and scrollbar
        self.frame_list = tk.Frame(self.frame_main)
        self.frame_list.pack(side=tk.LEFT, fill="both", expand=True)

        self.scrollbar = tk.Scrollbar(self.frame_list)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox = tk.Listbox(self.frame_list, selectmode=tk.SINGLE, yscrollcommand=self.scrollbar.set)
        self.listbox.pack(fill="both", expand=True)
        self.scrollbar.config(command=self.listbox.yview)

        # Right-click delete menu
        self.menu = tk.Menu(self.root, tearoff=0)
        self.menu.add_command(label="Delete", command=self.delete_selected)
        self.listbox.bind("<Button-3>", self.show_context_menu)

        # Bottom buttons (Add & Combine)
        frame_buttons = tk.Frame(root)
        frame_buttons.pack(pady=10)

        self.btn_add = tk.Button(frame_buttons, text="Add", width=12, command=self.add_file)
        self.btn_add.pack(side=tk.LEFT, padx=5)

        self.btn_combine = tk.Button(frame_buttons, text="Combine", width=12, command=self.combine_files)
        self.btn_combine.pack(side=tk.LEFT, padx=5)

        # Status bar
        self.status = tk.StringVar(value="Ready")
        ttk.Label(self.root, textvariable=self.status, relief=tk.SUNKEN, anchor="w").pack(side=tk.BOTTOM, fill="x")

    def center_window(self, w, h):
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def add_file(self):
        file_paths = filedialog.askopenfilenames(title="Select PDF File", filetypes=[("PDF files", "*.pdf")])
        if file_paths:
            self.status.set("")
            for file_path in file_paths:
                if file_path not in self.listbox.get(0, tk.END):
                    count = get_pdf_page_count(file_path)
                    display = f"{file_path} ({count} page{'s' if count > 1 else ''})" if count >= 0 else f"{file_path} (unknown # of pages)"
                    self.listbox.insert(tk.END, display)
                    self.status.set(self.status.get() + f"Added: {os.path.basename(file_path)}\n")
                else:
                    messagebox.showinfo("Duplicate", "This file is already added.")

    def delete_selected(self):
        selected = self.listbox.curselection()
        idx = selected[0]
        file_name = self.listbox.get(idx)
        if selected:
            self.status.set(f"Removed: {os.path.basename(file_name.split(' (')[0])}")
            self.listbox.delete(selected)

    def move_up(self):
        selected = self.listbox.curselection()
        if selected and selected[0] > 0:
            index = selected[0]
            text = self.listbox.get(index)
            self.listbox.delete(index)
            self.listbox.insert(index - 1, text)
            self.listbox.selection_set(index - 1)

    def move_down(self):
        selected = self.listbox.curselection()
        if selected and selected[0] < self.listbox.size() - 1:
            index = selected[0]
            text = self.listbox.get(index)
            self.listbox.delete(index)
            self.listbox.insert(index + 1, text)
            self.listbox.selection_set(index + 1)

    def show_context_menu(self, event):
        try:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(self.listbox.nearest(event.y))
            self.menu.post(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()

    def combine_files(self):
        files = [f.split(' (')[0] for f in list(self.listbox.get(0, tk.END))]
        if len(files) < 2:
            messagebox.showwarning("Add File", "Please add at least two PDF to merge.")
            return

        output_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                   filetypes=[("PDF files", "*.pdf")],
                                                   title="Save Merged PDF")
        if not output_path:
            return

        try:
            combine_pdfs(files, output_path)
            messagebox.showinfo("Success", f"Saved merged PDF as:\n  - {output_path}")
            self.status.set(f"Saved merged PDF as: \n  - {output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to combine PDFs:\n{e}")

if __name__ == "__main__":
    if isMergePDFsRunning("MergePDFs_Mutex_001"):    
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Already Running", "PDF Merger is already open.", icon="info")
        sys.exit(0)
    
    root = tk.Tk()
    app = PDFMergerApp(root)
    root.mainloop()
