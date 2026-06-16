import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from pypdf import PdfReader, PdfWriter

class PDFPageCutter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Page Cutter")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        self.input_file = None
        self.reader = None
        self.total_pages = 0
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="PDF Page Cutter", font=("Arial", 14, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=10)
        
        # File selection frame
        file_frame = ttk.LabelFrame(main_frame, text="Step 1: Select PDF File", padding="10")
        file_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        self.file_label = ttk.Label(file_frame, text="No file selected", foreground="gray")
        self.file_label.grid(row=0, column=0, sticky=tk.W, pady=5)
        
        browse_btn = ttk.Button(file_frame, text="Browse PDF", command=self.load_pdf)
        browse_btn.grid(row=0, column=1, sticky=tk.E, padx=5)
        
        # Page info frame
        info_frame = ttk.LabelFrame(main_frame, text="Step 2: Select Pages to Extract", padding="10")
        info_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        self.page_count_label = ttk.Label(info_frame, text="Total Pages: 0", font=("Arial", 11))
        self.page_count_label.grid(row=0, column=0, sticky=tk.W, pady=5)
        
        # Page selection method frame
        method_frame = ttk.LabelFrame(info_frame, text="Selection Method", padding="10")
        method_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        self.selection_var = tk.StringVar(value="range")
        
        ttk.Radiobutton(method_frame, text="Page Range (e.g., 1,3,5-10)", 
                       variable=self.selection_var, value="range", 
                       command=self.update_selection_method).grid(row=0, column=0, sticky=tk.W)
        
        ttk.Radiobutton(method_frame, text="Select Individual Pages", 
                       variable=self.selection_var, value="individual", 
                       command=self.update_selection_method).grid(row=1, column=0, sticky=tk.W)
        
        # Range input frame
        self.range_frame = ttk.Frame(info_frame)
        self.range_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        ttk.Label(self.range_frame, text="Enter pages:").pack(side=tk.LEFT)
        self.range_entry = ttk.Entry(self.range_frame, width=40)
        self.range_entry.pack(side=tk.LEFT, padx=5)
        self.range_entry.insert(0, "1-5")
        
        # Individual pages frame
        self.pages_frame = ttk.Frame(info_frame)
        
        # Scrollable frame for checkboxes
        canvas = tk.Canvas(self.pages_frame)
        scrollbar = ttk.Scrollbar(self.pages_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Extract button frame
        extract_frame = ttk.Frame(main_frame)
        extract_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        extract_btn = ttk.Button(extract_frame, text="Extract Pages", command=self.extract_pages)
        extract_btn.pack(side=tk.LEFT, padx=5)
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Ready", foreground="blue")
        self.status_label.grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=10)
        
        # Configure grid weights
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
    
    def load_pdf(self):
        file_path = filedialog.askopenfilename(
            title="Select PDF file",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                self.input_file = file_path
                self.reader = PdfReader(file_path)
                self.total_pages = len(self.reader.pages)
                
                # Update labels
                file_name = os.path.basename(file_path)
                self.file_label.config(text=file_name, foreground="black")
                self.page_count_label.config(text=f"Total Pages: {self.total_pages}")
                
                # Create checkboxes for pages
                self.page_vars = []
                for widget in self.scrollable_frame.winfo_children():
                    widget.destroy()
                
                for i in range(1, self.total_pages + 1):
                    var = tk.BooleanVar(value=False)
                    self.page_vars.append(var)
                    check = ttk.Checkbutton(
                        self.scrollable_frame,
                        text=f"Page {i}",
                        variable=var
                    )
                    check.pack(anchor=tk.W)
                
                # Update range entry default
                self.range_entry.delete(0, tk.END)
                self.range_entry.insert(0, f"1-{self.total_pages}")
                
                self.status_label.config(text="PDF loaded successfully!", foreground="green")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load PDF: {str(e)}")
                self.status_label.config(text="Error loading PDF", foreground="red")
    
    def update_selection_method(self):
        if self.selection_var.get() == "range":
            self.range_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
            self.pages_frame.grid_remove()
        else:
            self.range_frame.grid_remove()
            self.pages_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
    
    def parse_page_range(self, range_str):
        """Parse page range string like '1,3,5-10' into list of page numbers"""
        pages = []
        try:
            parts = range_str.replace(" ", "").split(",")
            for part in parts:
                if "-" in part:
                    start, end = part.split("-")
                    start, end = int(start.strip()), int(end.strip())
                    pages.extend(range(start, end + 1))
                else:
                    pages.append(int(part.strip()))
            
            # Validate pages
            pages = [p for p in pages if 1 <= p <= self.total_pages]
            return sorted(list(set(pages)))  # Remove duplicates and sort
        except:
            raise ValueError("Invalid page range format. Use format like: 1,3,5-10")
    
    def extract_pages(self):
        if not self.input_file:
            messagebox.showwarning("Warning", "Please load a PDF file first!")
            return
        
        try:
            # Determine which pages to extract
            if self.selection_var.get() == "range":
                range_str = self.range_entry.get()
                selected_pages = self.parse_page_range(range_str)
            else:
                selected_pages = [i + 1 for i, var in enumerate(self.page_vars) if var.get()]
            
            if not selected_pages:
                messagebox.showwarning("Warning", "Please select at least one page!")
                return
            
            # Create output PDF
            writer = PdfWriter()
            for page_num in selected_pages:
                writer.add_page(self.reader.pages[page_num - 1])
            
            # Save file dialog
            output_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
                initialfile=f"extracted_pages.pdf"
            )
            
            if output_path:
                with open(output_path, "wb") as output_file:
                    writer.write(output_file)
                
                self.status_label.config(
                    text=f"Successfully extracted {len(selected_pages)} page(s) to {os.path.basename(output_path)}",
                    foreground="green"
                )
                messagebox.showinfo("Success", 
                    f"Successfully extracted {len(selected_pages)} page(s)!\nSaved to: {output_path}")
        
        except ValueError as e:
            messagebox.showerror("Error", str(e))
            self.status_label.config(text="Error: Invalid input", foreground="red")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract pages: {str(e)}")
            self.status_label.config(text="Error extracting pages", foreground="red")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFPageCutter(root)
    root.mainloop()
