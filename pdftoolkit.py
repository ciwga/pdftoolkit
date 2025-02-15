import os
import sys
import json
import tkinter as tk
from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import filedialog, messagebox, ttk


class PDFToolkitApp:
    def __init__(self, root: TkinterDnD.Tk, icon_path: Path) -> None:
        self.root = root
        self.root.title("PDF Toolkit")
        self.root.geometry("508x680")
        self.root.iconbitmap(icon_path)
        self.root.resizable(False, False)
        
        self.style = ttk.Style()
        self.style.configure('TFrame', background='#f0f0f0')
        self.style.configure('Header.TLabel', font=('Arial', 10, 'bold'), background='#f0f0f0')
        
        self.create_widgets()
        self.setup_drag_and_drop()

        self.center_window()

    def center_window(self) -> None:
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'+{x}+{y}')

    def create_widgets(self) -> None:
        main_frame = ttk.Frame(self.root)
        main_frame.pack(padx=20, pady=15, fill='both', expand=True)

        # File Selection
        file_frame = ttk.LabelFrame(main_frame, text=" PDF File Operations ", padding=10)
        file_frame.grid(row=0, column=0, columnspan=2, sticky='ew', pady=5)
        
        self.file_path = tk.StringVar()
        self.file_entry = ttk.Entry(file_frame, textvariable=self.file_path, width=60)
        self.file_entry.grid(row=0, column=0, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.load_pdf, width=8).grid(row=0, column=1, padx=5)

        # Metadata Section
        meta_frame = ttk.LabelFrame(main_frame, text=" Metadata Editor ", padding=10)
        meta_frame.grid(row=1, column=0, sticky='nsew', pady=10)

        self.metadata_fields = {
            '/Title': tk.StringVar(),
            '/Author': tk.StringVar(),
            '/Subject': tk.StringVar(),
            '/Keywords': tk.StringVar(),
            '/Creator': tk.StringVar(),
            '/Producer': tk.StringVar(),
            '/CreationDate': tk.StringVar(),
            '/ModDate': tk.StringVar(),
        }

        for i, (key, var) in enumerate(self.metadata_fields.items()):
            ttk.Label(meta_frame, text=f"{key}:", width=13, anchor='e').grid(row=i, column=0, padx=5, pady=2)
            ttk.Entry(meta_frame, textvariable=var, width=40).grid(row=i, column=1, pady=2)
        
        # Metadata Buttons
        btn_frame = ttk.Frame(meta_frame)
        btn_frame.grid(row=8, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="Save PDF", command=self.save_pdf, width=14).pack(side='left', padx=3)
        ttk.Button(btn_frame, text="Save JSON", command=self.save_json_metadata, width=14).pack(side='left', padx=3)
        ttk.Button(btn_frame, text="Load JSON", command=self.load_json_metadata, width=14).pack(side='left', padx=3)

        # Page Operations
        page_frame = ttk.LabelFrame(main_frame, text=" Page Operations ", padding=10)
        page_frame.grid(row=2, column=0, sticky='ew', pady=5)

        ttk.Label(page_frame, text="Page Range (e.g., 1-3,5):").grid(row=0, column=0, sticky='w')
        self.page_range_var = tk.StringVar()
        ttk.Entry(page_frame, textvariable=self.page_range_var, width=30).grid(row=0, column=1, padx=10)
        ttk.Button(page_frame, text="Extract Pages", command=self.extract_pages, width=14).grid(row=0, column=2, padx=5)

        # Image Operations
        img_frame = ttk.LabelFrame(main_frame, text=" Image Operations ", padding=10)
        img_frame.grid(row=3, column=0, sticky='ew', pady=5)

        self.image_dir_var = tk.StringVar()
        ttk.Label(img_frame, text="Output Directory:").grid(row=0, column=0, sticky='w')
        ttk.Entry(img_frame, textvariable=self.image_dir_var, width=30).grid(row=0, column=1, padx=10)
        ttk.Button(img_frame, text="Browse", command=self.choose_image_dir, width=8).grid(row=0, column=2, padx=5)
        ttk.Button(img_frame, text="Extract Images", command=self.extract_images, width=14).grid(row=1, column=2, pady=5)

        # Utility Buttons
        util_frame = ttk.Frame(main_frame)
        util_frame.grid(row=4, column=0, sticky='ew', pady=15)
        ttk.Button(util_frame, text="Clear All Fields", command=self.clear_fields, width=18).pack(side='left', padx=5)
        ttk.Button(util_frame, text="Exit", command=self.exit_app, width=18).pack(side='left', padx=5)

        # Status Bar
        self.status_bar = ttk.Label(main_frame, text="Ready", relief='sunken', padding=5)
        self.status_bar.grid(row=5, column=0, sticky='ew')

    def setup_drag_and_drop(self) -> None:
        def handle_drop(event):
            file_path = self.clean_dropped_path(event.data)
            
            if self.is_valid_pdf(file_path):
                self.file_path.set(file_path)
                self.load_pdf(file_path=file_path)
            else:
                self.clear_fields()
                self.update_status("Only PDF files are allowed!", 'red')

        # Register drop target for entire window hierarchy
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', handle_drop)

    def clean_dropped_path(self, raw_path: str) -> str:
        path = raw_path.strip()
        if path.startswith('{') and path.endswith('}'):
            path = path[1:-1]
        return path

    def is_valid_pdf(self, file_path: str) -> bool:
        try:
            with open(file_path, 'rb') as f:
                header = f.read(5)
                return header.hex() == '255044462d'
        except Exception as e:
            return False

    def load_pdf(self, file_path: Path = None) -> None:
        if not file_path:
            file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if not file_path:
                return
            self.file_path.set(file_path)

        try:
            file_name = os.path.basename(file_path)
            reader = PdfReader(file_path)
            if not reader.pages:
                raise ValueError("PDF contains no pages")
                
            metadata = reader.metadata
            if metadata is None:
                self.clear_fields()
                self.update_status(f"No metadata found in {file_name}", 'purple')
                return
            for key, var in self.metadata_fields.items():
                var.set(metadata.get(key, ""))
            self.update_status(f"Loaded: {file_name}", 'green')
            
        except Exception as e:
            self.update_status(f"Error: {str(e)}", 'red')
            self.file_path.set("")

    def update_status(self, message: str, color: str = 'black') -> None:
        self.status_bar.config(text=message, foreground=color)

    def save_pdf(self) -> None: 
        """Save the PDF file with updated metadata."""
        file_path = self.file_path.get()
        if not file_path:
            messagebox.showerror("Error", "Please select a PDF file.")
            return

        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not output_path:
            return

        try:
            reader = PdfReader(file_path)
            writer = PdfWriter()

            for page in reader.pages:
                writer.add_page(page)

            metadata = {key: var.get() for key, var in self.metadata_fields.items()}
            writer.add_metadata(metadata)

            with open(output_path, "wb") as output_pdf:
                writer.write(output_pdf)

            messagebox.showinfo("Success", "PDF saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save PDF: {e}")

    def extract_pages(self) -> None:
        """Extract specified pages from the PDF and save as new file."""
        file_path = self.file_path.get()
        if not file_path:
            messagebox.showerror("Error", "Please select a PDF file.")
            return

        try:
            reader = PdfReader(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open PDF: {e}")
            return

        total_pages = len(reader.pages)
        if total_pages == 0:
            messagebox.showerror("Error", "The PDF has no pages.")
            return

        page_range = self.page_range_var.get().strip()
        if not page_range:
            messagebox.showerror("Error", "Please enter a page range.")
            return

        try:
            pages = self.parse_page_range(page_range, total_pages)
        except ValueError as e:
            messagebox.showerror("Error", str(e))
            return

        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not output_path:
            return

        try:
            writer = PdfWriter()
            writer.add_metadata({"/Creator":"Created by PDF Toolkit"})
            for page_num in pages:
                writer.add_page(reader.pages[page_num])

            with open(output_path, "wb") as out_file:
                writer.write(out_file)

            messagebox.showinfo("Success", "Pages extracted successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract pages: {e}")

    def parse_page_range(self, page_range_str: str, total_pages: int) -> list[int]:
        """Parse user input page range into list of page numbers."""
        parts = page_range_str.split(',')
        pages = []
        for part in parts:
            part = part.strip()
            if not part:
                continue
            if '-' in part:
                try:
                    start_str, end_str = part.split('-', 1)
                    start = int(start_str) - 1
                    end = int(end_str) - 1
                except ValueError:
                    raise ValueError(f"Invalid range format: {part}")

                if start < 0 or end >= total_pages:
                    raise ValueError(f"Page range {part} out of bounds (1-{total_pages})")
                if start > end:
                    raise ValueError(f"Invalid range order: {part}")
                pages.extend(range(start, end + 1))
            else:
                try:
                    page = int(part) - 1
                except ValueError:
                    raise ValueError(f"Invalid page number: {part}")

                if page < 0 or page >= total_pages:
                    raise ValueError(f"Page {part} out of bounds (1-{total_pages})")
                pages.append(page)

        # Remove duplicates while preserving order
        seen = set()
        return [p for p in pages if not (p in seen or seen.add(p))]

    def choose_image_dir(self) -> None:
        """Select output directory for image extraction."""
        directory = filedialog.askdirectory()
        if directory:
            self.image_dir_var.set(directory)

    def extract_images(self) -> None:
        """Extract all images from the PDF and save to selected directory."""
        file_path = self.file_path.get()
        if not file_path:
            messagebox.showerror("Error", "Please select a PDF file.")
            return

        output_dir = self.image_dir_var.get()
        if not output_dir:
            messagebox.showerror("Error", "Please select an output directory.")
            return

        try:
            reader = PdfReader(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open PDF: {e}")
            return

        try:
            image_count = 0
            for page_num in range(len(reader.pages)):
                page = reader.pages[page_num]
                xobjects = page.get('/Resources', {}).get('/XObject', {})

                for obj_key in xobjects:
                    obj = xobjects[obj_key]
                    if obj.get('/Subtype') == '/Image':
                        image = obj
                        image_count += 1

                        # Determine image type
                        filters = image.get('/Filter', [])
                        if not isinstance(filters, list):
                            filters = [filters]

                        if '/DCTDecode' in filters:
                            ext = '.jpg'
                        elif '/FlateDecode' in filters:
                            ext = '.png'
                        elif '/CCITTFaxDecode' in filters:
                            ext = '.tiff'
                        else:
                            ext = '.bin'

                        # Save image
                        filename = f"image_{image_count}{ext}"
                        output_path = os.path.join(output_dir, filename)
                        with open(output_path, "wb") as img_file:
                            img_file.write(image.get_data())

            if image_count == 0:
                messagebox.showinfo("Info", "No images found in the PDF.")
            else:
                messagebox.showinfo("Success", f"Extracted {image_count} images to:\n{output_dir}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract images: {e}")

    def save_json_metadata(self) -> None:
        """Save current metadata to JSON file."""
        metadata = {key: var.get() for key, var in self.metadata_fields.items()}
        output_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if not output_path:
            return

        try:
            with open(output_path, 'w', encoding='utf-8') as json_file:
                json.dump(metadata, json_file, ensure_ascii=False, indent=4)
            messagebox.showinfo("Success", "Metadata saved to JSON successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save JSON: {e}")

    def load_json_metadata(self) -> None:
        """Load metadata from JSON file into the application."""
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if not file_path:
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as json_file:
                metadata = json.load(json_file)

            for key, var in self.metadata_fields.items():
                var.set(metadata.get(key, ""))

            self.update_status("Metadata loaded successfully.", "green")
        except Exception as e:
            self.update_status(f"Failed to load metadata: {e}", "red")

    def clear_fields(self) -> None:
        """Clear all input fields."""
        self.file_path.set("")
        self.page_range_var.set("")
        self.image_dir_var.set("")
        for var in self.metadata_fields.values():
            var.set("")
        self.update_status("All fields cleared.", "blue")

    def exit_app(self) -> None:
        """Close the application."""
        self.root.destroy()


if __name__ == "__main__":
    icon_path = Path("assets/pdftoolkit.ico")
    root = TkinterDnD.Tk()
    root.withdraw()
    root.update()
    PDFToolkitApp(root, icon_path)
    root.after(50, root.deiconify)
    root.mainloop()