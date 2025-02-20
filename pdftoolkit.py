import tkinter as tk
from pathlib import Path
from pdf_ops import PdfTools
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import messagebox, filedialog, ttk


class PDFToolkitApp(TkinterDnD.Tk):
    """GUI application for PDF manipulation tools."""

    def __init__(self):
        super().__init__()
        self.withdraw()  # Hide window until centered
        self.title("PDF Toolkit")
        self.geometry("620x680")
        self.iconbitmap(default="assets/pdftoolkit.ico")
        self.pdf_tools = PdfTools()

        self._setup_ui()
        self.after(50, self.deiconify)  # Show window after setup

    def _setup_ui(self):
        """Initialize the UI components."""
        self.create_menu()
        self.create_main_interface()
        self.setup_drag_and_drop()
        self.center_window()

    def center_window(self):
        """Center the window on the screen."""
        self.update_idletasks()
        width, height = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"+{x}+{y}")

    def create_menu(self):
        """Create the application menu bar."""
        menu_bar = tk.Menu(self)
        self.config(menu=menu_bar)

        # File menu
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Open", command=self.load_pdf)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.exit_application)
        menu_bar.add_cascade(label="File", menu=file_menu)

        # Edit menu
        edit_menu = tk.Menu(menu_bar, tearoff=0)
        edit_menu.add_command(label="Clear All", command=self.clear_fields)
        menu_bar.add_cascade(label="Edit", menu=edit_menu)

        # Help menu
        help_menu = tk.Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="About", command=self.about)
        menu_bar.add_cascade(label="Help", menu=help_menu)

    def about(self):
        """Display the About dialog."""
        messagebox.showinfo(
            "About",
            "PDF Toolkit | A simple PDF utility\n"
            "Version 1.3\n"
            "Source: github.com/ciwga/pdftoolkit"
        )

    def create_main_interface(self):
        """Set up the scrollable main interface with notebook tabs."""
        # Container for scrollable content
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)

        # Scrollable canvas
        self.canvas = tk.Canvas(container)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Scrollable frame inside canvas
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.scrollable_window = self.canvas.create_window(
            (0, 0), window=self.scrollable_frame, anchor="nw"
        )
        self.canvas.bind(
            "<Configure>",
            lambda e: self.canvas.itemconfig(self.scrollable_window, width=e.width)
        )

        # Notebook for tabs
        self.notebook = ttk.Notebook(self.scrollable_frame)
        self.notebook.pack(expand=True, fill="both")

        # Add tabs
        self.create_pdf_operations_tab()
        self.create_pdf_merger_tab()
        self.create_image_to_pdf_tab()

        # Status bar
        status_frame = ttk.Frame(self.scrollable_frame)
        status_frame.pack(fill="both", padx=20, pady=15)
        ttk.Style().configure("Status.TLabel", relief="sunken", padding=5)
        self.status_bar = ttk.Label(
            status_frame, text="Ready", style="Status.TLabel", anchor="w"
        )
        self.status_bar.pack(fill="x")

    def create_pdf_operations_tab(self):
        """Create the PDF Operations tab with file selection, metadata, and tools."""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="PDF Operations")
        main_frame = ttk.Frame(frame, padding=10)
        main_frame.pack(expand=True, fill="both", padx=10, pady=10)

        # PDF File Selection
        pdf_frame = ttk.LabelFrame(main_frame, text="PDF File", padding=10)
        pdf_frame.pack(fill="x", pady=5)
        self.pdf_file_path = tk.StringVar()
        ttk.Entry(pdf_frame, textvariable=self.pdf_file_path, width=50).pack(
            side="left", fill="x", expand=True, padx=5
        )
        ttk.Button(pdf_frame, text="Browse", command=self.load_pdf).pack(side="left")

        # Metadata Editor
        metadata_frame = ttk.LabelFrame(main_frame, text="Metadata Editor", padding=10)
        metadata_frame.pack(fill="x", pady=10)
        self.metadata_fields = {
            "/Title": tk.StringVar(),
            "/Subject": tk.StringVar(),
            "/Keywords": tk.StringVar(),
            "/Author": tk.StringVar(),
            "/Creator": tk.StringVar(),
            "/Producer": tk.StringVar(),
            "/CreationDate": tk.StringVar(),
            "/ModDate": tk.StringVar(),
        }
        for i, (key, var) in enumerate(self.metadata_fields.items()):
            ttk.Label(metadata_frame, text=f"{key}:", width=13, anchor="e").grid(
                row=i, column=0, padx=5, pady=2
            )
            ttk.Entry(metadata_frame, textvariable=var, width=40).grid(
                row=i, column=1, pady=2, padx=6
            )

        # Metadata Buttons
        btn_frame = ttk.Frame(metadata_frame)
        btn_frame.grid(row=len(self.metadata_fields), columnspan=2, pady=10)
        buttons = [
            ("Save PDF", self.save_pdf_with_metadata),
            ("Save JSON", self.dump_metadata),
            ("Load JSON", self.load_metadata_from_json),
        ]
        for col, (text, cmd) in enumerate(buttons, 1):
            ttk.Button(btn_frame, text=text, width=16, command=cmd).grid(
                row=0, column=col, padx=4, pady=2
            )

        # PDF Splitter
        splitter_frame = ttk.LabelFrame(main_frame, text="PDF Splitter", padding=10)
        splitter_frame.pack(fill="x", pady=5)
        ttk.Label(splitter_frame, text="Page Range (e.g., 1-3,5):").grid(
            row=0, column=0, sticky="w"
        )
        self.page_range_var = tk.StringVar()
        ttk.Entry(splitter_frame, textvariable=self.page_range_var, width=30).grid(
            row=0, column=1, padx=10
        )
        ttk.Button(splitter_frame, text="Split PDF", width=12, command=self.extract_pages).grid(
            row=0, column=2, padx=5, pady=2
        )

        # PDF Image Extractor
        extractor_frame = ttk.LabelFrame(main_frame, text="PDF Image Extractor", padding=10)
        extractor_frame.pack(fill="x", pady=5)
        ttk.Label(extractor_frame, text="Output Directory:").grid(row=0, column=0, sticky="w")
        self.image_dir_var = tk.StringVar()
        ttk.Entry(extractor_frame, textvariable=self.image_dir_var, width=33).grid(
            row=0, column=1, padx=10
        )
        ttk.Button(extractor_frame, text="Browse", width=16, command=self.browse_image_dir).grid(
            row=0, column=2, padx=2
        )
        ttk.Button(
            extractor_frame, text="Extract Images", width=16, command=self.extract_all_images
        ).grid(row=1, column=2, pady=2)

    def create_pdf_merger_tab(self):
        """Create the PDF Merger tab with listbox and controls."""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="PDF Merger")
        main_frame = ttk.Frame(frame, padding=10)
        main_frame.pack(expand=True, fill="both", padx=10, pady=10)

        # Listbox for PDFs
        self.merger_listbox = self._create_listbox(main_frame)
        self.merger_listbox.drop_target_register(DND_FILES)
        self.merger_listbox.dnd_bind("<<Drop>>", self.handle_pdf_drop)

        # Control Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill="x", pady=5)
        self._add_listbox_controls(btn_frame, self.merger_listbox, self.add_pdfs_to_merge)

        # Merge Button
        ttk.Button(main_frame, text="Merge PDFs", command=self.merge_pdfs).pack(pady=10, side="top", fill="both")

    def create_image_to_pdf_tab(self):
        """Create the Image to PDF tab with listbox and conversion options."""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Image to PDF")
        main_frame = ttk.Frame(frame, padding=10)
        main_frame.pack(expand=True, fill="both", padx=10, pady=10)

        # Listbox for Images
        self.image_listbox = self._create_listbox(main_frame)
        self.image_listbox.drop_target_register(DND_FILES)
        self.image_listbox.dnd_bind("<<Drop>>", self.handle_image_drop)

        # Control Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill="x", pady=5)
        self._add_listbox_controls(btn_frame, self.image_listbox, self.add_images_to_convert)

        # Conversion Options
        options_frame = ttk.LabelFrame(main_frame, text="Conversion Options", padding=10)
        options_frame.pack(fill="x", pady=5)

        # Page Size
        ttk.Label(options_frame, text="Page Size:").grid(row=0, column=0, padx=5, sticky="e")
        self.page_size_var = tk.StringVar(value="A4")
        ttk.Combobox(
            options_frame,
            textvariable=self.page_size_var,
            values=["A4", "A3", "Letter", "Legal", "A5", "A6", "B5", "B4", "B3", "Tabloid", "Ledger", 
                    "Executive", "Foolscap", "Quarto", "10x14", "11x17", "Statement", "Folio"],
            state="readonly",
            width=10,
        ).grid(row=0, column=1, sticky="w")

        # Orientation
        ttk.Label(options_frame, text="Orientation:").grid(row=0, column=2, padx=5, sticky="e")
        self.orientation_var = tk.StringVar(value="Portrait")
        ttk.Radiobutton(options_frame, text="Portrait", variable=self.orientation_var, value="Portrait").grid(
            row=0, column=3, padx=5, sticky="w"
        )
        ttk.Radiobutton(options_frame, text="Landscape", variable=self.orientation_var, value="Landscape").grid(
            row=0, column=4, padx=5, sticky="w"
        )

        # Scaling Options
        ttk.Label(options_frame, text="Scaling:").grid(row=0, column=5, padx=5, sticky="e")
        self.scaling_var = tk.StringVar(value="Scale to Fit")
        ttk.Combobox(
            options_frame,
            textvariable=self.scaling_var,
            values=["Scale to Fit", "Stretch to Fit", "Actual Size", "Stretch to Fill"],
            state="readonly",
            width=12,
        ).grid(row=0, column=6, sticky="w")

        # Margins
        margins_frame = ttk.Frame(options_frame)
        margins_frame.grid(row=1, column=0, columnspan=5, pady=5, sticky="w")
        ttk.Label(margins_frame, text="Margins (mm):").pack(side="left")
        self.margins = {}
        for label, attr in [("L", "left"), ("R", "right"), ("T", "top"), ("B", "bottom")]:
            entry = ttk.Entry(margins_frame, width=5)
            entry.insert(0, "10")
            entry.pack(side="left", padx=2)
            ttk.Label(margins_frame, text=label).pack(side="left")
            self.margins[attr] = entry

        # Convert Button
        ttk.Button(main_frame, text="Convert to PDF", command=self.convert_images_to_pdf).pack(pady=10)

    def _create_listbox(self, parent):
        """Create a scrollable listbox with a scrollbar."""
        list_frame = ttk.Frame(parent)
        list_frame.pack(fill="both", expand=True, pady=5)
        listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED, height=6)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        return listbox

    def _add_listbox_controls(self, parent, listbox, add_command):
        """Add control buttons for a listbox."""
        buttons = [
            ("Add", add_command),
            ("Remove Selected", lambda: self.remove_selected(listbox)),
            ("Move Up", lambda: self.move_list_item(listbox, -1)),
            ("Move Down", lambda: self.move_list_item(listbox, 1)),
            ("Clear All", lambda: listbox.delete(0, tk.END)),
        ]
        for text, cmd in buttons:
            ttk.Button(parent, text=text, command=cmd).pack(side="left", padx=2, fill="x", expand=True)

    def setup_drag_and_drop(self):
        """Enable drag-and-drop for loading PDF files into the main window."""
        self.drop_target_register(DND_FILES)
        self.dnd_bind("<<Drop>>", self.handle_drop)

    def handle_drop(self, event):
        """Handle dropped files for the main window."""
        file_path = Path(self.clean_dropped_path(event.data))
        if self.pdf_tools.is_valid_pdf(file_path):
            self.pdf_file_path.set(str(file_path))
            self.load_pdf(file_path)
        else:
            self.clear_fields()
            self.update_status("Only PDF files are allowed!", "red")

    def handle_pdf_drop(self, event):
        """Handle dropped PDF files for the merger listbox."""
        file_path = Path(self.clean_dropped_path(event.data))
        if self.pdf_tools.is_valid_pdf(file_path):
            self._handle_file_drop(event, self.merger_listbox, [".pdf"])

    def handle_image_drop(self, event):
        """Handle dropped image files for the image listbox."""
        self._handle_file_drop(event, self.image_listbox, [".jpg", ".jpeg", ".png", ".bmp", ".gif"])

    def _handle_file_drop(self, event, listbox, extensions):
        """Generic handler for dropping files into a listbox."""
        files = self.tk.splitlist(event.data)
        for file in files:
            if Path(file).suffix.lower() in extensions:
                listbox.insert(tk.END, file)

    @staticmethod
    def clean_dropped_path(raw_path: str) -> str:
        """Clean and normalize a dropped file path."""
        path = raw_path.strip()
        return path[1:-1] if path.startswith("{") and path.endswith("}") else path

    def load_pdf(self, file_path=None):
        """Load a PDF file and populate metadata fields."""
        if not file_path:
            file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if not file_path:
                return
            file_path = Path(file_path)
            self.pdf_file_path.set(str(file_path))

        try:
            file_name, metadata = self.pdf_tools.read_pdf(file_path)
            if metadata is None:
                self.update_status(f"No metadata found in {file_name}", "purple")
                return
            for key, var in self.metadata_fields.items():
                var.set(metadata.get(key, ""))
            self.update_status(f"Loaded: {file_name}", "green")
        except Exception as e:
            self.update_status(f"Error: {str(e)}", "red")
            self.pdf_file_path.set("")

    def update_status(self, message: str, color: str = "black"):
        """Update the status bar with a message and color."""
        self.status_bar.config(text=message, foreground=color)

    def clear_fields(self):
        """Reset all input fields to empty and clear merger and image listboxes."""
        self.pdf_file_path.set("")
        self.page_range_var.set("")
        self.image_dir_var.set("")
        for var in self.metadata_fields.values():
            var.set("")
        self.merger_listbox.delete(0, tk.END)
        self.image_listbox.delete(0, tk.END)
        self.update_status("All fields cleared.", "blue")

    def browse_image_dir(self):
        """Select an output directory for image extraction."""
        directory = filedialog.askdirectory()
        if directory:
            self.image_dir_var.set(directory)

    def load_metadata_from_json(self):
        """Load metadata from a JSON file into fields."""
        json_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if not json_path:
            return
        try:
            metadata = self.pdf_tools.get_metadata_from_json(Path(json_path))
            if metadata:
                for key, var in self.metadata_fields.items():
                    var.set(metadata.get(key, ""))
                self.update_status(f"Loaded metadata from {Path(json_path).name}", "green")
            else:
                self.update_status(f"No metadata in {Path(json_path).name}", "purple")
        except Exception as e:
            self._show_error(f"An error occurred: {e}", "red")

    def dump_metadata(self):
        """Save current metadata to a JSON file."""
        metadata = {key: var.get() for key, var in self.metadata_fields.items()}
        if not any(metadata.values()):
            self.update_status("No metadata to save!", "purple")
            return
        output_path = self._save_file_dialog(".json", [("JSON files", "*.json")])
        if not output_path:
            return
        self._save_to_file(self.pdf_tools.save_metadata_to_json, metadata, output_path, "Metadata")

    def save_pdf_with_metadata(self):
        """Save the PDF with updated metadata."""
        input_file = self.pdf_file_path.get()
        if not input_file:
            self._show_error("Please load a PDF file first!", "purple")
            return
        metadata = {key: var.get() for key, var in self.metadata_fields.items()}
        if not any(metadata.values()):
            self.update_status("No metadata to save!", "purple")
            return
        output_path = self._save_file_dialog(".pdf", [("PDF files", "*.pdf")])
        if not output_path:
            return
        self._save_to_file(self.pdf_tools.save_pdf, metadata, output_path, "PDF")

    def _save_file_dialog(self, extension: str, filetypes: list) -> Path:
        """Open a save file dialog and return the selected path."""
        path = filedialog.asksaveasfilename(defaultextension=extension, filetypes=filetypes)
        return Path(path) if path else None

    def _save_to_file(self, save_func, data, output_path: Path, item_type: str):
        """Generic method to save data to a file with error handling."""
        try:
            if item_type == "PDF":
                input_file = self.pdf_file_path.get()
                save_func(input_file, data, output_path)
            else:
                save_func(data, output_path)
            self.update_status(f"{item_type} saved to {output_path}", "green")
            messagebox.showinfo("Success", f"{item_type} saved to {output_path.name}")
        except Exception as e:
            self._show_error(f"An error occurred: {e}", "red")

    def extract_pages(self):
        """Extract specified pages from the loaded PDF."""
        input_file = self.pdf_file_path.get()
        if not input_file:
            self._show_error("Please load a PDF file first!", "purple")
            return
        page_range = self.page_range_var.get().strip()
        if not page_range:
            self._show_error("Please enter a page range!", "purple")
            return
        output_path = self._save_file_dialog(".pdf", [("PDF files", "*.pdf")])
        if not output_path:
            self.update_status("No output path selected.", "purple")
            return
        try:
            self.pdf_tools.extract_pages(input_file, page_range, output_path)
            self.update_status("Pages extracted successfully!", "green")
            messagebox.showinfo("Success", "Pages extracted successfully!")
        except Exception as e:
            self._show_error(str(e), "red")

    def extract_all_images(self):
        """Extract all images from the loaded PDF."""
        input_file = self.pdf_file_path.get()
        if not input_file:
            self._show_error("Please load a PDF file first!", "purple")
            return
        output_dir = self.image_dir_var.get()
        if not output_dir:
            self._show_error("Please select an output directory!", "purple")
            return
        output_dir = Path(output_dir)
        try:
            extracted = self.pdf_tools.extract_images_from_pdf(input_file, output_dir)
            self.update_status(f"{len(extracted)} images extracted to {output_dir}", "green")
            messagebox.showinfo("Success", f"{len(extracted)} images extracted successfully!")
        except Exception as e:
            self._show_error(str(e), "red")

    def _show_error(self, message: str, color: str):
        """Display an error message in the status bar and a dialog."""
        self.update_status(message, color)
        messagebox.showerror("Error", message)

    def add_pdfs_to_merge(self):
        """Add PDF files to the merger listbox."""
        self._add_files_to_listbox(self.merger_listbox, [("PDF files", "*.pdf")])

    def add_images_to_convert(self):
        """Add image files to the image listbox."""
        self._add_files_to_listbox(
            self.image_listbox, [("Image files", "*.jpg *.jpeg *.png *.bmp *.gif")]
        )

    def _add_files_to_listbox(self, listbox, filetypes):
        """Generic method to add files to a listbox."""
        files = filedialog.askopenfilenames(filetypes=filetypes)
        for file in files:
            listbox.insert(tk.END, file)

    def remove_selected(self, listbox):
        """Remove selected items from a listbox."""
        for index in reversed(listbox.curselection()):
            listbox.delete(index)

    def move_list_item(self, listbox, direction):
        """Move selected items up or down in a listbox."""
        selected = listbox.curselection()
        if not selected:
            return
        for index in selected:
            if direction == -1 and index > 0:  # Move up
                item = listbox.get(index)
                listbox.delete(index)
                listbox.insert(index - 1, item)
                listbox.select_set(index - 1)
            elif direction == 1 and index < listbox.size() - 1:  # Move down
                item = listbox.get(index)
                listbox.delete(index)
                listbox.insert(index + 1, item)
                listbox.select_set(index + 1)

    def merge_pdfs(self):
        """Merge PDFs from the listbox (placeholder)."""
        files = list(self.merger_listbox.get(0, tk.END))
        if not files:
            self._show_error("No PDFs selected to merge!", "purple")
            return
        output_path = self._save_file_dialog(".pdf", [("PDF files", "*.pdf")])
        if not output_path:
            return
        try:
            self.pdf_tools.concatenate_pdfs(files, output_path)
            self.update_status(f"PDFs merged successfully to {output_path}", "green")
            messagebox.showinfo("Success", "PDFs merged successfully!")
        except Exception as e:
            self._show_error(str(e), "red")

    def convert_images_to_pdf(self):
        """Convert images to PDF (placeholder)."""
        images = list(self.image_listbox.get(0, tk.END))
        if not images:
            self._show_error("No images selected to convert!", "purple")
            return
        output_path = self._save_file_dialog(".pdf", [("PDF files", "*.pdf")])
        if not output_path:
            return
        try:
            self.pdf_tools.create_pdf_from_images(
                images,
                float(self.margins["left"].get()),
                float(self.margins["right"].get()),
                float(self.margins["top"].get()),
                float(self.margins["bottom"].get()),
                self.page_size_var.get(),
                self.orientation_var.get(),
                self.scaling_var.get(),
                output_path
            )
            self.update_status(f"Images converted to PDF: {output_path}", "green")
            messagebox.showinfo("Success", "Images converted to PDF successfully!")
        except Exception as e:
            self._show_error(str(e), "red")

    def exit_application(self):
        """Close the application."""
        self.destroy()


if __name__ == "__main__":
    app = PDFToolkitApp()
    app.mainloop()