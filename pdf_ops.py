import json
from PIL import Image
from pathlib import Path
from datetime import datetime, timezone
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from typing import Tuple, Optional, Union, List, Dict


class PdfTools:
    """A utility class for handling PDF files with various operations."""

    def read_pdf(self, file_path: Path) -> Tuple[str, Dict]:
        """
        Read the PDF file and return its name and metadata.

        Args:
            file_path: Path to the PDF file.

        Returns:
            Tuple containing the file name and metadata dictionary.

        Raises:
            ValueError: If the PDF contains no pages.
            IOError: If the file cannot be read.
        """
        try:
            reader = PdfReader(file_path)
            if not reader.pages:
                raise ValueError(f"{file_path.name} contains no pages")
            metadata = reader.metadata or {}
            return file_path.name, metadata
        except Exception as e:
            raise IOError(f"Error reading {file_path}: {e}")

    def is_valid_pdf(self, file_path: Path) -> bool:
        """
        Check if the file is a valid PDF.

        Args:
            file_path: Path to the file.

        Returns:
            True if the file is a valid PDF, False otherwise.
        """
        try:
            with file_path.open('rb') as f:
                header = f.read(5)
                return header.hex() == "255044462d"
        except Exception as e:
            return False

    def save_metadata_to_json(self, metadata: Dict, output_path: Path) -> None:
            """
            Save metadata to a JSON file.

            Args:
                metadata: Metadata dictionary to save.
                output_path: Path to the output JSON file.

            Raises:
                IOError: If the file cannot be written.
            """
            try:
                with output_path.open("w", encoding="utf-8") as f:
                    json.dump(metadata, f, ensure_ascii=False, indent=4)
            except Exception as e:
                raise IOError(f"Error saving metadata to {output_path}: {e}")

    def get_metadata_from_json(self, json_path: Path) -> Dict:
            """
            Read metadata from a JSON file.

            Args:
                json_path: Path to the JSON file.

            Returns:
                Metadata dictionary, or empty dict if file cannot be read.
            """
            try:
                with json_path.open("r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return {}

    def save_pdf(self, file_path: Path, metadata: Optional[Dict], output_path: Path) -> None:
            """
            Save the PDF with optional metadata.

            Args:
                file_path: Path to the input PDF file.
                metadata: Metadata to add, or None.
                output_path: Path to the output PDF file.

            Raises:
                IOError: If the PDF cannot be read or written.
            """
            try:
                reader = PdfReader(file_path)
                writer = PdfWriter()
                for page in reader.pages:
                    writer.add_page(page)
                if metadata:
                    writer.add_metadata(metadata)
                with output_path.open("wb") as f:
                    writer.write(f)
            except Exception as e:
                raise IOError(f"Error saving PDF to {output_path}: {e}")

    def parse_page_range(self, page_range: str, total_pages: int) -> List[int]:
            """
            Parse a page range string into a list of page indices.

            Args:
                page_range: Page range string (e.g., "1-3,5,7-9").
                total_pages: Total number of pages in the PDF.

            Returns:
                List of page indices (0-based).

            Raises:
                ValueError: If the page range is invalid.
            """
            pages = set()
            for part in page_range.split(","):
                part = part.strip()
                if not part:
                    continue
                if "-" in part:
                    try:
                        start, end = map(int, part.split("-", 1))
                        start -= 1
                        end -= 1
                        if not (0 <= start <= end < total_pages):
                            raise ValueError(f"Range {part} out of bounds (1-{total_pages})")
                        pages.update(range(start, end + 1))
                    except ValueError as e:
                        raise ValueError(f"Invalid range: {part}") from e
                else:
                    try:
                        page = int(part) - 1
                        if not 0 <= page < total_pages:
                            raise ValueError(f"Page {part} out of bounds (1-{total_pages})")
                        pages.add(page)
                    except ValueError:
                        raise ValueError(f"Invalid page: {part}")
            return sorted(pages)
    
    def extract_pages(self, file_path: Path, page_range: str, output_path: Path) -> None:
            """
            Extract specified pages from the PDF and save to a new PDF.

            Args:
                file_path: Path to the input PDF file.
                page_range: Page range string (e.g., "1-3,5,7-9").
                output_path: Path to the output PDF file.

            Raises:
                ValueError: If the page range is invalid or no pages to extract.
                IOError: If the PDF cannot be read or written.
            """
            try:
                reader = PdfReader(file_path)
                total_pages = len(reader.pages)
                if not total_pages:
                    raise ValueError(f"No pages in {file_path.name}")
                pages = self.parse_page_range(page_range, total_pages)
                if not pages:
                    raise ValueError("No pages to extract")
                writer = PdfWriter()
                writer.add_metadata(
                    {"/Creator":"Created by PDF Toolkit", 
                    "/Producer":"PDF Toolkit with PyPDF2",
                    "/CreationDate": datetime.now(timezone.utc).strftime("D:%Y%m%d%H%M%S+00'00'"),
                    "/ModDate": datetime.now(timezone.utc).strftime("D:%Y%m%d%H%M%S+00'00'")
                    }
                )
                for page_idx in pages:
                    writer.add_page(reader.pages[page_idx])
                with output_path.open("wb") as f:
                    writer.write(f)
            except Exception as e:
                raise IOError(f"Error extracting pages to {output_path}: {e}")

    def get_image_extension(self, filters):
        """
        Determine the file extension based on the image filters.
        
        Args:
        filters (list): List of filter names.
        
        Returns:
        str: File extension.
        """
        if "/DCTDecode" in filters:
            return ".jpg"
        elif "/JPXDecode" in filters:
            return ".jp2"
        elif "/FlateDecode" in filters or "/LZWDecode" in filters:
            return ".png"
        elif "/CCITTFaxDecode" in filters:
            return ".tiff"
        else:
            return ".bin"

    def extract_images_from_pdf(self, file_path: Path, output_dir: Path) -> List[Path]:
        """
        Extract images from a PDF file and save them to a specified directory.
        
        Args:
        file_path (Path): Path to the PDF file.
        output_dir (Path): Directory to save the extracted images.
        
        Returns:
        List[Path]: List of paths to the extracted image files.
        """
        if not output_dir.exists():
            output_dir.mkdir(parents=True, exist_ok=True)
        try:
            image_count = 0
            reader = PdfReader(file_path)
            image_paths = []
            for page_num in range(len(reader.pages)):
                page = reader.pages[page_num]
                # Get the XObject dictionary from the page's resources
                x_objects = page.get('/Resources', {}).get('/XObject', {})
                for object_key in x_objects:
                    x_object = x_objects[object_key]
                    if x_object.get('/Subtype') == '/Image':
                        image_count += 1
                        image_obj = x_object
                        filters = image_obj.get('/Filter', [])
                        if not isinstance(filters, list):
                            filters = [filters]
                        ext = self.get_image_extension(filters)
                        file_name = f"page{page_num+1}_img{image_count}{ext}"
                        image_file_path = output_dir / file_name
                        with image_file_path.open("wb") as f:
                            f.write(image_obj.get_data())
                        image_paths.append(image_file_path)
            return image_paths
        except Exception as e:
            raise IOError(f"Error extracting images from {file_path}: {e}")

    def concatenate_pdfs(self, pdf_files: List[Path], output_path: Path) -> None:
            """
            Merge multiple PDFs into a single PDF.

            Args:
                pdf_files: List of paths to the PDF files to merge.
                output_path: Path to the output merged PDF file.

            Raises:
                IOError: If the PDFs cannot be read or merged.
            """
            try:
                merger = PdfMerger()
                merger.add_metadata(
                    {
                    "/Creator": "Created by PDF Toolkit",
                    "/Producer": "PDF Toolkit with PyPDF2",
                    "/CreationDate": datetime.now(timezone.utc).strftime("D:%Y%m%d%H%M%S+00'00'"),
                    "/ModDate": datetime.now(timezone.utc).strftime("D:%Y%m%d%H%M%S+00'00'")
                    }
                )
                for pdf_file in pdf_files:
                    merger.append(pdf_file)
                with output_path.open("wb") as f:
                    merger.write(f)
            except Exception as e:
                raise IOError(f"Error merging PDFs to {output_path}: {e}")

    def create_pdf_from_images(
        self,
        image_files: List[Path],
        margin_left: float,
        margin_right: float,
        margin_top: float,
        margin_bottom: float,
        page_size: str,
        orientation: str,
        scaling_option: Union[str, int],
        output_path: Path
    ) -> None:
        """
        Convert images to a PDF with custom settings.

        Args:
            image_files: List of paths to the image files.
            margin_left: Left margin in millimeters.
            margin_right: Right margin in millimeters.
            margin_top: Top margin in millimeters.
            margin_bottom: Bottom margin in millimeters.
            page_size: Page size (e.g., "A4", "Letter").
            orientation: Page orientation ("portrait" or "landscape").
            scaling_option: Scaling option ("scale to fit", "actual size", "stretch to fill").
            output_path: Path to the output PDF file.

        Raises:
            ValueError: If parameters are invalid.
            IOError: If images cannot be read or PDF cannot be written.
        """
        page_sizes = {
            "A4": (595, 842), "A3": (842, 1191), "Letter": (612, 792),
            "Legal": (612, 1008), "A5": (420, 595), "A6": (298, 420),
            "B5": (498, 708), "B4": (708, 1000), "B3": (1000, 1414),
            "Tabloid": (792, 1224), "Ledger": (1224, 792), "Executive": (522, 756),
            "Foolscap": (648, 936), "Quarto": (610, 780), "10x14": (720, 1008),
            "11x17": (792, 1224), "Statement": (396, 612), "Folio": (612, 936),
        }

        if page_size not in page_sizes:
            raise ValueError(f"Invalid page size: {page_size}")

        page_width, page_height = page_sizes[page_size]
        if orientation.lower() == "landscape":
            page_width, page_height = page_height, page_width

        MM_TO_PT = 2.83465
        margins = [margin * MM_TO_PT for margin in (margin_left, margin_right,
                                                    margin_top, margin_bottom)]
        available_width = page_width - margins[0] - margins[1]
        available_height = page_height - margins[2] - margins[3]

        scaling = (scaling_option.lower() if isinstance(scaling_option, str)
                else {1: "scale to fit", 2: "stretch to fit", 3: "actual size", 4: "stretch to fill"}
                .get(scaling_option, "scale to fit"))

        pdf_pages = []
        for image_file in image_files:
            try:
                img = Image.open(image_file).convert("RGB")
                img_width, img_height = img.size
                if scaling == "stretch to fill":
                    new_width, new_height = int(available_width), int(available_height)
                elif scaling == "stretch to fit":
                    ratio = max(available_width / img_width, available_height / img_height)
                    new_width, new_height = int(img_width * ratio), int(img_height * ratio)
                else:
                    ratio = (1.0 if scaling == "actual size" else
                            min(available_width / img_width, available_height / img_height)) # stretch to fill
                    new_width, new_height = int(img_width * ratio), int(img_height * ratio)

                resized_img = img.resize((new_width, new_height), Image.LANCZOS)
                page = Image.new("RGB", (page_width, page_height), "white")
                x = int(margins[0] + (available_width - new_width) / 2)
                y = int(margins[2] + (available_height - new_height) / 2)
                page.paste(resized_img, (x, y))
                pdf_pages.append(page)
            except Exception as e:
                raise IOError(f"Error processing {image_file}: {e}")

        if not pdf_pages:
            raise ValueError("No images to convert")
        try:
            pdf_pages[0].save(
                output_path, 
                save_all=True, 
                append_images=pdf_pages[1:],
                title=str(output_path.name),
                creator="Created by PDF Toolkit",
                producer="PDF Toolkit with PyPDF2",
                creationDate=datetime.now(timezone.utc).strftime("D:%Y%m%d%H%M%S+00'00'"),
                modDate=datetime.now(timezone.utc).strftime("D:%Y%m%d%H%M%S+00'00'")
            )
        except Exception as e:
            raise IOError(f"Error saving PDF to {output_path}: {e}")
 