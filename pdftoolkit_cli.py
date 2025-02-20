import sys
import argparse
from pathlib import Path
from typing import Union
from pdf_ops import PdfTools


def scaling_value(value: Union[str, int]) -> Union[str, int]:
    valid_scaling = {"scale to fit", "stretch to fit", "actual size", "stretch to fill"}
    try:
        ivalue = int(value)
        if not 1 <= ivalue <= 4:
            raise argparse.ArgumentTypeError("Scaling integer must be between 1 and 4")
        return ivalue
    except ValueError:
        lower_value = value.lower()
        if lower_value not in valid_scaling:
            raise argparse.ArgumentTypeError(
                f"Invalid scaling option. Choose from {valid_scaling} or 1-4."
            )
        return lower_value


def main():
    parser = argparse.ArgumentParser(
        description="PDF Toolkit CLI - perform various PDF operations."
    )
    subparsers = parser.add_subparsers(dest="command", help="Sub-command help")

    # Subparser for reading PDF metadata
    parser_read = subparsers.add_parser("read", help="Read PDF metadata")
    parser_read.add_argument("pdf_file", type=Path, help="Path to the PDF file")

    # Subparser for validating a PDF file
    parser_validate = subparsers.add_parser("validate", help="Validate a PDF file")
    parser_validate.add_argument("pdf_file", type=Path, help="Path to the PDF file")

    # Subparser for extracting pages
    parser_extract = subparsers.add_parser(
        "extract-pages", help="Extract specific pages from a PDF"
    )
    parser_extract.add_argument("pdf_file", type=Path, help="Path to the PDF file")
    parser_extract.add_argument(
        "page_range", type=str, help="Page range (e.g., '1-3,5')"
    )
    parser_extract.add_argument(
        "output_pdf", type=Path, help="Path to the output PDF file"
    )

    # Subparser for extracting images
    parser_extract_images = subparsers.add_parser(
        "extract-images", help="Extract images from a PDF"
    )
    parser_extract_images.add_argument("pdf_file", type=Path, help="Path to the PDF file")
    parser_extract_images.add_argument(
        "output_dir", type=Path, help="Directory to save the extracted images"
    )

    # Subparser for merging PDFs
    parser_merge = subparsers.add_parser("merge", help="Merge multiple PDFs")
    parser_merge.add_argument(
        "pdf_files", type=Path, nargs="+", help="List of PDF files to merge"
    )
    parser_merge.add_argument(
        "output_pdf", type=Path, help="Path to the merged output PDF file"
    )

    # Subparser for creating a PDF from images
    parser_create = subparsers.add_parser(
        "create-from-images", help="Create a PDF from images"
    )
    parser_create.add_argument(
        "image_files", type=Path, nargs="+", help="List of image files"
    )
    parser_create.add_argument(
        "output_pdf", type=Path, help="Path to the output PDF file"
    )
    parser_create.add_argument(
    "--margins",
    type=float,
    nargs=4,
    metavar=("LEFT", "RIGHT", "TOP", "BOTTOM"),
    help="Margin values in mm for left, right, top, and bottom (overrides individual margins if provided)."
    )
    parser_create.add_argument(
        "--margin-left", type=float, default=10, help="Left margin in mm (default: 10)"
    )
    parser_create.add_argument(
        "--margin-right", type=float, default=10, help="Right margin in mm (default: 10)"
    )
    parser_create.add_argument(
        "--margin-top", type=float, default=10, help="Top margin in mm (default: 10)"
    )
    parser_create.add_argument(
        "--margin-bottom", type=float, default=10, help="Bottom margin in mm (default: 10)"
    )
    parser_create.add_argument(
        "--page-size", type=str, default="A4", help="Page size (e.g., A4, Letter)"
    )
    parser_create.add_argument(
        "--orientation",
        type=str,
        default="portrait",
        help="Page orientation ('portrait' or 'landscape')",
    )
    parser_create.add_argument(
        "--scaling",
        type=scaling_value,
        default="scale to fit",
        help=("Scaling option: accepts 'scale to fit', 'stretch to fit', "
              "'actual size', 'stretch to fill' or an integer (1-4) mapping to these options."),
    )

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    tools = PdfTools()

    try:
        if args.command == "read":
            name, metadata = tools.read_pdf(args.pdf_file)
            print(f"File: {name}")
            print("Metadata:")
            for key, value in metadata.items():
                print(f"  {key}: {value}")
        elif args.command == "validate":
            valid = tools.is_valid_pdf(args.pdf_file)
            status = "valid" if valid else "invalid"
            print(f"{args.pdf_file} is {status}.")
        elif args.command == "extract-pages":
            tools.extract_pages(args.pdf_file, args.page_range, args.output_pdf)
            print(f"Extracted pages saved to {args.output_pdf}")
        elif args.command == "extract-images":
            images = tools.extract_images_from_pdf(args.pdf_file, args.output_dir)
            print(f"Extracted {len(images)} image(s):")
            for image in images:
                print(f"  {image}")
        elif args.command == "merge":
            tools.concatenate_pdfs(args.pdf_files, args.output_pdf)
            print(f"Merged PDF saved to {args.output_pdf}")
        elif args.command == "create-from-images":
            if args.margins is not None:
                margin_left, margin_right, margin_top, margin_bottom = args.margins
            else:
                margin_left = args.margin_left
                margin_right = args.margin_right
                margin_top = args.margin_top
                margin_bottom = args.margin_bottom

            tools.create_pdf_from_images(
                args.image_files,
                margin_left,
                margin_right,
                margin_top,
                margin_bottom,
                args.page_size,
                args.orientation,
                args.scaling,
                args.output_pdf,
            )
            print(f"PDF created from images saved to {args.output_pdf}")
        else:
            parser.print_help()
            sys.exit(1)
    except Exception as err:
        print(f"Error: {err}", file=sys.stderr)
        sys.exit(1)

    sys.exit(0)


if __name__ == "__main__":
    main()
