#!/usr/bin/env python3
"""
HTML Slides to PDF and PPTX Converter

This script converts local HTML presentation files (SPA-style slides with JavaScript interactions)
into PDF and PPTX formats by automating browser screenshots through each slide.

Usage:
    python convert_slides.py <html_file1> [<html_file2> ...]

Requirements:
    - playwright
    - Pillow (PIL)
    - python-pptx
    - Install with: pip install playwright pillow python-pptx
    - After installing playwright, run: playwright install chromium
"""

import argparse
import io
import os
import tempfile
import time
from pathlib import Path

from PIL import Image
from playwright.sync_api import sync_playwright
from pptx import Presentation
from pptx.util import Inches


def convert_html_to_images(html_path: str, viewport_size: tuple = (1920, 1080)) -> list:
    """
    Convert an HTML presentation file to a list of screenshot images.

    Args:
        html_path: Path to the HTML file
        viewport_size: Tuple of (width, height) for browser viewport

    Returns:
        List of PIL Image objects representing each slide
    """
    images = []

    with sync_playwright() as p:
        # Launch browser in headless mode
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': viewport_size[0], 'height': viewport_size[1]})
        page = context.new_page()

        # Open the local HTML file
        file_url = f"file://{os.path.abspath(html_path)}"
        page.goto(file_url)

        # Wait for page to load completely
        page.wait_for_load_state('networkidle')

        while True:
            # Try to find the presentation container, fallback to full page
            container_selector = ".presentation-wrapper, .presentation-container"
            if page.locator(container_selector).count() > 0:
                # Screenshot the container
                screenshot_bytes = page.locator(container_selector).screenshot()
            else:
                # Full page screenshot
                screenshot_bytes = page.screenshot()

            # Convert bytes to PIL Image
            image = Image.open(io.BytesIO(screenshot_bytes))
            images.append(image)

            # Check if nextBtn is disabled
            next_btn = page.locator("#nextBtn")
            if next_btn.count() == 0:
                # If no nextBtn found, assume single slide or end
                break

            disabled_attr = next_btn.get_attribute("disabled")
            if disabled_attr is not None:
                # Button is disabled, we've reached the last slide
                break

            # Click next button and wait for transition
            next_btn.click()
            time.sleep(0.6)  # Wait for CSS transitions

        browser.close()

    return images


def images_to_pdf(images: list, output_path: str):
    """
    Convert a list of PIL Images to a single PDF file.

    Args:
        images: List of PIL Image objects
        output_path: Path to save the PDF
    """
    if not images:
        raise ValueError("No images provided for PDF creation")

    # Convert all images to RGB mode for PDF compatibility
    rgb_images = [img.convert('RGB') for img in images]

    # Save as PDF
    rgb_images[0].save(output_path, save_all=True, append_images=rgb_images[1:])


def images_to_pptx(images: list, output_path: str, slide_size: tuple = (16, 9)):
    """
    Convert a list of PIL Images to a PPTX presentation.

    Args:
        images: List of PIL Image objects
        output_path: Path to save the PPTX
        slide_size: Tuple of (width, height) in inches for slides
    """
    if not images:
        raise ValueError("No images provided for PPTX creation")

    # Create presentation with custom slide size (16:9 aspect ratio)
    prs = Presentation()
    prs.slide_width = Inches(slide_size[0])
    prs.slide_height = Inches(slide_size[1])

    for img in images:
        # Add a blank slide
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout

        # Save image temporarily for insertion
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
            img.save(temp_file.name, 'PNG')
            temp_path = temp_file.name

        try:
            # Add picture to fill the entire slide
            slide.shapes.add_picture(
                temp_path,
                left=Inches(0),
                top=Inches(0),
                width=prs.slide_width,
                height=prs.slide_height
            )
        finally:
            # Clean up temporary file
            os.unlink(temp_path)

    prs.save(output_path)


def process_html_file(html_path: str):
    """
    Process a single HTML file: convert to PDF and PPTX.

    Args:
        html_path: Path to the HTML file or directory containing HTML files
    """
    path = Path(html_path)

    if path.is_dir():
        # If it's a directory, process all .html files in it
        html_files = list(path.glob("*.html"))
        if not html_files:
            print(f"Warning: No .html files found in directory {html_path}")
            return
        for html_file in html_files:
            process_single_file(html_file)
    else:
        # It's a file
        process_single_file(path)


def process_single_file(html_path: Path):
    """
    Process a single HTML file.

    Args:
        html_path: Path object to the HTML file
    """
    if not html_path.exists():
        print(f"Error: File {html_path} does not exist")
        return

    print(f"Processing {html_path}...")

    # Create output directory if it doesn't exist
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)

    # Generate output paths in the output directory
    base_name = html_path.stem
    pdf_path = output_dir / f"{base_name}.pdf"
    pptx_path = output_dir / f"{base_name}.pptx"

    try:
        # Convert HTML to images
        images = convert_html_to_images(str(html_path))

        if not images:
            print(f"Warning: No slides found in {html_path}")
            return

        print(f"Found {len(images)} slides")

        # Create PDF
        images_to_pdf(images, str(pdf_path))
        print(f"PDF saved to {pdf_path}")

        # Create PPTX
        images_to_pptx(images, str(pptx_path))
        print(f"PPTX saved to {pptx_path}")

    except Exception as e:
        print(f"Error processing {html_path}: {e}")


def main():
    parser = argparse.ArgumentParser(
        description="Convert HTML presentations to PDF and PPTX formats"
    )
    parser.add_argument(
        'html_files',
        nargs='+',
        help='Paths to HTML presentation files'
    )

    args = parser.parse_args()

    for html_file in args.html_files:
        process_html_file(html_file)


if __name__ == "__main__":
    main()