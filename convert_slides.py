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
import re
import tempfile
import time
import unicodedata
from importlib import util
from pathlib import Path
from textwrap import wrap
from urllib.parse import urlparse

from PIL import Image, ImageDraw, ImageFont
from playwright.sync_api import sync_playwright
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from pptx.util import Inches, Pt


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
        context = browser.new_context(viewport={'width': viewport_size[0], 'height': viewport_size[1], 'deviceScaleFactor': 2})
        page = context.new_page()

        # Open the local HTML file
        file_url = f"file://{os.path.abspath(html_path)}"
        page.goto(file_url)

        # Wait for page to load completely
        page.wait_for_load_state('networkidle')
        page.wait_for_timeout(1000)

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


def is_url(source: str) -> bool:
    parsed = urlparse(source)
    return parsed.scheme in ('http', 'https', 'file')


def detect_html_type(page) -> str:
    """Detect the type of HTML presentation: 'slide' or 'qa' (question-answer)."""
    # Check for traditional slides
    if page.locator('.slide').count() > 0:
        return 'slide'
    # Check for Q&A style presentations (GPU.html, etc.)
    if page.locator('.presentation').count() > 0 or page.locator('.card').count() > 0:
        return 'qa'
    return 'unknown'


def extract_slide_data(source: str) -> list:
    """Extract slide text content and formatting from an HTML presentation loaded in a browser."""
    url = source
    if not is_url(source):
        source_path = Path(source)
        if source_path.exists():
            url = f"file://{os.path.abspath(source)}"
        else:
            raise ValueError(f"Source not found: {source}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 1920, 'height': 1080})
        page = context.new_page()
        page.goto(url)
        page.wait_for_load_state('networkidle')

        # Detect HTML type: 'slide' or 'qa' (question-answer)
        has_slides = page.locator('.slide').count() > 0
        has_cards = page.locator('.card').count() > 0
        has_presentations = page.locator('.presentation').count() > 0
        
        if has_cards or has_presentations:
            # Q&A style presentation (GPU.html, etc.)
            # First, expand all content blocks by adding force-open class
            page.evaluate('''() => {
                // Force all presentations to be visible
                document.querySelectorAll('.presentation').forEach(pres => {
                    pres.style.display = 'block';
                    pres.style.visibility = 'visible';
                    pres.style.opacity = '1';
                });

                // Add force-open class to all content blocks
                document.querySelectorAll('.content-block').forEach(block => {
                    block.classList.add('force-open');
                });
            }''')
            page.wait_for_timeout(1000)  # Wait for animations to complete
            
            # Now extract the expanded content
            slide_data = page.evaluate(r'''() => {
                const result = [];
                
                // Get all presentation sections
                const presentations = Array.from(document.querySelectorAll('.presentation'));
                
                presentations.forEach((pres, presIdx) => {
                    // Get presentation title
                    const presTitle = pres.querySelector('h2');
                    const presTitleText = presTitle ? presTitle.textContent.trim() : `Presentation ${presIdx + 1}`;
                    
                    // Create title slide for each presentation
                    result.push({
                        rect: { width: 1920, height: 1080 },
                        backgroundColor: '#ffffff',
                        blocks: [{
                            rect: { x: 50, y: 400, width: 1820, height: 280 },
                            type: 'header',
                            elements: [{
                                tag: 'h1',
                                align: 'center',
                                list: false,
                                segments: [{ text: presTitleText, fontSize: '32px', fontWeight: '700', color: '#1e40af' }]
                            }],
                            backgroundColor: '#f8fafc',
                            borderLeftWidth: 0,
                            borderLeftColor: null,
                            borderBottomWidth: 4,
                            borderBottomColor: '#2563eb'
                        }]
                    });
                    
                    // Get all cards in this presentation
                    const cards = Array.from(pres.querySelectorAll('.card'));
                    
                    cards.forEach((card, cardIdx) => {
                        const cardBlocks = [];
                        
                        // Get card header (question)
                        const header = card.querySelector('.card-header');
                        if (header) {
                            cardBlocks.push({
                                rect: { x: 50, y: 30, width: 1820, height: 100 },
                                type: 'header',
                                elements: [{
                                    tag: 'h2',
                                    align: 'left',
                                    list: false,
                                    segments: [{ text: header.textContent.trim(), fontSize: '18px', fontWeight: '700', color: '#0f172a' }]
                                }],
                                backgroundColor: '#f1f5f9',
                                borderLeftWidth: 0,
                                borderLeftColor: null,
                                borderBottomWidth: 3,
                                borderBottomColor: '#2563eb'
                            });
                        }
                        
                        // Get all content blocks (theory, task, solution)
                        const contentBlocks = Array.from(card.querySelectorAll('.content-block'));
                        
                        contentBlocks.forEach((block, blockIdx) => {
                            // Check if block has force-open class or is visible
                            const hasForceOpen = block.classList.contains('force-open');
                            const style = window.getComputedStyle(block);
                            if (!hasForceOpen && style.display === 'none' && block.style.display !== 'block') return;
                            
                            const sectionTitle = block.querySelector('.section-title');
                            const content = block.querySelector('p');
                            const solutionDiv = block.querySelector('.solution');
                            
                            let titleText = '';
                            let contentText = '';
                            let tag = 'p';
                            let bgColor = '#ffffff';
                            let borderColor = '#667eea';
                            
                            if (sectionTitle) {
                                titleText = sectionTitle.textContent.trim();
                            }
                            
                            if (solutionDiv) {
                                contentText = solutionDiv.textContent.trim();
                                tag = 'h3';
                                bgColor = '#ecfdf5';
                                borderColor = '#10b981';
                            } else if (content) {
                                contentText = content.textContent.trim();
                                if (titleText.includes('Теория') || titleText.includes('Конспект')) {
                                    tag = 'h3';
                                    bgColor = '#f0f9ff';
                                    borderColor = '#2563eb';
                                } else if (titleText.includes('Задача') || titleText.includes('Практич')) {
                                    tag = 'h3';
                                    bgColor = '#fffbeb';
                                    borderColor = '#f59e0b';
                                }
                            }
                            
                            if (contentText && titleText) {
                                // Create separate slide for each content block
                                const paragraphs = contentText.split(/\n\n|\n/).filter(p => p.trim());
                                const elements = [];
                                
                                // Add section title
                                elements.push({
                                    tag: 'h3',
                                    align: 'left',
                                    list: false,
                                    segments: [{ text: titleText, fontSize: '16px', fontWeight: '600', color: '#1e40af' }]
                                });
                                
                                // Add content paragraphs
                                paragraphs.forEach((para, paraIdx) => {
                                    elements.push({
                                        tag: 'p',
                                        align: 'left',
                                        list: para.match(/^\d+[\)\.]\s/) !== null,
                                        segments: [{ 
                                            text: para.trim(), 
                                            fontSize: '14px', 
                                            fontWeight: '400', 
                                            color: '#334155' 
                                        }]
                                    });
                                });
                                
                                // Create separate slide for this content block
                                result.push({
                                    rect: { width: 1920, height: 1080 },
                                    backgroundColor: '#ffffff',
                                    blocks: [{
                                        rect: { x: 50, y: 30, width: 1820, height: Math.min(1020, Math.max(200, elements.length * 70)) },
                                        type: 'content',
                                        elements: elements,
                                        backgroundColor: bgColor,
                                        borderLeftWidth: 5,
                                        borderLeftColor: borderColor
                                    }]
                                });
                            }
                        });
                    });
                });
                
                return result;
            }''')
        else:
            # Traditional slide presentation
            slide_data = page.evaluate(r'''() => {
            const wrapper = document.querySelector('.presentation-wrapper') || document.body;
            const slides = Array.from(document.querySelectorAll('.slide'));
            const originalClasses = slides.map(slide => slide.className);
            const originalStyles = slides.map(slide => ({
                display: slide.style.display || '',
                visibility: slide.style.visibility || '',
                opacity: slide.style.opacity || '',
            }));

            function normalize(text) {
                return text.replace(/\u00a0/g, ' ').replace(/\\s+/g, ' ').trim();
            }

            function colorToHex(color) {
                if (!color) return null;
                const rgba = color.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/);
                if (!rgba) return null;
                const alpha = rgba[4] !== undefined ? parseFloat(rgba[4]) : 1;
                if (alpha === 0) return null;
                return '#' + [1,2,3].map(i => Number(rgba[i]).toString(16).padStart(2, '0')).join('');
            }

            function getBackgroundColor(el) {
                const style = window.getComputedStyle(el);
                if (style.backgroundImage && style.backgroundImage !== 'none') {
                    const gradient = style.backgroundImage;
                    const stopMatch = gradient.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
                    if (stopMatch) {
                        return '#' + [1,2,3].map(i => Number(stopMatch[i]).toString(16).padStart(2, '0')).join('');
                    }
                    return null;
                }
                return colorToHex(style.backgroundColor);
            }

            function parseNode(node) {
                const segments = [];
                if (node.nodeType === Node.TEXT_NODE) {
                    const text = normalize(node.nodeValue || '');
                    if (text) {
                        segments.push({ text });
                    }
                    return segments;
                }
                if (node.nodeType !== Node.ELEMENT_NODE) {
                    return segments;
                }

                const style = window.getComputedStyle(node);
                const base = {
                    fontSize: style.fontSize,
                    fontWeight: style.fontWeight,
                    fontStyle: style.fontStyle,
                    color: colorToHex(style.color),
                    underline: style.textDecorationLine.includes('underline'),
                    highlight: node.classList.contains('highlight'),
                };

                if (node.childNodes.length === 0) {
                    const text = normalize(node.innerText || '');
                    if (text) {
                        segments.push(Object.assign({ text }, base));
                    }
                    return segments;
                }

                node.childNodes.forEach(child => {
                    parseNode(child).forEach(segment => {
                        segments.push(Object.assign({}, base, segment));
                    });
                });
                return segments;
            }

            function extractBlock(el, slideRect) {
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                const tags = ['h1','h2','h3','h4','h5','h6','p','li'];
                const elements = [];
                if (tags.includes(el.tagName.toLowerCase())) {
                    elements.push(el);
                }
                Array.from(el.querySelectorAll('h1,h2,h3,h4,h5,h6,p,li')).forEach(sub => {
                    if (!elements.includes(sub)) {
                        elements.push(sub);
                    }
                });
                const parsedElements = elements.map(sub => {
                    const segments = parseNode(sub);
                    return {
                        tag: sub.tagName.toLowerCase(),
                        align: window.getComputedStyle(sub).textAlign,
                        list: sub.tagName.toLowerCase() === 'li',
                        segments,
                    };
                }).filter(item => item.segments.length > 0);

                return {
                    rect: {
                        x: rect.x - slideRect.x,
                        y: rect.y - slideRect.y,
                        width: rect.width,
                        height: rect.height,
                    },
                    type: el.classList.contains('slide-header') || el.classList.contains('header') ? 'header' : 'content',
                    elements: parsedElements,
                    backgroundColor: getBackgroundColor(el),
                    borderLeftColor: colorToHex(style.borderLeftColor),
                    borderLeftWidth: parseFloat(style.borderLeftWidth) || 0,
                    borderBottomColor: colorToHex(style.borderBottomColor),
                    borderBottomWidth: parseFloat(style.borderBottomWidth) || 0,
                };
            }

            function showSlide(slide) {
                slides.forEach(s => {
                    s.style.display = 'none';
                    s.style.visibility = 'hidden';
                    s.style.opacity = '0';
                    s.classList.remove('active');
                });
                slide.style.display = '';
                slide.style.visibility = 'visible';
                slide.style.opacity = '1';
                slide.classList.add('active');
                slide.scrollIntoView({ behavior: 'instant', block: 'center', inline: 'center' });
            }

            function shouldUseBlock(el) {
                return !el.matches('.slide-number, .language-bar, .progress-bar, .nav-btn, .lang-btn') && el.textContent.trim();
            }

            const result = slides.map(slide => {
                showSlide(slide);
                const slideRect = slide.getBoundingClientRect();
                const blocks = [];
                const blockSelectors = [
                        ':scope > .slide-header',
                        ':scope > .header',
                        ':scope > .content-box',
                        ':scope .content-box',
                        ':scope .card',
                        ':scope .card-header',
                        ':scope .card-body',
                        ':scope > .box',
                        ':scope .box',
                        ':scope > .content-ru.active',
                        ':scope > .content-en.active',
                        ':scope > .content-zh.active',
                        ':scope .content-ru.active',
                        ':scope .content-en.active',
                        ':scope .content-zh.active',
                        ':scope > .presentation',
                        ':scope .presentation',
                        ':scope > .content-block',
                        ':scope .content-block',
                        ':scope > h1',
                        ':scope > h2',
                        ':scope > h3',
                        ':scope > h4',
                        ':scope > h5',
                        ':scope > h6',
                        ':scope > p',
                        ':scope > ul',
                        ':scope > ol',
                        ':scope .h1',
                        ':scope .h2',
                        ':scope .h3',
                        ':scope .h4',
                        ':scope .h5',
                        ':scope .h6',
                        ':scope .p',
                        ':scope .ul',
                        ':scope .ol',
                        ':scope .q-grid'
                    ];
                blockSelectors.forEach(selector => {
                    Array.from(slide.querySelectorAll(selector)).forEach(el => {
                        if (shouldUseBlock(el) && !blocks.includes(el)) {
                            blocks.push(el);
                        }
                    });
                });
                if (blocks.length === 0) {
                    Array.from(slide.children).forEach(child => {
                        const style = window.getComputedStyle(child);
                        if (style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0' && child.textContent.trim()) {
                            blocks.push(child);
                        }
                    });
                }
                const blockData = blocks.map(block => extractBlock(block, slideRect)).filter(item => item.elements.length > 0);
                return {
                    rect: {
                        width: slideRect.width,
                        height: slideRect.height,
                    },
                    backgroundColor: getBackgroundColor(slide),
                    blocks: blockData,
                };
            });

            slides.forEach((slide, i) => {
                slide.className = originalClasses[i];
                slide.style.display = originalStyles[i].display;
                slide.style.visibility = originalStyles[i].visibility;
                slide.style.opacity = originalStyles[i].opacity;
            });
            return result;
        }''')

        browser.close()

    if not slide_data:
        raise ValueError('No slides found in the HTML file')
    return slide_data


def slide_data_to_pptx(slide_data: list, output_path: str, slide_size: tuple = (16, 9)):
    """Create an editable PPTX from extracted slide text data."""
    if not slide_data:
        raise ValueError('No slide data provided for PPTX creation')

    prs = Presentation()
    prs.slide_width = Inches(slide_size[0])
    prs.slide_height = Inches(slide_size[1])

    def px_to_pt(px_value) -> Pt:
        try:
            if isinstance(px_value, (int, float)):
                return Pt(float(px_value) * 0.75)
            if isinstance(px_value, str):
                if px_value.endswith('px'):
                    return Pt(float(px_value[:-2]) * 0.75)
                if px_value.endswith('pt'):
                    return Pt(float(px_value[:-2]))
        except Exception:
            pass
        return Pt(14)

    def hex_to_rgb(hex_color: str) -> RGBColor:
        if not hex_color:
            return RGBColor(0, 0, 0)
        hex_color = hex_color.lstrip('#')
        if len(hex_color) == 6:
            return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))
        return RGBColor(0, 0, 0)

    def build_text_shape(left, top, width, height, block):
        bg_color = block.get('backgroundColor')
        if block.get('type') == 'content' and bg_color:
            shape = prs_slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
            try:
                shape.roundness = 0.08
            except Exception:
                pass
            shape.fill.solid()
            shape.fill.fore_color.rgb = hex_to_rgb(bg_color)
            shape.line.fill.background()
            if block.get('borderLeftWidth', 0) > 0 and block.get('borderLeftColor'):
                border_width = px_to_pt(block['borderLeftWidth'])
                border_shape = prs_slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    left,
                    top,
                    border_width,
                    height
                )
                border_shape.fill.solid()
                border_shape.fill.fore_color.rgb = hex_to_rgb(block['borderLeftColor'])
                border_shape.line.fill.background()
            return shape

        text_box = prs_slide.shapes.add_textbox(left, top, width, height)
        text_box.fill.background()
        return text_box

    def apply_paragraph_style(paragraph, element, block_type):
        paragraph.alignment = {
            'center': PP_PARAGRAPH_ALIGNMENT.CENTER,
            'right': PP_PARAGRAPH_ALIGNMENT.RIGHT,
            'left': PP_PARAGRAPH_ALIGNMENT.LEFT,
            'justify': PP_PARAGRAPH_ALIGNMENT.JUSTIFY,
        }.get(element.get('align', 'left'), PP_PARAGRAPH_ALIGNMENT.LEFT)
        if element.get('list'):
            paragraph.level = 0
            paragraph.bullet = True
        tag = element.get('tag', '')
        if tag in ('h1', 'h2'):
            paragraph.space_before = Pt(6)
            paragraph.space_after = Pt(8)
        elif tag in ('h3', 'h4', 'h5', 'h6'):
            paragraph.space_before = Pt(4)
            paragraph.space_after = Pt(6)
        elif block_type == 'header':
            paragraph.space_after = Pt(8)
        else:
            paragraph.space_after = Pt(5)
        paragraph.line_spacing = 1.2

    def apply_run_style(run, segment, bg_color=None):
        run.font.name = 'Segoe UI'
        run.font.size = px_to_pt(segment.get('fontSize', '14px'))
        font_weight = segment.get('fontWeight', '')
        run.font.bold = font_weight in ('700', '800', '900', 'bold') or (isinstance(font_weight, str) and font_weight.isdigit() and int(font_weight) >= 600)
        run.font.italic = segment.get('fontStyle') == 'italic'
        
        # 确定文本颜色
        text_color = None
        if bg_color and bg_color.startswith('#') and len(bg_color) == 7:
            # 计算背景亮度
            r = int(bg_color[1:3], 16)
            g = int(bg_color[3:5], 16)
            b = int(bg_color[5:7], 16)
            brightness = 0.299 * r + 0.587 * g + 0.114 * b
            if brightness < 128:  # 深色背景
                text_color = RGBColor(255, 255, 255)  # 白色文本
            else:
                text_color = hex_to_rgb(segment.get('color')) if segment.get('color') else RGBColor(0, 0, 0)
        else:
            text_color = hex_to_rgb(segment.get('color')) if segment.get('color') else RGBColor(0, 0, 0)
        
        if text_color:
            run.font.color.rgb = text_color
        
        if segment.get('underline'):
            run.font.underline = True
        if segment.get('highlight'):
            run.font.bold = True
            run.font.color.rgb = hex_to_rgb('#667eea')
            run.font.highlight_color = MSO_THEME_COLOR.ACCENT_1

    for slide in slide_data:
        prs_slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide_width = slide['rect']['width'] or 1920
        slide_height = slide['rect']['height'] or 1080

        background_color = slide.get('backgroundColor')
        if background_color:
            fill = prs_slide.background.fill
            fill.solid()
            fill.fore_color.rgb = hex_to_rgb(background_color)

        for block in slide['blocks']:
            left = int(block['rect']['x'] / slide_width * prs.slide_width)
            top = int(block['rect']['y'] / slide_height * prs.slide_height)
            width = int(block['rect']['width'] / slide_width * prs.slide_width)
            height = int(block['rect']['height'] / slide_height * prs.slide_height)
            if width < Inches(1):
                width = Inches(1)
            if height < Inches(0.5):
                height = Inches(0.5)

            text_box = build_text_shape(left, top, width, height, block)
            text_frame = text_box.text_frame
            text_frame.word_wrap = True
            text_frame.margin_left = Inches(0.8)
            text_frame.margin_right = Inches(0.8)
            text_frame.margin_top = Inches(0.6)
            text_frame.margin_bottom = Inches(0.6)

            if block.get('type') == 'header' and block.get('borderBottomWidth', 0) > 0:
                bottom_width = px_to_pt(block['borderBottomWidth'])
                border_rect = prs_slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    left,
                    top + height - bottom_width,
                    width,
                    bottom_width
                )
                border_rect.fill.solid()
                border_rect.fill.fore_color.rgb = hex_to_rgb(block['borderBottomColor'] or '#667eea')
                border_rect.line.fill.background()

            for element_index, element in enumerate(block['elements']):
                paragraph = text_frame.paragraphs[0] if element_index == 0 else text_frame.add_paragraph()
                paragraph.text = ''
                apply_paragraph_style(paragraph, element, block.get('type'))

                for segment in element['segments']:
                    run = paragraph.add_run()
                    run.text = segment.get('text', '')
                    apply_run_style(run, segment, block.get('backgroundColor'))

            if len(block['elements']) == 0:
                text_frame.text = ''

    prs.save(output_path)


def overlay_text_on_image(image: Image.Image, text: str, position: str = 'bottom', padding: int = 20, font_size: int = 32) -> Image.Image:
    """Overlay text on a PIL Image and return a new image."""
    image = image.copy()
    draw = ImageDraw.Draw(image)
    font = ImageFont.load_default()

    if hasattr(font, 'getsize'):
        char_width, char_height = font.getsize('A')
    else:
        char_width, char_height = (8, 16)

    max_chars = max(20, (image.width - padding * 2) // max(char_width, 1))
    lines = []
    for raw_line in text.splitlines() or ['']:
        if raw_line.strip() == '':
            lines.append('')
        else:
            lines.extend(wrap(raw_line, width=max_chars))

    line_height = char_height + 4
    text_height = line_height * len(lines)
    box_top = padding if position == 'top' else image.height - padding - text_height - padding
    box_bottom = box_top + text_height + padding
    box_left = padding
    box_right = image.width - padding

    draw.rectangle(
        [(box_left - 10, box_top - 10), (box_right + 10, box_bottom + 10)],
        fill=(0, 0, 0)
    )

    y = box_top
    for line in lines:
        draw.text((padding, y), line, font=font, fill=(255, 255, 255))
        y += line_height

    return image


def images_to_pdf(images: list, output_path: str, overlay_text: str | None = None, text_position: str = 'bottom'):
    """
    Convert a list of PIL Images to a single PDF file.

    Args:
        images: List of PIL Image objects
        output_path: Path to save the PDF
        overlay_text: Optional text to overlay on each slide image
        text_position: Position for overlay text on slide images
    """
    if not images:
        raise ValueError("No images provided for PDF creation")

    if overlay_text:
        images = [overlay_text_on_image(img, overlay_text, position=text_position) for img in images]

    # Convert all images to RGB mode for PDF compatibility
    rgb_images = [img.convert('RGB') for img in images]

    # Save as PDF
    rgb_images[0].save(output_path, save_all=True, append_images=rgb_images[1:])


def images_to_pptx(images: list, output_path: str, slide_size: tuple = (16, 9), overlay_text: str | None = None, text_position: str = 'bottom'):
    """
    Convert a list of PIL Images to a PPTX presentation.

    Args:
        images: List of PIL Image objects
        output_path: Path to save the PPTX
        slide_size: Tuple of (width, height) in inches for slides
        overlay_text: Optional text to add as a slide text box
        text_position: Position for overlay text on slides
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

        if overlay_text:
            text_box = slide.shapes.add_textbox(
                Inches(0.5),
                Inches(0.5) if text_position == 'top' else prs.slide_height - Inches(1.5),
                prs.slide_width - Inches(1.0),
                Inches(1.0)
            )
            text_frame = text_box.text_frame
            text_frame.text = overlay_text

    prs.save(output_path)


def sanitize_filename(name: str) -> str:
    """Return an ASCII-safe base filename by normalizing and stripping unsupported characters."""
    name = unicodedata.normalize('NFKD', name)
    name = name.encode('ascii', 'ignore').decode('ascii')
    name = name.strip()
    name = re.sub(r'[^A-Za-z0-9._-]+', '_', name)
    name = name.strip('._-')
    return name or 'output'


def process_html_file(html_path: str, overlay_text: str | None = None, text_position: str = 'bottom', editable: bool = False, no_pdf: bool = False):
    """
    Process a single HTML file or URL: convert to PDF and/or PPTX.

    Args:
        html_path: Path to the HTML file, directory, or URL
        overlay_text: Optional text to add to each slide
        text_position: Position for overlay text on slides
        editable: Whether to create an editable PPTX from slide text instead of image-based slides
        no_pdf: Whether to skip PDF generation
    """
    path = Path(html_path)

    if path.is_dir():
        html_files = list(path.glob("*.html"))
        if not html_files:
            print(f"Warning: No .html files found in directory {html_path}")
            return
        for html_file in html_files:
            process_single_file(html_file, overlay_text=overlay_text, text_position=text_position, editable=editable, no_pdf=no_pdf)
    else:
        process_single_file(path, overlay_text=overlay_text, text_position=text_position, editable=editable, no_pdf=no_pdf)


def process_single_file(html_path: Path, overlay_text: str | None = None, text_position: str = 'bottom', editable: bool = False, no_pdf: bool = False):
    """
    Process a single HTML file.

    Args:
        html_path: Path object to the HTML file or URL
        overlay_text: Optional text to add to each slide
        text_position: Position for overlay text on slides
        editable: Whether to create an editable PPTX from slide text instead of image-based slides
        no_pdf: Whether to skip PDF generation
    """
    source = str(html_path)
    if html_path.exists() or is_url(source):
        print(f"Processing {source}...")
    else:
        print(f"Error: Source {source} does not exist")
        return

    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)

    if is_url(source):
        url_path = urlparse(source).path
        base_name = sanitize_filename(Path(url_path).stem or 'output')
    else:
        base_name = sanitize_filename(html_path.stem)
    pdf_path = output_dir / f"{base_name}.pdf"
    pptx_path = output_dir / f"{base_name}.pptx"

    try:
        if editable:
            slide_data = extract_slide_data(source)
            slide_data_to_pptx(slide_data, str(pptx_path))
            print(f"Editable PPTX saved to {pptx_path}")
            return

        images = convert_html_to_images(source)
        if not images:
            print(f"Warning: No slides found in {source}")
            return

        print(f"Found {len(images)} slides")

        if not no_pdf:
            images_to_pdf(images, str(pdf_path), overlay_text=overlay_text, text_position=text_position)
            print(f"PDF saved to {pdf_path}")

        images_to_pptx(images, str(pptx_path), overlay_text=overlay_text, text_position=text_position)
        print(f"PPTX saved to {pptx_path}")

    except Exception as e:
        print(f"Error processing {source}: {e}")


def get_clipboard_text() -> str:
    pyperclip_spec = util.find_spec('pyperclip')
    if pyperclip_spec is None:
        raise RuntimeError("pyperclip is not installed. Install it with: pip install pyperclip")

    pyperclip = __import__('pyperclip')
    clipboard_text = pyperclip.paste()
    if not clipboard_text:
        raise RuntimeError("Clipboard is empty or contains no text")
    return clipboard_text


def main():
    parser = argparse.ArgumentParser(
        description="Convert HTML presentations to PDF and PPTX formats"
    )
    parser.add_argument(
        'html_files',
        nargs='+',
        help='Paths or URLs to HTML presentation files'
    )

    group = parser.add_mutually_exclusive_group()
    group.add_argument('--text', help='Text to add to each slide')
    group.add_argument('--text-file', help='Load text from a file to add to each slide')
    group.add_argument('--clipboard', action='store_true', help='Read text from clipboard to add to each slide')
    parser.add_argument('--position', choices=['top', 'bottom'], default='bottom', help='Position for overlay text on slides')
    parser.add_argument('--editable', action='store_true', help='Create editable PPTX from slide text instead of image slides')
    parser.add_argument('--no-pdf', action='store_true', help='Skip PDF generation')

    args = parser.parse_args()

    overlay_text = None
    if args.text_file:
        with open(args.text_file, 'r', encoding='utf-8') as f:
            overlay_text = f.read()
    elif args.clipboard:
        overlay_text = get_clipboard_text()
    elif args.text:
        overlay_text = args.text

    for html_file in args.html_files:
        process_html_file(
            html_file,
            overlay_text=overlay_text,
            text_position=args.position,
            editable=args.editable,
            no_pdf=args.no_pdf
        )


if __name__ == "__main__":
    main()