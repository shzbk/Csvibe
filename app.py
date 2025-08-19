import streamlit as st
import pandas as pd
import os
import tempfile
import csv
import time
import zipfile
import shutil
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus.flowables import Flowable

# PDF to PNG conversion imports
try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

try:
    from pdf2image import convert_from_path
    PDF2IMAGE_AVAILABLE = True
except ImportError:
    PDF2IMAGE_AVAILABLE = False

# Configure Streamlit page
st.set_page_config(
    page_title="CSVibe - Dictionary Generator",
    page_icon="📄",
    layout="wide"
)


# Line flowable for horizontal lines
class LineFlowable(Flowable):
    def __init__(self, width, color=(0, 0, 0), height=6, line_width=4):
        Flowable.__init__(self)
        self.width = width
        self.height = height
        self.color = color
        self.line_width = line_width

    def draw(self):
        from reportlab.lib.colors import Color
        self.canv.setLineWidth(self.line_width)
        self.canv.setStrokeColor(Color(*self.color))
        self.canv.line(0, 0, self.width, 0)

# PDF Generation Function
def create_pdf_from_csv(csv_file, output_file, 
                        term_font=None, term_size=84, term_spacing=48,
                        pronunciation_font=None, pronunciation_size=28, pronunciation_spacing=36,
                        definition_font=None, definition_size=28, page_color="#FFFFFF",
                        term_color="#000000", pronunciation_color="#000000", line_color="#000000", definition_color="#000000",
                        page_width_inches=11, page_height_inches=14, text_alignment="left", page_position="bottom-left"):
    """Create PDF with customizable fonts and sizes"""
    
    def register_font_safe(font_path, base_name):
        """Safely register a font with unique name to avoid conflicts"""
        if font_path and font_path.endswith('.ttf') and os.path.exists(font_path):
            try:
                # Use timestamp to create unique font name
                unique_name = f"{base_name}_{int(time.time() * 1000)}"
                pdfmetrics.registerFont(TTFont(unique_name, font_path))
                return unique_name
            except Exception as e:
                st.warning(f"Failed to register {font_path}: {e}")
                return 'Times-Bold'
        elif font_path in ['Times-Bold', 'Helvetica-Bold', 'Helvetica', 'Times-Roman', 'Courier', 'Courier-Bold']:
            return font_path
        else:
            return 'Times-Bold'
    
    # Register fonts with unique names
    title_font = register_font_safe(term_font, 'TermFont')
    pronunciation_font_name = register_font_safe(pronunciation_font, 'PronunciationFont')
    definition_font_name = register_font_safe(definition_font, 'DefinitionFont')
    
    # Enhanced Unicode support for built-in fonts
    unicode_font = 'Helvetica'
    unicode_font_bold = 'Helvetica-Bold'
    
    if pronunciation_font_name in ['Helvetica-Bold', 'Helvetica'] or definition_font_name in ['Helvetica-Bold', 'Helvetica']:
        try:
            font_configs = [
                ('C:/Windows/Fonts/segoeui.ttf', 'C:/Windows/Fonts/segoeuib.ttf'),
                ('/mnt/c/Windows/Fonts/segoeui.ttf', '/mnt/c/Windows/Fonts/segoeuib.ttf'),
            ]
            
            for regular_path, bold_path in font_configs:
                if os.path.exists(regular_path):
                    try:
                        unicode_name = f"Unicode_{int(time.time() * 1000)}"
                        unicode_bold_name = f"UnicodeBold_{int(time.time() * 1000)}"
                        
                        pdfmetrics.registerFont(TTFont(unicode_name, regular_path))
                        unicode_font = unicode_name
                        
                        if os.path.exists(bold_path):
                            pdfmetrics.registerFont(TTFont(unicode_bold_name, bold_path))
                            unicode_font_bold = unicode_bold_name
                        
                        break
                    except Exception:
                        continue
        except Exception:
            pass
    
    # Update built-in font references to Unicode versions
    if pronunciation_font_name == 'Helvetica-Bold':
        pronunciation_font_name = unicode_font_bold
    elif pronunciation_font_name == 'Helvetica':
        pronunciation_font_name = unicode_font
    
    if definition_font_name == 'Helvetica-Bold':
        definition_font_name = unicode_font_bold
    elif definition_font_name == 'Helvetica':
        definition_font_name = unicode_font
    
    # Page setup
    page_width = page_width_inches * inch
    page_height = page_height_inches * inch
    page_size = (page_width, page_height)
    
    # Calculate scaling factors based on 11x14 baseline
    baseline_width = 11.0
    baseline_height = 14.0
    width_scale = page_width_inches / baseline_width
    height_scale = page_height_inches / baseline_height
    # Use average scale for consistent proportions
    scale_factor = (width_scale + height_scale) / 2
    
    # Scale all size and spacing values
    scaled_term_size = int(term_size * scale_factor)
    scaled_term_spacing = int(term_spacing * scale_factor)
    scaled_pronunciation_size = int(pronunciation_size * scale_factor)
    scaled_pronunciation_spacing = int(pronunciation_spacing * scale_factor)
    scaled_definition_size = int(definition_size * scale_factor)
    
    # Scale margins proportionally
    scaled_margin = 1.5 * scale_factor * inch
    
    # Scale line width and positioning
    scaled_line_width = int(4 * scale_factor)
    scaled_line_length = page_width - 2*scaled_margin  # Match frame content width
    scaled_definition_space_before = int(54 * scale_factor)
    
    # Content-aware positioning functions
    def calculate_content_height(row, title_style, pronunciation_style, definition_style, available_width):
        """Calculate the total height needed for all content elements"""
        try:
            # Create temporary paragraphs to measure their height
            term_para = Paragraph(row['term'].lower(), title_style)
            pronunciation_text = f"{row['pronunciation']} • ({row['type']})"
            pronunciation_para = Paragraph(pronunciation_text, pronunciation_style)
            definition_para = Paragraph(row['definition'], definition_style)
            
            # Measure actual heights using ReportLab's wrap method
            # Use a reasonable max height to force proper wrapping calculation
            max_height = 20 * inch  # Large enough to not constrain wrapping
            
            term_height = term_para.wrap(available_width, max_height)[1]
            pronunciation_height = pronunciation_para.wrap(available_width, max_height)[1]
            definition_height = definition_para.wrap(available_width, max_height)[1]
            
            # Add spacing between elements including all style spacing
            total_height = (term_height + scaled_term_spacing + 
                          pronunciation_height + scaled_pronunciation_spacing + 
                          scaled_line_width + scaled_definition_space_before + 
                          definition_height)
            
            # Font-specific adjustments - different fonts need different spacing
            font_padding = 0
            
            # Check for custom TTF fonts (they often have unpredictable metrics)
            if hasattr(title_style, 'fontName') and ('_' in str(title_style.fontName) or any(char.isdigit() for char in str(title_style.fontName)[-10:])):
                font_padding += scaled_term_size * 0.20  # Increased from 15% to 20%
                
            if hasattr(pronunciation_style, 'fontName') and ('_' in str(pronunciation_style.fontName) or any(char.isdigit() for char in str(pronunciation_style.fontName)[-10:])):
                font_padding += scaled_pronunciation_size * 0.20  # Increased from 15% to 20%
                
            if hasattr(definition_style, 'fontName') and ('_' in str(definition_style.fontName) or any(char.isdigit() for char in str(definition_style.fontName)[-10:])):
                font_padding += scaled_definition_size * 0.20  # Increased from 15% to 20%
            
            total_height += font_padding
            
            # Add extra spacing that ReportLab applies (leading adjustments, etc.)
            extra_leading = int(12 * scale_factor)  # From definition_style leading
            pronunciation_leading = int(4 * scale_factor)  # From pronunciation_style leading  
            
            total_height += extra_leading + pronunciation_leading
            
            # ReportLab can add unexpected spacing for long text blocks
            # Add extra padding for multi-line content (the sneaky overflow culprit)
            definition_line_count = max(1, len(row['definition']) // 80)  # Rough estimate of lines
            if definition_line_count > 2:  # Multi-line definitions need extra space
                multiline_padding = definition_line_count * int(8 * scale_factor)
                total_height += multiline_padding
            
            # Character-specific adjustments for special characters that can affect height
            text_content = f"{row['term']} {row['pronunciation']} {row['definition']}"
            if any(ord(char) > 127 for char in text_content):  # Non-ASCII characters
                unicode_padding = int(6 * scale_factor)
                total_height += unicode_padding
            
            # Increased safety margin from 5% to 10% to catch edge cases
            safety_margin = total_height * 0.10
                
            return total_height + safety_margin
            
        except Exception as e:
            # Fallback to conservative estimate if measurement fails
            import streamlit as st
            st.warning(f"Height calculation failed, using conservative estimate: {e}")
            return 8 * scale_factor * inch  # Conservative fallback
    
    def get_end_positioned_spacer_amount(position, page_h, scaled_margin, content_height):
        """Your way: content ENDS at consistent margins, not starts"""
        content_area_height = page_h - 2 * scaled_margin
        
        if position == "top":
            return 0  # Start at top margin, flow down (same as before)
        elif position == "middle":
            # Center the content block in the page
            spacer = (content_area_height - content_height) / 2
            return max(0, spacer)  # Prevent negative spacer if content is too large
        elif position == "bottom":
            # END at bottom margin - calculate where to START so it ends there
            # Just like top starts at top margin, bottom ENDS at bottom margin
            spacer = content_area_height - content_height
            # Add extra conservative margin for bottom positioning (most problematic)
            conservative_bottom_margin = int(20 * scale_factor)
            spacer = spacer - conservative_bottom_margin
            return max(0, spacer)  # Prevent negative spacer if content is too large
        else:
            return 0  # fallback to top
    
    # Get alignment constant
    def get_alignment_constant(alignment):
        from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
        alignments = {
            "left": TA_LEFT,
            "center": TA_CENTER, 
            "right": TA_RIGHT
        }
        return alignments.get(alignment, TA_LEFT)
    
    text_align_const = get_alignment_constant(text_alignment)
    # spacer_amount will be calculated per-row based on content
    
    # Convert hex color to RGB for ReportLab
    def hex_to_rgb(hex_color):
        """Convert hex color to RGB tuple (0-1 range)"""
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i+2], 16)/255.0 for i in (0, 2, 4))
    
    bg_color = hex_to_rgb(page_color)
    term_rgb = hex_to_rgb(term_color)
    pronunciation_rgb = hex_to_rgb(pronunciation_color)
    line_rgb = hex_to_rgb(line_color)
    definition_rgb = hex_to_rgb(definition_color)
    
    doc = SimpleDocTemplate(output_file, pagesize=page_size,
                          rightMargin=scaled_margin, leftMargin=scaled_margin,
                          topMargin=scaled_margin, bottomMargin=scaled_margin)
    
    # Create custom page template with background color
    from reportlab.platypus import PageTemplate, Frame
    from reportlab.lib.colors import Color
    
    class ColoredPageTemplate(PageTemplate):
        def __init__(self, id, frames, bg_color, **kwargs):
            super().__init__(id, frames, **kwargs)
            self.bg_color = bg_color
        
        def beforeDrawPage(self, canvas, doc):
            # Fill entire page with background color
            canvas.saveState()
            canvas.setFillColor(Color(*self.bg_color))
            canvas.rect(0, 0, page_width, page_height, fill=1, stroke=0)
            canvas.restoreState()
    
    # Create frame for content (same margins as SimpleDocTemplate)
    frame = Frame(scaled_margin, scaled_margin, 
                  page_width - 2*scaled_margin, page_height - 2*scaled_margin,
                  id='normal')
    
    # Create custom document with colored background
    from reportlab.platypus import BaseDocTemplate
    doc = BaseDocTemplate(output_file, pagesize=page_size)
    
    # Add the colored page template
    colored_template = ColoredPageTemplate('colored', [frame], bg_color)
    doc.addPageTemplates([colored_template])
    
    story = []
    styles = getSampleStyleSheet()
    
    # Custom styles using user-selected fonts, sizes, and colors
    from reportlab.lib.colors import Color
    
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=scaled_term_size,
        spaceAfter=scaled_term_spacing,
        alignment=text_align_const,
        fontName=title_font,
        leading=scaled_term_size,
        textColor=Color(*term_rgb)
    )
    
    pronunciation_style = ParagraphStyle(
        'Pronunciation',
        parent=styles['Normal'],
        fontSize=scaled_pronunciation_size,
        spaceAfter=scaled_pronunciation_spacing,
        alignment=text_align_const,
        fontName=pronunciation_font_name,
        leading=scaled_pronunciation_size + int(4 * scale_factor),
        textColor=Color(*pronunciation_rgb)
    )
    
    definition_style = ParagraphStyle(
        'Definition',
        parent=styles['Normal'],
        fontSize=scaled_definition_size,
        spaceBefore=scaled_definition_space_before,
        alignment=text_align_const,
        fontName=definition_font_name,
        leading=scaled_definition_size + int(12 * scale_factor),
        textColor=Color(*definition_rgb)
    )
    
    # Read CSV and create pages
    with open(csv_file, 'r', encoding='utf-8-sig') as file:
        reader = csv.DictReader(file)
        first_page = True
        available_width = page_width - 2 * scaled_margin
        
        for row in reader:
            if not first_page:
                story.append(PageBreak())
            first_page = False
            
            # Calculate basic content height (simple measurement, no complex safety margins)
            content_height = calculate_content_height(row, title_style, pronunciation_style, 
                                                    definition_style, available_width)
            
            # Your way: content ENDS at consistent margins
            spacer_amount = get_end_positioned_spacer_amount(page_position, page_height, 
                                                           scaled_margin, content_height)
            
            # Add spacing for positioning (from top)
            story.append(Spacer(1, spacer_amount))
            
            # Title (lowercase)
            title = Paragraph(row['term'].lower(), title_style)
            story.append(title)
            
            # Pronunciation with type
            pronunciation_text = f"{row['pronunciation']} • ({row['type']})"
            pronunciation = Paragraph(pronunciation_text, pronunciation_style)
            story.append(pronunciation)
            
            # Horizontal line (scaled)
            story.append(LineFlowable(scaled_line_length, line_rgb, line_width=scaled_line_width))
            
            # Definition
            definition = Paragraph(row['definition'], definition_style)
            story.append(definition)
    
    doc.build(story)
    return True

# Quote Generation Function
def create_quotes_pdf_from_csv(csv_file, output_file, 
                              quote_font=None, quote_size=48, 
                              page_color="#FFFFFF", quote_color="#000000",
                              page_width_inches=11, page_height_inches=14, text_alignment="center", page_position="middle"):
    """Create PDF with quotes from CSV file (one quote per line)"""
    
    def register_font_safe(font_path, base_name):
        """Safely register a font with unique name to avoid conflicts"""
        if font_path and font_path.endswith('.ttf') and os.path.exists(font_path):
            try:
                unique_name = f"{base_name}_{int(time.time() * 1000)}"
                pdfmetrics.registerFont(TTFont(unique_name, font_path))
                return unique_name
            except Exception as e:
                st.warning(f"Failed to register {font_path}: {e}")
                return 'Times-Bold'
        elif font_path in ['Times-Bold', 'Helvetica-Bold', 'Helvetica', 'Times-Roman', 'Courier', 'Courier-Bold']:
            return font_path
        else:
            return 'Times-Bold'
    
    # Register font with Unicode support (same as dictionary mode)
    quote_font_name = register_font_safe(quote_font, 'QuoteFont')
    
    # Enhanced Unicode support for built-in fonts (same as dictionary mode)
    unicode_font = 'Helvetica'
    unicode_font_bold = 'Helvetica-Bold'
    
    if quote_font_name in ['Helvetica-Bold', 'Helvetica']:
        try:
            font_configs = [
                ('C:/Windows/Fonts/segoeui.ttf', 'C:/Windows/Fonts/segoeuib.ttf'),
                ('/mnt/c/Windows/Fonts/segoeui.ttf', '/mnt/c/Windows/Fonts/segoeuib.ttf'),
            ]
            
            for regular_path, bold_path in font_configs:
                if os.path.exists(regular_path):
                    try:
                        unicode_name = f"Unicode_{int(time.time() * 1000)}"
                        unicode_bold_name = f"UnicodeBold_{int(time.time() * 1000)}"
                        
                        pdfmetrics.registerFont(TTFont(unicode_name, regular_path))
                        unicode_font = unicode_name
                        
                        if os.path.exists(bold_path):
                            pdfmetrics.registerFont(TTFont(unicode_bold_name, bold_path))
                            unicode_font_bold = unicode_bold_name
                        
                        break
                    except Exception:
                        continue
        except Exception:
            pass
    
    # Update built-in font references to Unicode versions
    if quote_font_name == 'Helvetica-Bold':
        quote_font_name = unicode_font_bold
    elif quote_font_name == 'Helvetica':
        quote_font_name = unicode_font
    
    # Page setup
    page_width = page_width_inches * inch
    page_height = page_height_inches * inch
    page_size = (page_width, page_height)
    
    # Calculate scaling factors
    baseline_width = 11.0
    baseline_height = 14.0
    width_scale = page_width_inches / baseline_width
    height_scale = page_height_inches / baseline_height
    scale_factor = (width_scale + height_scale) / 2
    
    # Scale font size and margins
    scaled_quote_size = int(quote_size * scale_factor)
    scaled_margin = 1.5 * scale_factor * inch  # Same margins as dictionary
    
    # Get alignment constant
    def get_alignment_constant(alignment):
        from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
        alignments = {
            "left": TA_LEFT,
            "center": TA_CENTER, 
            "right": TA_RIGHT
        }
        return alignments.get(alignment, TA_CENTER)  # Default center for quotes
    
    text_align_const = get_alignment_constant(text_alignment)
    
    # Robust content height calculation (same logic as dictionary mode)
    def calculate_quote_height(quote_text, quote_style, available_width):
        """Calculate the total height needed for quote content with proper safety margins"""
        try:
            # Create temporary paragraph to measure its height
            quote_para = Paragraph(quote_text, quote_style)
            
            # Measure actual height using ReportLab's wrap method
            # Use a reasonable max height to force proper wrapping calculation
            max_height = 20 * inch  # Large enough to not constrain wrapping
            
            quote_height = quote_para.wrap(available_width, max_height)[1]
            
            # Font-specific adjustments - different fonts need different spacing
            font_padding = 0
            
            # Check for custom TTF fonts (they often have unpredictable metrics)
            if hasattr(quote_style, 'fontName') and ('_' in str(quote_style.fontName) or any(char.isdigit() for char in str(quote_style.fontName)[-10:])):
                font_padding += scaled_quote_size * 0.20  # Increased from 15% to 20%
            
            quote_height += font_padding
            
            # Add extra spacing that ReportLab applies (leading adjustments, etc.)
            extra_leading = int(12 * scale_factor)  # From quote_style leading
            quote_height += extra_leading
            
            # ReportLab can add unexpected spacing for long text blocks
            # Add extra padding for multi-line quotes (the sneaky overflow culprit)
            quote_line_count = max(1, len(quote_text) // 80)  # Rough estimate of lines
            if quote_line_count > 2:  # Multi-line quotes need extra space
                multiline_padding = quote_line_count * int(8 * scale_factor)
                quote_height += multiline_padding
            
            # Character-specific adjustments for special characters that can affect height
            if any(ord(char) > 127 for char in quote_text):  # Non-ASCII characters
                unicode_padding = int(6 * scale_factor)
                quote_height += unicode_padding
            
            # Increased safety margin from 5% to 10% to catch edge cases
            safety_margin = quote_height * 0.10
                
            return quote_height + safety_margin
            
        except Exception as e:
            # Fallback to conservative estimate if measurement fails
            import streamlit as st
            st.warning(f"Quote height calculation failed, using conservative estimate: {e}")
            return 8 * scale_factor * inch  # Conservative fallback
    
    # Exact same positioning approach as dictionary mode (your way: content ENDS at consistent margins)
    def get_quote_spacer_amount(position, page_h, scaled_margin, content_height):
        """Your way: content ENDS at consistent margins, not starts"""
        content_area_height = page_h - 2 * scaled_margin
        
        if position == "top":
            return 0  # Start at top margin, flow down (same as before)
        elif position == "middle":
            # Center the content block in the page
            spacer = (content_area_height - content_height) / 2
            return max(0, spacer)  # Prevent negative spacer if content is too large
        elif position == "bottom":
            # END at bottom margin - calculate where to START so it ends there
            # Just like top starts at top margin, bottom ENDS at bottom margin
            spacer = content_area_height - content_height
            # Add extra conservative margin for bottom positioning (most problematic)
            conservative_bottom_margin = int(20 * scale_factor)
            spacer = spacer - conservative_bottom_margin
            return max(0, spacer)  # Prevent negative spacer if content is too large
        else:
            return 0  # fallback to top
    
    # Convert colors
    def hex_to_rgb(hex_color):
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i+2], 16)/255.0 for i in (0, 2, 4))
    
    bg_color = hex_to_rgb(page_color)
    quote_rgb = hex_to_rgb(quote_color)
    
    # Create document with colored background
    from reportlab.platypus import PageTemplate, Frame, BaseDocTemplate
    from reportlab.lib.colors import Color
    
    class ColoredPageTemplate(PageTemplate):
        def __init__(self, id, frames, bg_color, **kwargs):
            super().__init__(id, frames, **kwargs)
            self.bg_color = bg_color
        
        def beforeDrawPage(self, canvas, doc):
            canvas.saveState()
            canvas.setFillColor(Color(*self.bg_color))
            canvas.rect(0, 0, page_width, page_height, fill=1, stroke=0)
            canvas.restoreState()
    
    frame = Frame(scaled_margin, scaled_margin, 
                  page_width - 2*scaled_margin, page_height - 2*scaled_margin,
                  id='normal')
    
    doc = BaseDocTemplate(output_file, pagesize=page_size)
    colored_template = ColoredPageTemplate('colored', [frame], bg_color)
    doc.addPageTemplates([colored_template])
    
    story = []
    
    # Quote style with proper leading calculation (same as dictionary mode)
    quote_style = ParagraphStyle(
        'Quote',
        fontSize=scaled_quote_size,
        alignment=text_align_const,
        fontName=quote_font_name,
        leading=scaled_quote_size + int(12 * scale_factor),  # Same leading calculation as dictionary
        textColor=Color(*quote_rgb)
    )
    
    # Read CSV (one quote per line)
    available_width = page_width - 2 * scaled_margin
    
    with open(csv_file, 'r', encoding='utf-8-sig') as file:
        first_page = True
        
        for line_num, line in enumerate(file, 1):
            quote_text = line.strip()
            if not quote_text:  # Skip empty lines
                continue
                
            if not first_page:
                story.append(PageBreak())
            first_page = False
            
            # Calculate content height for this quote
            content_height = calculate_quote_height(quote_text, quote_style, available_width)
            
            # Calculate positioning spacer
            spacer_amount = get_quote_spacer_amount(page_position, page_height, 
                                                  scaled_margin, content_height)
            
            # Add positioning spacer
            story.append(Spacer(1, spacer_amount))
            
            # Add quote without quotation marks
            quote_paragraph = Paragraph(quote_text, quote_style)
            story.append(quote_paragraph)
    
    doc.build(story)
    return True

# Authored Quote Generation Function
def create_authored_quotes_pdf_from_csv(csv_file, output_file, 
                                       quote_font=None, quote_size=48, 
                                       author_font=None, author_size=24,
                                       page_color="#FFFFFF", quote_color="#000000", author_color="#000000",
                                       page_width_inches=11, page_height_inches=14, text_alignment="left", page_position="bottom"):
    """Create PDF with authored quotes from CSV file (quote,author format)"""
    
    def register_font_safe(font_path, base_name):
        """Safely register a font with unique name to avoid conflicts"""
        if font_path and font_path.endswith('.ttf') and os.path.exists(font_path):
            try:
                unique_name = f"{base_name}_{int(time.time() * 1000)}"
                pdfmetrics.registerFont(TTFont(unique_name, font_path))
                return unique_name
            except Exception as e:
                st.warning(f"Failed to register {font_path}: {e}")
                return 'Times-Bold'
        elif font_path in ['Times-Bold', 'Helvetica-Bold', 'Helvetica', 'Times-Roman', 'Courier', 'Courier-Bold']:
            return font_path
        else:
            return 'Times-Bold'
    
    # Register fonts with Unicode support (same as dictionary mode)
    quote_font_name = register_font_safe(quote_font, 'QuoteFont')
    author_font_name = register_font_safe(author_font, 'AuthorFont')
    
    # Enhanced Unicode support for built-in fonts (same as dictionary mode)
    unicode_font = 'Helvetica'
    unicode_font_bold = 'Helvetica-Bold'
    
    for font_name in [quote_font_name, author_font_name]:
        if font_name in ['Helvetica-Bold', 'Helvetica']:
            try:
                font_configs = [
                    ('C:/Windows/Fonts/segoeui.ttf', 'C:/Windows/Fonts/segoeuib.ttf'),
                    ('/mnt/c/Windows/Fonts/segoeui.ttf', '/mnt/c/Windows/Fonts/segoeuib.ttf'),
                ]
                
                for regular_path, bold_path in font_configs:
                    if os.path.exists(regular_path):
                        try:
                            unicode_name = f"Unicode_{int(time.time() * 1000)}"
                            unicode_bold_name = f"UnicodeBold_{int(time.time() * 1000)}"
                            
                            pdfmetrics.registerFont(TTFont(unicode_name, regular_path))
                            unicode_font = unicode_name
                            
                            if os.path.exists(bold_path):
                                pdfmetrics.registerFont(TTFont(unicode_bold_name, bold_path))
                                unicode_font_bold = unicode_bold_name
                            
                            break
                        except Exception:
                            continue
            except Exception:
                pass
    
    # Update built-in font references to Unicode versions
    if quote_font_name == 'Helvetica-Bold':
        quote_font_name = unicode_font_bold
    elif quote_font_name == 'Helvetica':
        quote_font_name = unicode_font
        
    if author_font_name == 'Helvetica-Bold':
        author_font_name = unicode_font_bold
    elif author_font_name == 'Helvetica':
        author_font_name = unicode_font
    
    # Page setup
    page_width = page_width_inches * inch
    page_height = page_height_inches * inch
    page_size = (page_width, page_height)
    
    # Calculate scaling factors
    baseline_width = 11.0
    baseline_height = 14.0
    width_scale = page_width_inches / baseline_width
    height_scale = page_height_inches / baseline_height
    scale_factor = (width_scale + height_scale) / 2
    
    # Scale font sizes and margins
    scaled_quote_size = int(quote_size * scale_factor)
    scaled_author_size = int(author_size * scale_factor)
    scaled_margin = 1.5 * scale_factor * inch
    scaled_author_spacing = int(24 * scale_factor)  # Space between quote and author
    
    # Get alignment constant
    def get_alignment_constant(alignment):
        from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
        alignments = {
            "left": TA_LEFT,
            "center": TA_CENTER, 
            "right": TA_RIGHT
        }
        return alignments.get(alignment, TA_LEFT)  # Default left for authored quotes
    
    text_align_const = get_alignment_constant(text_alignment)
    
    # Robust content height calculation (same logic as dictionary mode)
    def calculate_authored_quote_height(quote_text, author_text, quote_style, author_style, available_width):
        """Calculate the total height needed for quote + author content with proper safety margins"""
        try:
            # Create temporary paragraphs to measure their height
            quote_para = Paragraph(quote_text, quote_style)
            author_para = Paragraph(author_text, author_style)
            
            # Measure actual heights using ReportLab's wrap method
            max_height = 20 * inch  # Large enough to not constrain wrapping
            
            quote_height = quote_para.wrap(available_width, max_height)[1]
            author_height = author_para.wrap(available_width, max_height)[1]
            
            # Add spacing between quote and author
            total_height = quote_height + scaled_author_spacing + author_height
            
            # Font-specific adjustments - different fonts need different spacing
            font_padding = 0
            
            # Check for custom TTF fonts (they often have unpredictable metrics)
            if hasattr(quote_style, 'fontName') and ('_' in str(quote_style.fontName) or any(char.isdigit() for char in str(quote_style.fontName)[-10:])):
                font_padding += scaled_quote_size * 0.20  # 20% extra for custom fonts
                
            if hasattr(author_style, 'fontName') and ('_' in str(author_style.fontName) or any(char.isdigit() for char in str(author_style.fontName)[-10:])):
                font_padding += scaled_author_size * 0.20
            
            total_height += font_padding
            
            # Add extra spacing that ReportLab applies (leading adjustments, etc.)
            extra_leading = int(12 * scale_factor)  # From quote_style leading
            author_leading = int(4 * scale_factor)  # From author_style leading  
            
            total_height += extra_leading + author_leading
            
            # ReportLab can add unexpected spacing for long text blocks
            # Add extra padding for multi-line content (the sneaky overflow culprit)
            quote_line_count = max(1, len(quote_text) // 80)  # Rough estimate of lines
            if quote_line_count > 2:  # Multi-line quotes need extra space
                multiline_padding = quote_line_count * int(8 * scale_factor)
                total_height += multiline_padding
            
            # Character-specific adjustments for special characters that can affect height
            text_content = f"{quote_text} {author_text}"
            if any(ord(char) > 127 for char in text_content):  # Non-ASCII characters
                unicode_padding = int(6 * scale_factor)
                total_height += unicode_padding
            
            # Increased safety margin from 5% to 10% to catch edge cases
            safety_margin = total_height * 0.10
                
            return total_height + safety_margin
            
        except Exception as e:
            # Fallback to conservative estimate if measurement fails
            import streamlit as st
            st.warning(f"Authored quote height calculation failed, using conservative estimate: {e}")
            return 8 * scale_factor * inch  # Conservative fallback
    
    # Exact same positioning approach as dictionary mode (your way: content ENDS at consistent margins)
    def get_authored_quote_spacer_amount(position, page_h, scaled_margin, content_height):
        """Your way: content ENDS at consistent margins, not starts"""
        content_area_height = page_h - 2 * scaled_margin
        
        if position == "top":
            return 0  # Start at top margin, flow down (same as before)
        elif position == "middle":
            # Center the content block in the page
            spacer = (content_area_height - content_height) / 2
            return max(0, spacer)  # Prevent negative spacer if content is too large
        elif position == "bottom":
            # END at bottom margin - calculate where to START so it ends there
            # Just like top starts at top margin, bottom ENDS at bottom margin
            spacer = content_area_height - content_height
            # Add extra conservative margin for bottom positioning (most problematic)
            conservative_bottom_margin = int(20 * scale_factor)
            spacer = spacer - conservative_bottom_margin
            return max(0, spacer)  # Prevent negative spacer if content is too large
        else:
            return 0  # fallback to top
    
    # Convert colors
    def hex_to_rgb(hex_color):
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i+2], 16)/255.0 for i in (0, 2, 4))
    
    bg_color = hex_to_rgb(page_color)
    quote_rgb = hex_to_rgb(quote_color)
    author_rgb = hex_to_rgb(author_color)
    
    # Create document with colored background
    from reportlab.platypus import PageTemplate, Frame, BaseDocTemplate
    from reportlab.lib.colors import Color
    
    class ColoredPageTemplate(PageTemplate):
        def __init__(self, id, frames, bg_color, **kwargs):
            super().__init__(id, frames, **kwargs)
            self.bg_color = bg_color
        
        def beforeDrawPage(self, canvas, doc):
            canvas.saveState()
            canvas.setFillColor(Color(*self.bg_color))
            canvas.rect(0, 0, page_width, page_height, fill=1, stroke=0)
            canvas.restoreState()
    
    frame = Frame(scaled_margin, scaled_margin, 
                  page_width - 2*scaled_margin, page_height - 2*scaled_margin,
                  id='normal')
    
    doc = BaseDocTemplate(output_file, pagesize=page_size)
    colored_template = ColoredPageTemplate('colored', [frame], bg_color)
    doc.addPageTemplates([colored_template])
    
    story = []
    
    # Quote and author styles with proper leading calculation (same as dictionary mode)
    quote_style = ParagraphStyle(
        'Quote',
        fontSize=scaled_quote_size,
        alignment=text_align_const,
        fontName=quote_font_name,
        leading=scaled_quote_size + int(12 * scale_factor),  # Same leading calculation as dictionary
        textColor=Color(*quote_rgb)
    )
    
    author_style = ParagraphStyle(
        'Author',
        fontSize=scaled_author_size,
        spaceAfter=0,
        alignment=text_align_const,
        fontName=author_font_name,
        leading=scaled_author_size + int(4 * scale_factor),  # Same leading calculation as dictionary
        textColor=Color(*author_rgb)
    )
    
    # Read CSV (quote,author format)
    available_width = page_width - 2 * scaled_margin
    
    with open(csv_file, 'r', encoding='utf-8-sig') as file:
        reader = csv.reader(file)
        first_page = True
        
        for line_num, row in enumerate(reader, 1):
            if len(row) < 2:  # Skip malformed rows
                continue
                
            quote_text = row[0].strip()
            author_text = row[1].strip()
            
            if not quote_text:  # Skip empty quotes
                continue
                
            if not first_page:
                story.append(PageBreak())
            first_page = False
            
            # Calculate content height for this quote + author
            content_height = calculate_authored_quote_height(quote_text, author_text, quote_style, author_style, available_width)
            
            # Calculate positioning spacer
            spacer_amount = get_authored_quote_spacer_amount(page_position, page_height, 
                                                           scaled_margin, content_height)
            
            # Add positioning spacer
            story.append(Spacer(1, spacer_amount))
            
            # Add quote without quotation marks
            quote_paragraph = Paragraph(quote_text, quote_style)
            story.append(quote_paragraph)
            
            # Add spacing between quote and author
            story.append(Spacer(1, scaled_author_spacing))
            
            # Add author attribution
            author_paragraph = Paragraph(author_text, author_style)
            story.append(author_paragraph)
    
    doc.build(story)
    return True

def convert_pdf_to_png(pdf_path, output_folder, dpi=150, progress_callback=None):
    """Convert PDF to PNG images with high quality using PyMuPDF"""
    if not PYMUPDF_AVAILABLE:
        st.error("PyMuPDF library not available. Please install: pip install PyMuPDF")
        return []
    
    try:
        # Open PDF document
        pdf_document = fitz.open(pdf_path)
        png_files = []
        total_pages = pdf_document.page_count
        
        # Ensure output folder exists
        os.makedirs(output_folder, exist_ok=True)
        
        # Convert each page to PNG
        for page_num in range(total_pages):
            page = pdf_document[page_num]
            
            # Create transformation matrix for DPI scaling
            # 72 DPI is default, so scale factor = target_dpi / 72
            scale_factor = dpi / 72.0
            matrix = fitz.Matrix(scale_factor, scale_factor)
            
            # Render page to pixmap (image)
            pixmap = page.get_pixmap(matrix=matrix)
            
            # Save as PNG
            png_filename = f"page_{page_num + 1:03d}.png"
            png_path = os.path.join(output_folder, png_filename)
            pixmap.save(png_path)
            png_files.append(png_path)
            
            # Update progress if callback provided
            if progress_callback:
                progress = (page_num + 1) / total_pages
                progress_callback(progress, f"Converting page {page_num + 1}/{total_pages}")
            
            # Clean up pixmap memory
            pixmap = None
        
        # Close PDF document
        pdf_document.close()
        
        return png_files
    except Exception as e:
        st.error(f"Failed to convert PDF to PNG: {str(e)}")
        return []

def generate_pngs_from_csv(csv_file, term_font, term_size, term_spacing,
                          pronunciation_font, pronunciation_size, pronunciation_spacing,
                          definition_font, definition_size, page_color,
                          term_color, pronunciation_color, line_color, definition_color,
                          text_alignment, page_position, generate_all_sizes, 
                          page_width_inches=11, page_height_inches=14, document_type="Dictionary"):
    """Generate PNGs by creating PDFs and converting them"""
    
    if not PYMUPDF_AVAILABLE:
        st.error("PNG generation requires PyMuPDF library. Please install: pip install PyMuPDF")
        return None
    
    png_folders = []
    temp_files = []
    
    try:
        if generate_all_sizes:
            # Generate PNGs for all standard sizes
            standard_sizes = [
                ("11x14", 11, 14),
                ("16x20", 16, 20),
                ("18x24", 18, 24),
                ("24x36", 24, 36),
                ("A0", 33.1, 46.8)
            ]
            
            for size_name, width, height in standard_sizes:
                # Create temporary PDF
                with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp_pdf:
                    temp_pdf_path = tmp_pdf.name
                    temp_files.append(temp_pdf_path)
                
                # Generate PDF based on document type
                if document_type == "Dictionary":
                    success = create_pdf_from_csv(
                        csv_file, temp_pdf_path,
                        term_font, term_size, term_spacing,
                        pronunciation_font, pronunciation_size, pronunciation_spacing,
                        definition_font, definition_size, page_color,
                        term_color, pronunciation_color, line_color, definition_color,
                        width, height, text_alignment, page_position
                    )
                elif document_type == "Authored Quotes":
                    success = create_authored_quotes_pdf_from_csv(
                        csv_file, temp_pdf_path,
                        term_font, term_size,
                        pronunciation_font, pronunciation_size,
                        page_color, term_color, pronunciation_color,
                        width, height, text_alignment, page_position
                    )
                else:  # Regular Quotes mode
                    success = create_quotes_pdf_from_csv(
                        csv_file, temp_pdf_path,
                        term_font, term_size,
                        page_color, term_color,
                        width, height, text_alignment, page_position
                    )
                
                if success:
                    # Create folder for this size
                    folder_name = f"{document_type.lower().replace(' ', '_')}_pngs_{size_name}"
                    
                    def progress_update(progress, message):
                        # Update progress for this specific size conversion
                        overall_progress = (len(png_folders) + progress) / len(standard_sizes)
                        st.session_state.conversion_progress = overall_progress
                        st.session_state.conversion_message = f"{size_name}: {message}"
                    
                    png_files = convert_pdf_to_png(temp_pdf_path, folder_name, dpi=150, progress_callback=progress_update)
                    
                    if png_files:
                        png_folders.append((folder_name, size_name, len(png_files)))
        
        else:
            # Generate single size PNG
            with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp_pdf:
                temp_pdf_path = tmp_pdf.name
                temp_files.append(temp_pdf_path)
            
            # Generate PDF based on document type
            if document_type == "Dictionary":
                success = create_pdf_from_csv(
                    csv_file, temp_pdf_path,
                    term_font, term_size, term_spacing,
                    pronunciation_font, pronunciation_size, pronunciation_spacing,
                    definition_font, definition_size, page_color,
                    term_color, pronunciation_color, line_color, definition_color,
                    page_width_inches, page_height_inches, text_alignment, page_position
                )
            elif document_type == "Authored Quotes":
                success = create_authored_quotes_pdf_from_csv(
                    csv_file, temp_pdf_path,
                    term_font, term_size,
                    pronunciation_font, pronunciation_size,
                    page_color, term_color, pronunciation_color,
                    page_width_inches, page_height_inches, text_alignment, page_position
                )
            else:  # Regular Quotes mode
                success = create_quotes_pdf_from_csv(
                    csv_file, temp_pdf_path,
                    term_font, term_size,
                    page_color, term_color,
                    page_width_inches, page_height_inches, text_alignment, page_position
                )
            
            if success:
                # Create folder for single size
                folder_name = f"{document_type.lower().replace(' ', '_')}_pngs_{page_width_inches}x{page_height_inches}"
                
                def progress_update(progress, message):
                    st.session_state.conversion_progress = progress
                    st.session_state.conversion_message = message
                
                png_files = convert_pdf_to_png(temp_pdf_path, folder_name, dpi=150, progress_callback=progress_update)
                
                if png_files:
                    png_folders.append((folder_name, f"{page_width_inches}x{page_height_inches}", len(png_files)))
        
        # Clean up temporary PDF files
        for temp_file in temp_files:
            if os.path.exists(temp_file):
                os.unlink(temp_file)
        
        return png_folders
        
    except Exception as e:
        # Clean up on error
        for temp_file in temp_files:
            if os.path.exists(temp_file):
                os.unlink(temp_file)
        raise e

# Initialize session state
if 'font_cache_initialized' not in st.session_state:
    st.session_state.font_cache_initialized = False
    st.session_state.available_fonts_list = None

# Font discovery and management
@st.cache_data
def get_available_fonts():
    """Scan system for available TTF fonts"""
    fonts = {
        "Times-Bold": "Times Bold (Default)",
        "Helvetica-Bold": "Helvetica Bold", 
        "Helvetica": "Helvetica Regular",
        "Times-Roman": "Times Roman",
        "Courier": "Courier",
        "Courier-Bold": "Courier Bold"
    }
    
    # System font directories to scan
    font_dirs = [
        "C:/Windows/Fonts/",
        "/mnt/c/Windows/Fonts/",
        "/System/Library/Fonts/",  # macOS
        "/usr/share/fonts/truetype/",  # Linux
    ]
    
    for font_dir in font_dirs:
        if os.path.exists(font_dir):
            try:
                for root, dirs, files in os.walk(font_dir):
                    for file in files:
                        if file.lower().endswith('.ttf'):
                            font_path = os.path.join(root, file)
                            
                            # Create readable display name
                            display_name = file.replace('.ttf', '').replace('.TTF', '')
                            display_name = display_name.replace('-', ' ').replace('_', ' ')
                            display_name = display_name.replace('Regular', '').replace('regular', '')
                            display_name = display_name.replace('Bold', 'Bold').replace('bold', 'Bold')
                            display_name = display_name.replace('Italic', 'Italic').replace('italic', 'Italic')
                            display_name = ' '.join(display_name.split())
                            
                            if not display_name:
                                display_name = file.replace('.ttf', '')
                            
                            fonts[font_path] = f"{display_name} ({file})"
            except Exception:
                continue
    
    return fonts

# Load fonts with caching
if not st.session_state.font_cache_initialized:
    with st.spinner("Scanning system fonts..."):
        st.session_state.available_fonts_list = get_available_fonts()
        st.session_state.font_cache_initialized = True

available_fonts = st.session_state.available_fonts_list

# Sidebar for settings
st.sidebar.header("Settings")

# Document Type
st.sidebar.subheader("Document Type")
document_type = st.sidebar.selectbox(
    "Select Document Type",
    options=["Dictionary", "Quotes", "Authored Quotes"],
    index=0,
    key="document_type",
    help="Choose the type of document to generate"
)

# Font management buttons
col_refresh1, col_refresh2 = st.sidebar.columns(2)
with col_refresh1:
    if st.button("Refresh Fonts", help="Rescan system fonts and reset colors"):
        st.session_state.font_cache_initialized = False
        # Reset all color selections to defaults
        for key in ['page_color', 'term_color', 'pronunciation_color', 'line_color', 'definition_color']:
            if key in st.session_state:
                del st.session_state[key]
        st.rerun()

with col_refresh2:
    if st.button("Clear Cache", help="Clear font cache and reset colors"):
        try:
            from reportlab.pdfbase.pdfmetrics import _fonts
            for font_name in list(_fonts.keys()):
                if 'Custom' in font_name or 'Unicode' in font_name:
                    del _fonts[font_name]
            # Reset all color selections to defaults
            for key in ['page_color', 'term_color', 'pronunciation_color', 'line_color', 'definition_color']:
                if key in st.session_state:
                    del st.session_state[key]
            st.sidebar.success("Cache cleared!")
        except Exception:
            st.sidebar.warning("Cache clear failed")

# Font selection
st.sidebar.subheader("Font Selection")
st.sidebar.info(f"Found {len(available_fonts)} fonts")

term_font = st.sidebar.selectbox(
    "Terms Font",
    options=list(available_fonts.keys()),
    format_func=lambda x: available_fonts.get(x, x),
    index=0,
    key="term_font_select",
    help="Font for the main term/word"
)

# Document-specific font selections
if document_type == "Dictionary":
    pronunciation_font = st.sidebar.selectbox(
        "Pronunciation Font", 
        options=list(available_fonts.keys()),
        format_func=lambda x: available_fonts.get(x, x),
        index=1 if len(available_fonts) > 1 else 0,
        key="pronunciation_font_select",
        help="Font for pronunciation and word type"
    )

    definition_font = st.sidebar.selectbox(
        "Definition Font",
        options=list(available_fonts.keys()),
        format_func=lambda x: available_fonts.get(x, x),
        index=2 if len(available_fonts) > 2 else 0,
        key="definition_font_select",
        help="Font for the definition text"
    )
elif document_type == "Authored Quotes":
    # Author font selection for authored quotes
    pronunciation_font = st.sidebar.selectbox(
        "Author Font", 
        options=list(available_fonts.keys()),
        format_func=lambda x: available_fonts.get(x, x),
        index=1 if len(available_fonts) > 1 else 0,
        key="pronunciation_font_select",
        help="Font for author attribution"
    )
    definition_font = term_font  # Not used in authored quotes
else:
    # For regular quotes mode, use term font for everything
    pronunciation_font = term_font
    definition_font = term_font


# Font sizes
st.sidebar.subheader("Font Sizes")
if document_type == "Dictionary":
    term_size = st.sidebar.slider("Terms Size", 60, 120, 84, key="term_size")
    pronunciation_size = st.sidebar.slider("Pronunciation Size", 20, 40, 28, key="pronunciation_size")
    definition_size = st.sidebar.slider("Definition Size", 20, 40, 28, key="definition_size")
elif document_type == "Authored Quotes":
    term_size = st.sidebar.slider("Quote Text Size", 20, 80, 48, key="term_size", help="Font size for quote text")
    pronunciation_size = st.sidebar.slider("Author Size", 16, 36, 24, key="pronunciation_size", help="Font size for author attribution")
    definition_size = term_size  # Not used in authored quotes
else:  # Regular Quotes mode
    term_size = st.sidebar.slider("Quote Text Size", 20, 80, 48, key="term_size", help="Font size for quote text")
    # Use same size for all quote elements
    pronunciation_size = term_size
    definition_size = term_size

# Page Size
st.sidebar.subheader("Page Size")
page_size_options = {
    "11 × 14 inch (3300 × 4200 px)": (11, 14),
    "16 × 20 inch (4800 × 6000 px)": (16, 20),
    "18 × 24 inch (5400 × 7200 px)": (18, 24),
    "24 × 36 inch (7200 × 10800 px)": (24, 36),
    "33.1 × 46.8 inch - A0 (9930 × 14040 px)": (33.1, 46.8),
    "Custom": (None, None)
}

selected_size = st.sidebar.selectbox(
    "Select Page Size",
    options=list(page_size_options.keys()),
    index=0,
    key="page_size_select"
)

if selected_size == "Custom":
    col_w, col_h = st.sidebar.columns(2)
    with col_w:
        custom_width = st.sidebar.number_input("Width (inches)", min_value=1.0, max_value=50.0, value=11.0, step=0.1, key="custom_width")
    with col_h:
        custom_height = st.sidebar.number_input("Height (inches)", min_value=1.0, max_value=50.0, value=14.0, step=0.1, key="custom_height")
    page_width_inches, page_height_inches = custom_width, custom_height
else:
    page_width_inches, page_height_inches = page_size_options[selected_size]

# Text Alignment & Positioning
st.sidebar.subheader("Text Alignment & Position")

# Text alignment (within the text group)
text_alignment = st.sidebar.selectbox(
    "Text Alignment",
    options=["left", "center", "right"],
    index=0,
    key="text_alignment",
    help="How text is aligned within the text group"
)

# Page positioning (vertical only)
page_position = st.sidebar.selectbox(
    "Vertical Position",
    options=["bottom", "middle", "top"],
    index=0,  # bottom as current default
    key="page_position",
    help="Vertical position of content on the page"
)

# Generation options
generate_all_sizes = st.sidebar.checkbox("Generate All Standard Sizes", 
                                         help="Generate PDFs in all standard sizes at once",
                                         key="generate_all_sizes")

# Spacing
if document_type == "Dictionary":
    st.sidebar.subheader("Spacing")
    term_spacing = st.sidebar.slider("After Terms", 20, 80, 48, key="term_spacing")
    pronunciation_spacing = st.sidebar.slider("After Pronunciation", 20, 60, 36, key="pronunciation_spacing")
else:
    # No spacing controls needed for quotes
    term_spacing = 48
    pronunciation_spacing = 36

# Colors
st.sidebar.subheader("Colors")
page_color = st.sidebar.color_picker("Background Color", "#FFFFFF", key="page_color")

if document_type == "Dictionary":
    term_color = st.sidebar.color_picker("Terms Color", "#000000", key="term_color")
    pronunciation_color = st.sidebar.color_picker("Pronunciation Color", "#000000", key="pronunciation_color")
    line_color = st.sidebar.color_picker("Dividing Line Color", "#000000", key="line_color")
    definition_color = st.sidebar.color_picker("Definition Color", "#000000", key="definition_color")
elif document_type == "Authored Quotes":
    term_color = st.sidebar.color_picker("Quote Text Color", "#000000", key="term_color")
    pronunciation_color = st.sidebar.color_picker("Author Color", "#000000", key="pronunciation_color")
    # Not used in authored quotes
    line_color = term_color
    definition_color = term_color
else:  # Regular Quotes mode
    term_color = st.sidebar.color_picker("Quote Text Color", "#000000", key="term_color")
    # Use same color for all quote elements
    pronunciation_color = term_color
    line_color = term_color
    definition_color = term_color

# Main content area - File Upload
st.header("File Upload")
uploaded_file = st.file_uploader("Choose CSV file", type=['csv'])

if uploaded_file is not None:
    # Validation based on document type
    if document_type == "Dictionary":
        df = pd.read_csv(uploaded_file)
        required_columns = ['term', 'pronunciation', 'type', 'definition']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            st.error(f"Dictionary mode requires columns: {', '.join(missing_columns)}")
            st.info("Your CSV should have: term, pronunciation, type, definition")
            csv_valid = False
        else:
            st.success(f"Dictionary CSV loaded: {len(df)} entries found")
            csv_valid = True
    elif document_type == "Authored Quotes":
        # For authored quotes, validate CSV format (quote,author)
        try:
            content = uploaded_file.read().decode('utf-8')
            lines = [line.strip() for line in content.splitlines() if line.strip()]
            valid_quotes = 0
            
            for line in lines:
                if ',' in line:
                    parts = line.split(',', 1)  # Split on first comma only
                    if len(parts) == 2 and parts[0].strip() and parts[1].strip():
                        valid_quotes += 1
            
            if valid_quotes > 0:
                st.success(f"Authored Quotes CSV loaded: {valid_quotes} quotes found")
                csv_valid = True
            else:
                st.error("No valid authored quotes found. Format should be: quote,author")
                st.info("Example: \"The world is beautiful,-Henry Wadsworth Longfellow\"")
                csv_valid = False
            
            # Reset file pointer for later use
            uploaded_file.seek(0)
        except Exception as e:
            st.error(f"Error reading authored quotes file: {e}")
            csv_valid = False
    else:  # Regular Quotes mode
        # For quotes, read line by line (no column validation needed)
        try:
            content = uploaded_file.read().decode('utf-8')
            quotes = [line.strip() for line in content.splitlines() if line.strip()]
            st.success(f"Quotes CSV loaded: {len(quotes)} quotes found")
            csv_valid = True
            # Reset file pointer for later use
            uploaded_file.seek(0)
        except Exception as e:
            st.error(f"Error reading quotes file: {e}")
            csv_valid = False
    
    if csv_valid:
        
        # Generate buttons side by side
        col1, col2 = st.columns(2)
        
        with col1:
            generate_pdf = st.button("Generate PDF", type="primary", use_container_width=True)
        with col2:
            generate_png = st.button("Generate PNGs", type="secondary", use_container_width=True)
        
        if generate_pdf:
            status_placeholder = st.empty()
            progress_bar = st.progress(0)
            
            try:
                # Save uploaded file temporarily
                with tempfile.NamedTemporaryFile(mode='wb', suffix='.csv', delete=False) as tmp_file:
                    tmp_file.write(uploaded_file.getbuffer())
                    temp_csv_path = tmp_file.name
                
                progress_bar.progress(25)
                status_placeholder.info("Processing CSV data...")
                
                if generate_all_sizes:
                    # Generate PDFs for all standard sizes
                    standard_sizes = [
                        ("11x14", 11, 14),
                        ("16x20", 16, 20),
                        ("18x24", 18, 24),
                        ("24x36", 24, 36),
                        ("A0", 33.1, 46.8)
                    ]
                    
                    zip_files = []
                    total_sizes = len(standard_sizes)
                    
                    for i, (size_name, width, height) in enumerate(standard_sizes):
                        progress = 25 + (i * 60 // total_sizes)
                        progress_bar.progress(progress)
                        status_placeholder.info(f"Generating {size_name} ({width}×{height} inch)...")
                        
                        if document_type == "Dictionary":
                            output_path = f"dictionary_{size_name}.pdf"
                            success = create_pdf_from_csv(
                                temp_csv_path, 
                                output_path,
                                term_font, term_size, term_spacing,
                                pronunciation_font, pronunciation_size, pronunciation_spacing,
                                definition_font, definition_size, page_color,
                                term_color, pronunciation_color, line_color, definition_color,
                                width, height, text_alignment, page_position
                            )
                        elif document_type == "Authored Quotes":
                            output_path = f"authored_quotes_{size_name}.pdf"
                            success = create_authored_quotes_pdf_from_csv(
                                temp_csv_path,
                                output_path,
                                term_font, term_size,  # Quote font/size
                                pronunciation_font, pronunciation_size,  # Author font/size
                                page_color, term_color, pronunciation_color,  # Quote & author colors
                                width, height, text_alignment, page_position
                            )
                        else:  # Regular Quotes mode
                            output_path = f"quotes_{size_name}.pdf"
                            success = create_quotes_pdf_from_csv(
                                temp_csv_path,
                                output_path,
                                term_font, term_size,  # Use term font/size for quotes
                                page_color, term_color,  # Use term color for quote text
                                width, height, text_alignment, page_position
                            )
                        
                        if success and os.path.exists(output_path):
                            filename = f"{document_type.lower().replace(' ', '_')}_{size_name}.pdf"
                            zip_files.append((output_path, filename))
                    
                    if zip_files:
                        # Create ZIP file with all PDFs
                        import zipfile
                        zip_path = "dictionary_all_sizes.zip"
                        with zipfile.ZipFile(zip_path, 'w') as zipf:
                            for file_path, archive_name in zip_files:
                                zipf.write(file_path, archive_name)
                                os.unlink(file_path)  # Clean up individual files
                        
                        progress_bar.progress(100)
                        status_placeholder.success(f"Generated {len(zip_files)} PDFs in all sizes!")
                        
                        # Provide ZIP download
                        with open(zip_path, "rb") as zip_file:
                            zip_bytes = zip_file.read()
                            st.download_button(
                                label="Download All Sizes (ZIP)",
                                data=zip_bytes,
                                file_name="dictionary_all_sizes.zip",
                                mime="application/zip"
                            )
                        os.unlink(zip_path)  # Clean up ZIP file
                    else:
                        status_placeholder.error("Failed to generate PDFs")
                
                else:
                    # Generate single PDF with selected size
                    progress_bar.progress(50)
                    status_placeholder.info(f"Applying fonts and generating {document_type} PDF...")
                    
                    if document_type == "Dictionary":
                        output_path = "generated_dictionary.pdf"
                        success = create_pdf_from_csv(
                            temp_csv_path, 
                            output_path,
                            term_font, term_size, term_spacing,
                            pronunciation_font, pronunciation_size, pronunciation_spacing,
                            definition_font, definition_size, page_color,
                            term_color, pronunciation_color, line_color, definition_color,
                            page_width_inches, page_height_inches, text_alignment, page_position
                        )
                    elif document_type == "Authored Quotes":
                        output_path = "generated_authored_quotes.pdf"
                        success = create_authored_quotes_pdf_from_csv(
                            temp_csv_path,
                            output_path,
                            term_font, term_size,  # Quote font/size
                            pronunciation_font, pronunciation_size,  # Author font/size
                            page_color, term_color, pronunciation_color,  # Quote & author colors
                            page_width_inches, page_height_inches, text_alignment, page_position
                        )
                    else:  # Regular Quotes mode
                        output_path = "generated_quotes.pdf"
                        success = create_quotes_pdf_from_csv(
                            temp_csv_path,
                            output_path,
                            term_font, term_size,  # Use term font/size for quotes
                            page_color, term_color,  # Use term color for quote text
                            page_width_inches, page_height_inches, text_alignment, page_position
                        )
                    
                    if success:
                        progress_bar.progress(100)
                        size_text = f"{page_width_inches}×{page_height_inches} inch"
                        status_placeholder.success(f"{document_type} PDF generated successfully! ({size_text})")
                        
                        # Provide download
                        with open(output_path, "rb") as pdf_file:
                            pdf_bytes = pdf_file.read()
                            download_filename = f"{document_type.lower().replace(' ', '_')}.pdf"
                            st.download_button(
                                label=f"Download {document_type} PDF",
                                data=pdf_bytes,
                                file_name=download_filename,
                                mime="application/pdf"
                            )
                
            except Exception as e:
                progress_bar.progress(0)
                status_placeholder.error(f"Error: {str(e)}")
            
            finally:
                # Clean up
                if 'temp_csv_path' in locals() and os.path.exists(temp_csv_path):
                    os.unlink(temp_csv_path)
        
        elif generate_png:
            if not PYMUPDF_AVAILABLE:
                st.error("PNG generation requires PyMuPDF library. Please install it:")
                st.code("pip install PyMuPDF", language="bash")
                st.info("PyMuPDF has no system dependencies - just install and it works!")
            else:
                status_placeholder = st.empty()
                progress_bar = st.progress(0)
                
                try:
                    # Save uploaded file temporarily
                    with tempfile.NamedTemporaryFile(mode='wb', suffix='.csv', delete=False) as tmp_file:
                        tmp_file.write(uploaded_file.getbuffer())
                        temp_csv_path = tmp_file.name
                    
                    progress_bar.progress(25)
                    status_placeholder.info("Processing CSV data...")
                    
                    # Generate PNGs
                    progress_bar.progress(50)
                    status_placeholder.info("Generating PNGs...")
                    
                    # Pass document type to PNG generation
                    png_folders = generate_pngs_from_csv(
                        temp_csv_path,
                        term_font, term_size, term_spacing,
                        pronunciation_font, pronunciation_size, pronunciation_spacing,
                        definition_font, definition_size, page_color,
                        term_color, pronunciation_color, line_color, definition_color,
                        text_alignment, page_position, generate_all_sizes,
                        page_width_inches, page_height_inches, document_type
                    )
                    
                    if png_folders:
                        progress_bar.progress(75)
                        status_placeholder.info("Creating ZIP archive...")
                        
                        # Create ZIP file with all PNG folders
                        zip_path = f"{document_type.lower().replace(' ', '_')}_pngs.zip"
                        with zipfile.ZipFile(zip_path, 'w') as zipf:
                            for folder_name, size_name, file_count in png_folders:
                                # Add all PNG files from each folder to ZIP
                                for root, dirs, files in os.walk(folder_name):
                                    for file in files:
                                        if file.endswith('.png'):
                                            file_path = os.path.join(root, file)
                                            # Preserve folder structure in ZIP
                                            arc_name = os.path.relpath(file_path, '.')
                                            zipf.write(file_path, arc_name)
                                
                                # Clean up folder after adding to ZIP
                                shutil.rmtree(folder_name, ignore_errors=True)
                        
                        progress_bar.progress(100)
                        total_folders = len(png_folders)
                        total_files = sum(count for _, _, count in png_folders)
                        status_placeholder.success(f"Generated {total_files} PNG files in {total_folders} size(s)!")
                        
                        # Provide ZIP download
                        with open(zip_path, "rb") as zip_file:
                            zip_bytes = zip_file.read()
                            st.download_button(
                                label=f"Download {document_type} PNGs (ZIP)",
                                data=zip_bytes,
                                file_name=f"{document_type.lower().replace(' ', '_')}_pngs.zip",
                                mime="application/zip"
                            )
                        
                        # Show folder structure
                        st.info("**Folder structure:**")
                        for folder_name, size_name, file_count in png_folders:
                            st.write(f"└── `{folder_name}/` ({file_count} PNG files)")
                        
                        os.unlink(zip_path)  # Clean up ZIP file
                    else:
                        status_placeholder.error("Failed to generate PNG files")
                
                except Exception as e:
                    progress_bar.progress(0)
                    status_placeholder.error(f"Error: {str(e)}")
                
                finally:
                    # Clean up
                    if 'temp_csv_path' in locals() and os.path.exists(temp_csv_path):
                        os.unlink(temp_csv_path)


# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #25656E; padding: 20px;'>
    <p><strong>CSVibe</strong> - CSV meets beautiful design</p>
    <p>Contact: <span style='color: white;'>shzbkh1@gmail.com</span> | Discord: <span style='color: white;'>shzbk</span></p>
</div>
""", unsafe_allow_html=True)
