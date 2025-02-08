from pandas.io.sas.sas_constants import page_size_length
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, Image, Frame
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.platypus.doctemplate import PageTemplate, BaseDocTemplate
import pandas as pd
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Register Calibri fonts
pdfmetrics.registerFont(TTFont('Calibri', 'Calibri.ttf'))
pdfmetrics.registerFont(TTFont('Calibri-Bold', 'Calibrib.ttf'))

class MilitaryRosterDocument(BaseDocTemplate):
    def __init__(self, filename, **kwargs):
        super().__init__(filename, **kwargs)
        self.page_width, self.page_height = landscape(letter)

        # Define the main content frame
        # Define the main content frame - start below header
        content_frame = Frame(
            x1=0.5 * inch,  # left margin
            y1=1.0 * inch,  # bottom margin (increased to make room for footer)
            width=self.page_width - inch,  # width with margins
            height=self.page_height - 2.7 * inch,  # reduced height to make room for header
            id='normal'
        )

        template = PageTemplate(
            id='military_roster',
            frames=content_frame,
            onPage=self.add_page_elements
        )
        self.addPageTemplates([template])

    def add_page_elements(self, canvas, doc):
        """Add header and footer to each page"""
        canvas.saveState()
        self.add_header(canvas, doc)
        canvas.setLineWidth(2.5)
        canvas.setStrokeColorRGB(0, 0.32, 0.65)  # Navy Blue color
        canvas.line(
            0 * inch,  # left margin
            self.page_height - 1.7 * inch,  # just below header
            self.page_width - 0 * inch,  # right margin
            self.page_height - 1.7 * inch
        )

        # Add footer top border
        canvas.setLineWidth(2.5)
        canvas.setStrokeColorRGB(0, 0.32, 0.65)  # Navy Blue color
        canvas.line(
            0 * inch,  # left margin
            1.2 * inch,  # just above footer
            self.page_width - 0 * inch,  # right margin
            1.2 * inch
        )
        self.add_footer(canvas, doc)
        canvas.restoreState()

    def add_header(self, canvas, doc):

        header_promotion_codes = {

        }
        # CUI Header at the very top
        canvas.setFont('Helvetica-Bold', 10)
        canvas.drawCentredString(
            self.page_width / 2,
            self.page_height - 0.3 * inch,
            'CUI// CONTROLLED UNCLASSIFIED INFORMATION'
        )

        # Start the main header content below the CUI line
        header_top = self.page_height - 0.8 * inch

        # Logo on the left
        logo_path = doc.logo_path
        logo = Image(logo_path, width=1 * inch, height=1 * inch)
        logo.drawOn(canvas, 0.5 * inch, header_top - 0.8 * inch)

        # Title "Unit Data"
        canvas.setFont('Helvetica-Bold', 12)  # Slightly larger font for title
        text_start_x = 2 * inch
        title_y = header_top + 0.1 * inch  # Position title above other text
        canvas.drawString(text_start_x, title_y, "Unit Data")

        # Add divider line under "Unit Data"
        line_width = 3 * inch  # Adjust width as needed
        line_y = title_y - 0.05 * inch  # Position slightly below the title
        canvas.line(text_start_x, line_y, text_start_x + line_width, line_y)

        # PAS Information - aligned with logo height
        canvas.setFont('Helvetica', 10)  # Back to regular font size
        line_height = 0.2 * inch  # Consistent line spacing
        text_start_y = header_top - 0.1 * inch  # Center text block with logo

        # Draw each line of text
        canvas.drawString(text_start_x, text_start_y, f"SRID: {doc.pas_info['srid']}")
        canvas.drawString(text_start_x, text_start_y - line_height, f"FD NAME: {doc.pas_info['fd name']}")
        canvas.drawString(text_start_x, text_start_y - 2 * line_height, f"FDID: {doc.pas_info['fdid']}")
        canvas.drawString(text_start_x, text_start_y - 3 * line_height, f"SRID MPF: {doc.pas_info['srid mpf']}")

        canvas.setFont('Helvetica-Bold', 12)  # Slightly larger font for title
        text_start_x = 5 * inch
        title_y = header_top + 0.1 * inch  # Position title above other text
        canvas.drawString(text_start_x, title_y, "Promotion Key")

    def add_footer(self, canvas, doc):
        footer_text = (
            "The information herein is FOR OFFICIAL USE ONLY (CUI) information which must be protected under "
            "the Freedom of Information Act (5 U.S.C. 552) and/or the Privacy Act of 1974 (5 U.S.C. 552a). "
            "Unauthorized disclosure or misuse of this PERSONAL INFORMATION may result in disciplinary action, "
            "criminal and/or civil penalties."
        )

        canvas.setFont('Helvetica', 8)
        footer_width = self.page_width - inch  # Total width for footer text
        x = 0.5 * inch  # Starting x position (left margin)
        y = 0.75 * inch  # Bottom margin for main paragraph

        # Split footer text into lines that fit within the page width
        from reportlab.pdfbase.pdfmetrics import stringWidth
        words = footer_text.split()
        lines = []
        current_line = []
        current_width = 0
        max_width = self.page_width - inch  # Available width for text

        for word in words:
            word_width = stringWidth(word + ' ', 'Helvetica', 8)
            if current_width + word_width <= max_width:
                current_line.append(word)
                current_width += word_width
            else:
                lines.append(' '.join(current_line))
                current_line = [word]
                current_width = word_width

        if current_line:
            lines.append(' '.join(current_line))

        # Draw footer lines centered
        for i, line in enumerate(lines):
            line_width = stringWidth(line, 'Helvetica', 8)
            center_x = (self.page_width - line_width) / 2
            canvas.drawString(center_x, y + (len(lines) - 1 - i) * 10, line)

        # Add the bottom footer line with three elements
        bottom_y = 0.3 * inch  # Position for the bottom line

        # Left aligned date (placeholder - you can modify this)
        canvas.setFont('Helvetica-Bold', 10)
        canvas.drawString(x, bottom_y, "15 March 2024")

        # Center aligned "CUI"
        canvas.setFont('Helvetica-Bold', 10)
        cui_text = "CUI"
        cui_width = stringWidth(cui_text, 'Helvetica-Bold', 8)
        canvas.drawString(self.page_width / 2 - cui_width / 2, bottom_y, cui_text)

        # Right aligned page numbers
        canvas.setFont('Helvetica-Bold', 10)
        page_text = f"Page {doc.page} of {doc.page}"  # doc.page will be current page number
        page_width = stringWidth(page_text, 'Helvetica', 8)
        canvas.drawString(self.page_width - 0.5 * inch - page_width, bottom_y, page_text)


def pdf_generator(dataframe, output_filename="military_roster.pdf", logo_path='images/Air_Force_Personnel_Center.png'):
    """Generate a military roster PDF from a DataFrame"""
    # Create PDF document with custom template
    doc = MilitaryRosterDocument(
        output_filename,
        pagesize=landscape(letter),
        rightMargin=0.5 * inch,
        leftMargin=0.5 * inch,
        topMargin=0.5 * inch,
        bottomMargin=0.5 * inch
    )

    # Store additional information needed for header/footer
    doc.logo_path = logo_path
    doc.pas_info = {
        'srid': '0R173',
        'fd name': 'BEEBE III, KENNETH B.',
        'fdid': '0R173FGDF',
        'srid mpf': '0P'
    }

    elements = []

    # Convert DataFrame to table data
    header_row = ['FULL NAME', 'GRADE', 'DAS', 'DAFSC', 'UNIT', 'DOR', 'TAFMSD', 'PASCODE']
    data_rows = dataframe.values.tolist()

    # Create a row that spans the full width with "Eligible"
    eligible_row = [["Eligible"] + [""] * (len(header_row) - 1)]

    # Combine all rows
    table_data = eligible_row + [header_row] + data_rows

    # Create main table
    table = Table(table_data)
    table.setStyle(TableStyle([
        # Eligible row styling
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('SPAN', (0, 0), (-1, 0)),  # Span across all columns
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Calibri'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('ROWHEIGHT', (0, 0), (-1, -1), 30),  # makes all rows 30 points high

        # Header styling
        ('BACKGROUND', (0, 1), (-1, 1), colors.lightgrey),
        ('TEXTCOLOR', (0, 1), (-1, 1), colors.black),
        ('ALIGN', (0, 1), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 1), (-1, 1), 'Calibri'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 1), (-1, 1), 4),

        # Data rows styling
        ('FONTNAME', (0, 2), (-1, -1), 'Calibri'),
        ('LINEBELOW', (0, 0), (-1, -1), .5, colors.lightgrey),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ]))

    elements.append(table)

    # Build PDF
    doc.build(elements)