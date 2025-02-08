from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, Image, Frame
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus.flowables import PageBreak
from reportlab.platypus.doctemplate import PageTemplate, BaseDocTemplate
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Register Calibri fonts
pdfmetrics.registerFont(TTFont('Calibri', 'Calibri.ttf'))
pdfmetrics.registerFont(TTFont('Calibri-Bold', 'Calibrib.ttf'))


class MilitaryRosterDocument(BaseDocTemplate):
    def __init__(self, filename, **kwargs):
        super().__init__(filename, **kwargs)
        self.page_width, self.page_height = landscape(letter)

        # Define the main content frame - start below header
        content_frame = Frame(
            x1=0.5 * inch,  # left margin
            y1=1.05 * inch,  # bottom margin (increased to make room for footer)
            width=self.page_width - inch,  # width with margins
            height=self.page_height - 2.98 * inch,  # reduced height to make room for header
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

        # Add footer border
        canvas.setLineWidth(0.1)
        canvas.setStrokeColorRGB(0, 0, 0)
        canvas.line(
            0.5 * inch,
            1.2 * inch,
            self.page_width - 0.5 * inch,
            1.2 * inch
        )
        self.add_footer(canvas, doc)
        canvas.restoreState()

    def add_header(self, canvas, doc):
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
        canvas.setFont('Calibri-Bold', 12)
        text_start_x = 2 * inch
        title_y = header_top + 0.1 * inch
        canvas.drawString(text_start_x, title_y, "Unit Data")

        # PAS Information
        canvas.setFont('Calibri-Bold', 10)
        line_height = 0.2 * inch
        text_start_y = header_top - 0.1 * inch

        canvas.drawString(text_start_x, text_start_y, f"SRID: {doc.pas_info['srid']}")
        canvas.drawString(text_start_x, text_start_y - line_height, f"FD NAME: {doc.pas_info['fd name']}")
        canvas.drawString(text_start_x, text_start_y - 2 * line_height, f"FDID: {doc.pas_info['fdid']}")
        canvas.drawString(text_start_x, text_start_y - 3 * line_height, f"SRID MPF: {doc.pas_info['srid mpf']}")

        # Promotion Key Title
        canvas.setFont('Calibri-Bold', 12)
        text_start_x = 5 * inch
        canvas.drawString(text_start_x, title_y, "Promotion Eligibility Data")

        # Promotion Key Information
        canvas.setFont('Calibri-Bold', 10)
        line_height = 0.2 * inch
        text_start_y = header_top - 0.1 * inch
        canvas.drawString(text_start_x, text_start_y, f"PROMOTE NOW: {doc.pas_info['promote now']}")
        canvas.drawString(text_start_x, text_start_y - line_height, f"MUST PROMOTE: {doc.pas_info['must promote']}")

        #Signature Block
        canvas.setFont('Calibri-Bold', 12)
        text_start_s = 7.8 * inch
        line_height_s = 0.2 * inch
        title_s = header_top - 0.75 * inch
        canvas.drawString(text_start_s, title_s, "KENNETH B. BEEBE III, Colonel, USAF")
        canvas.drawString(text_start_s, title_s - line_height_s, "Commander, 51st Maintenance Group")

    def add_footer(self, canvas, doc):
        footer_text = (
            "The information herein is FOR OFFICIAL USE ONLY (CUI) information which must be protected under "
            "the Freedom of Information Act (5 U.S.C. 552) and/or the Privacy Act of 1974 (5 U.S.C. 552a). "
            "Unauthorized disclosure or misuse of this PERSONAL INFORMATION may result in disciplinary action, "
            "criminal and/or civil penalties."
        )

        canvas.setFont('Calibri-Bold', 8)
        x = 0.5 * inch
        y = 0.75 * inch

        # Split footer text into lines
        from reportlab.pdfbase.pdfmetrics import stringWidth
        words = footer_text.split()
        lines = []
        current_line = []
        current_width = 0
        max_width = self.page_width - inch

        for word in words:
            word_width = stringWidth(word + ' ', 'Calibri-Bold', 8)
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
            line_width = stringWidth(line, 'Calibri', 8)
            center_x = (self.page_width - line_width) / 2
            canvas.drawString(center_x, y + (len(lines) - 1 - i) * 10, line)

        # Bottom footer elements
        bottom_y = 0.3 * inch
        canvas.setFillColorRGB(0, 0, 0)

        # Left: Date
        canvas.setFont('Calibri-Bold', 12)
        canvas.drawString(x, bottom_y, "15 March 2024")

        # Center: CUI and identifier
        cui_text = "CUI"
        cui_width = stringWidth(cui_text, 'Calibri-Bold', 12)
        cui_center_x = (self.page_width / 2) - (cui_width / 2)
        canvas.drawString(cui_center_x, bottom_y, cui_text)

        identifier_text = "25E7 - Final MEL"
        identifier_width = stringWidth(identifier_text, 'Calibri-Bold', 12)
        identifier_center_x = (self.page_width / 2) - (identifier_width / 2)
        canvas.drawString(identifier_center_x, bottom_y - 18, identifier_text)

        # Right: Page numbers
        page_text = f"Page {canvas.getPageNumber()} of {doc.page}"
        page_width = stringWidth(page_text, 'Calibri-Bold', 12)
        canvas.drawString(self.page_width - 0.5 * inch - page_width, bottom_y, page_text)


def generate_roster_pdf(eligible_df, ineligible_df, cycle, output_filename="military_roster.pdf",
                        logo_path='images/Air_Force_Personnel_Center.png'):
    """Generate a military roster PDF from eligible and ineligible DataFrames"""
    doc = MilitaryRosterDocument(
        output_filename,
        pagesize=landscape(letter),
        rightMargin=0.5 * inch,
        leftMargin=0.5 * inch,
        topMargin=0.5 * inch,
        bottomMargin=0.5 * inch
    )

    # Store additional information
    doc.logo_path = logo_path
    doc.pas_info = {
        'srid': '0R173',
        'fd name': 'BEEBE III, KENNETH B.',
        'fdid': '0R173FGDF',
        'srid mpf': '0P',
        'promote now': '2',
        'must promote': '3'

    }

    elements = []
    header_row = ['FULL NAME', 'GRADE', 'DAS', 'DAFSC', 'UNIT', 'DOR', 'TAFMSD', 'PASCODE']

    # Convert DataFrames to lists
    eligible_data = eligible_df.values.tolist()
    ineligible_data = ineligible_df.values.tolist()

    # Get unique PASCODEs from both eligible and ineligible data
    pascodes = set()
    for row in eligible_data + ineligible_data:
        pascodes.add(row[7])  # PASCODE is the 8th column
    unique_pascodes = sorted(list(pascodes))

    def create_table(data, header, table_type=None, count=None):
        """Create table with optional status row"""
        table_width = doc.page_width - inch
        col_widths = [table_width * x for x in [0.22, 0.07, 0.1, 0.08, 0.23, 0.1, 0.1, 0.1]]

        # Prepare table data
        table_data = [header] + data
        repeat_rows = 1

        # Add status row if provided
        if table_type and count is not None:
            status_row = [[table_type, "", "", "", "", "", "", f"Total: {count}"]]
            table_data = status_row + table_data
            repeat_rows = 2

        # Convert hex #17365d to RGB values (23, 54, 93)
        dark_blue = colors.Color(23 / 255, 54 / 255, 93 / 255)

        table = Table(table_data, repeatRows=repeat_rows, colWidths=col_widths)

        style = [
            # Header styling (for both status row and column headers)
            ('BACKGROUND', (0, 0), (-1, repeat_rows - 1), dark_blue),
            ('TEXTCOLOR', (0, 0), (-1, repeat_rows - 1), colors.white),
            ('FONTNAME', (0, 0), (-1, repeat_rows - 1), 'Calibri-Bold'),
            ('FONTSIZE', (0, 0), (-1, repeat_rows - 1), 12),
            ('BOTTOMPADDING', (0, 0), (-1, repeat_rows - 1), 4),
            ('ROWHEIGHT', (0, 0), (-1, -1), 30),

            # Data rows styling
            ('FONTNAME', (0, repeat_rows), (-1, -1), 'Calibri'),
            ('FONTSIZE', (0, repeat_rows), (-1, -1), 10),
            ('LINEBELOW', (0, 0), (-1, -1), .5, colors.lightgrey),

            # Column alignments
            ('ALIGN', (0, repeat_rows - 1), (0, -1), 'LEFT'),  # FULL NAME
            ('ALIGN', (1, repeat_rows - 1), (1, -1), 'CENTER'),  # GRADE
            ('ALIGN', (2, repeat_rows - 1), (2, repeat_rows - 1), 'CENTER'),  # DAS header
            ('ALIGN', (2, repeat_rows), (2, -1), 'RIGHT'),  # DAS data
            ('ALIGN', (3, repeat_rows - 1), (3, -1), 'LEFT'),  # DAFSC
            ('ALIGN', (4, repeat_rows - 1), (4, repeat_rows - 1), 'CENTER'),  # UNIT header
            ('ALIGN', (4, repeat_rows), (4, -1), 'LEFT'),  # UNIT data
            ('ALIGN', (5, repeat_rows - 1), (5, repeat_rows - 1), 'CENTER'),  # DOR header
            ('ALIGN', (5, repeat_rows), (5, -1), 'RIGHT'),  # DOR data
            ('ALIGN', (6, repeat_rows - 1), (6, -1), 'RIGHT'),  # TAFMSD
            ('ALIGN', (7, repeat_rows - 1), (7, -1), 'RIGHT'),  # PASCODE
        ]

        # Add status row styling if present
        if table_type:
            style.extend([
                ('SPAN', (0, 0), (6, 0)),  # Span ELIGIBLE/INELIGIBLE across
                ('ALIGN', (7, 0), (7, 0), 'RIGHT'),  # Right align the count
            ])

        table.setStyle(TableStyle(style))
        return table

    # Modify the processing section:
    for pascode in unique_pascodes:
        # Filter data for current PASCODE
        pascode_eligible = [row for row in eligible_data if row[7] == pascode]
        pascode_ineligible = [row for row in ineligible_data if row[7] == pascode]

        # If not first PASCODE, add page break
        if elements and pascode != unique_pascodes[0]:
            elements.append(PageBreak())

        # Create eligible section if there are eligible records
        if pascode_eligible:
            elements.append(create_table(
                data=pascode_eligible,
                header=header_row,
                table_type="ELIGIBLE",
                count=len(pascode_eligible)
            ))

        # Add page break before ineligible section
        if pascode_ineligible:
            elements.append(PageBreak())
            elements.append(create_table(
                data=pascode_ineligible,
                header=header_row,
                table_type="INELIGIBLE",
                count=len(pascode_ineligible)
            ))

    # Build PDF
    doc.build(elements)
