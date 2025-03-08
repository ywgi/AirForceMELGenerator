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

        # Get eligible row count (you'll need to pass this in or calculate it)
        eligible_count = len(self.eligible_rows) if hasattr(self, 'eligible_rows') else 0
        counter_text = f"Total: {eligible_count}"

        # Create thicker header border with background
        canvas.setFillColorRGB(0, 0.32, 0.65)  # Navy Blue color
        canvas.setStrokeColorRGB(0, 0.32, 0.65)  # Navy Blue color
        canvas.rect(
            0 * inch,  # left margin
            self.page_height - 2.0 * inch,  # bottom of header border
            self.page_width,  # width
            0.3 * inch,  # height of border
            fill=1
        )

        # Add "Eligible" text in white color on left
        canvas.setFont("Calibri-Bold", 12)
        canvas.setFillColorRGB(1, 1, 1)  # White color
        canvas.drawString(
            0.5 * inch,  # Left margin
            self.page_height - 1.9 * inch,  # Centered vertically in the border
            "ELIGIBLE"
        )

        # Add counter text in white color on right
        # Get width of counter text to position it correctly
        counter_width = canvas.stringWidth(counter_text, "Calibri-Bold", 12)
        canvas.drawString(
            self.page_width - counter_width - 0.5 * inch,  # Right margin
            self.page_height - 1.9 * inch,  # Same height as ELIGIBLE text
            counter_text
        )

        # Add footer top border
        canvas.setLineWidth(2.5)
        canvas.setStrokeColorRGB(0, 0.32, 0.65)
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

    # Combine header and data rows
    table_data = [header_row] + data_rows

    # Create main table
    table = Table(table_data)
    table.setStyle(TableStyle([
        # Header styling
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('FONTNAME', (0, 0), (-1, 0), 'Calibri-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('TOPPADDING', (0, 0), (-1, 0), 4),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 4),
        ('ROWHEIGHT', (0, 0), (-1, -1), 30),  # makes all rows 30 points high

        # Data rows styling
        ('FONTNAME', (0, 1), (-1, -1), 'Calibri'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('LINEBELOW', (0, 0), (-1, -1), .5, colors.lightgrey),

        # FULL NAME column alignment (Left)
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),

        # GRADE column alignment (Center)
        ('ALIGN', (1, 0), (1, -1), 'CENTER'),

        # DAS column alignment (Center for header, Right for data)
        ('ALIGN', (2, 0), (2, 0), 'CENTER'),  # Header
        ('ALIGN', (2, 1), (2, -1), 'RIGHT'),  # Data rows

        # DAFSC column alignment (Left)
        ('ALIGN', (3, 0), (3, -1), 'LEFT'),

        # UNIT column alignment (Center for header, Left for data)
        ('ALIGN', (4, 0), (4, 0), 'CENTER'),  # Header
        ('ALIGN', (4, 1), (4, -1), 'LEFT'),  # Data rows

        # DOR column alignment (Center for header, Right for data)
        ('ALIGN', (5, 0), (5, 0), 'CENTER'),  # Header
        ('ALIGN', (5, 1), (5, -1), 'RIGHT'),  # Data rows

        # TAFMSD column alignment (Right)
        ('ALIGN', (6, 0), (6, -1), 'RIGHT'),

        # PASCODE column alignment (Right)
        ('ALIGN', (7, 0), (7, -1), 'RIGHT'),
    ]))

    elements.append(table)

    # Build PDF
    doc.build(elements)

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
            # canvas.setLineWidth(2.5)
            # canvas.setStrokeColorRGB(0, 0.32, 0.65)  # Navy Blue color
            # canvas.line(
            #     0.5 * inch,  # left margin
            #     self.page_height - 1.7 * inch,  # just below header
            #     self.page_width - 9.0 * inch,  # right margin
            #     self.page_height - 1.7 * inch
            # )

            # Create blue rectangle
            canvas.setFillColorRGB(0.0902, 0.2118, 0.3647)  # Navy Blue color based on #17365d
            canvas.setStrokeColorRGB(0.0902, 0.2118, 0.3647)  # Navy Blue color based on #17365d
            canvas.rect(
                0.5 * inch,  # left margin (0.5 inches)
                self.page_height - 2.0 * inch,  # 2 inches from top
                self.page_width - inch,  # width minus margins (0.5 inches on each side)
                0.3 * inch,  # height of rectangle
                fill=1,  # fill the rectangle
                stroke=0  # no border
            )

            # Add footer top border with matching margins as the rectangle
            canvas.setLineWidth(0.1)  # Set the line width (keeping it thin as previously discussed)
            canvas.setStrokeColorRGB(0, 0, 0)  # Black color (keeping the same color)
            canvas.line(
                0.5 * inch,  # left margin (0.5 inches)
                1.2 * inch,  # just above footer
                self.page_width - 0.5 * inch,  # right margin, adjusted to match the rectangle
                1.2 * inch  # same y-coordinate for a straight line
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

            canvas.setFont('Calibri-Bold', 8)
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
                line_width = stringWidth(line, 'Helvetica', 8)  # Default font, change if needed
                center_x = (self.page_width - line_width) / 2
                canvas.drawString(center_x, y + (len(lines) - 1 - i) * 10, line)

            # Add the bottom footer line with three elements
            bottom_y = 0.3 * inch  # Position for the bottom line

            # Set all text to black
            canvas.setFillColorRGB(0, 0, 0)  # Set color to black

            # Left aligned date (placeholder - you can modify this)
            canvas.setFont('Calibri-Bold', 12)  # Set font to Calibri-Bold 12 pt
            canvas.drawString(x, bottom_y, "15 March 2024")

            # Center aligned "CUI" in Calibri-Bold
            cui_text = "CUI"
            cui_width = stringWidth(cui_text, 'Calibri-Bold', 12)
            cui_center_x = (self.page_width / 2) - (cui_width / 2)
            canvas.drawString(cui_center_x, bottom_y, cui_text)

            # Add "25E7 - Final MEL" below "CUI" in Calibri-Bold
            identifier_text = "25E7 - Final MEL"
            identifier_width = stringWidth(identifier_text, 'Calibri-Bold', 12)
            identifier_center_x = (self.page_width / 2) - (identifier_width / 2)
            canvas.drawString(identifier_center_x, bottom_y - 18, identifier_text)  # Adjust y-position for 12 pt font

            # Right aligned page numbers
            page_text = f"Page {doc.page} of {doc.page}"
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
            'srid mpf': '0P'
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

        def create_table(data, header):
            # Set table width to match page width minus margins
            table_width = doc.page_width - inch  # 0.5 inch margins on each side
            col_widths = [table_width * x for x in [0.25, 0.08, 0.1, 0.1, 0.2, 0.1, 0.1, 0.07]]  # Proportional widths

            table = Table([header] + data, repeatRows=1, colWidths=col_widths)
            table.setStyle(TableStyle([
                # Header styling
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('FONTNAME', (0, 0), (-1, 0), 'Calibri-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 4),
                ('ROWHEIGHT', (0, 0), (-1, -1), 30),

                # Data rows styling
                ('FONTNAME', (0, 1), (-1, -1), 'Calibri'),
                ('FONTSIZE', (0, 1), (-1, -1), 10),
                ('LINEBELOW', (0, 0), (-1, -1), .5, colors.lightgrey),

                # Column alignments
                ('ALIGN', (0, 0), (0, -1), 'LEFT'),  # FULL NAME
                ('ALIGN', (1, 0), (1, -1), 'CENTER'),  # GRADE
                ('ALIGN', (2, 0), (2, 0), 'CENTER'),  # DAS header
                ('ALIGN', (2, 1), (2, -1), 'RIGHT'),  # DAS data
                ('ALIGN', (3, 0), (3, -1), 'LEFT'),  # DAFSC
                ('ALIGN', (4, 0), (4, 0), 'CENTER'),  # UNIT header
                ('ALIGN', (4, 1), (4, -1), 'LEFT'),  # UNIT data
                ('ALIGN', (5, 0), (5, 0), 'CENTER'),  # DOR header
                ('ALIGN', (5, 1), (5, -1), 'RIGHT'),  # DOR data
                ('ALIGN', (6, 0), (6, -1), 'RIGHT'),  # TAFMSD
                ('ALIGN', (7, 0), (7, -1), 'RIGHT'),  # PASCODE
            ]))
            return table

        # Process each PASCODE
        for pascode in unique_pascodes:
            # Filter data for current PASCODE
            pascode_eligible = [row for row in eligible_data if row[7] == pascode]
            pascode_ineligible = [row for row in ineligible_data if row[7] == pascode]

            # Store eligible count for the counter
            doc.eligible_rows = pascode_eligible

            # If not first PASCODE, add page break
            if elements and pascode != unique_pascodes[0]:
                elements.append(PageBreak())

            # Create eligible section if there are eligible records
            if pascode_eligible:
                # Add blue header with "ELIGIBLE" and count
                elements.append(Paragraph(
                    f'<para textColor="white" backColor="17365d" leftIndent="36" rightIndent="36" fontSize="12" fontName="Calibri-Bold">'
                    f'ELIGIBLE<spacer length="350"/>Total: {len(pascode_eligible)}</para>',
                    ParagraphStyle('header')
                ))
                elements.append(Spacer(1, 0.2 * inch))
                elements.append(create_table(pascode_eligible, header_row))

            # Add page break before ineligible section
            if pascode_ineligible:
                elements.append(PageBreak())
                # Add blue header with "INELIGIBLE" and count
                elements.append(Paragraph(
                    f'<para textColor="white" backColor="17365d" leftIndent="36" rightIndent="36" fontSize="12" fontName="Calibri-Bold">'
                    f'INELIGIBLE<spacer length="350"/>Total: {len(pascode_ineligible)}</para>',
                    ParagraphStyle('header')
                ))
                elements.append(Spacer(1, 0.2 * inch))
                elements.append(create_table(pascode_ineligible, header_row))

        # Build PDF
        doc.build(elements)

