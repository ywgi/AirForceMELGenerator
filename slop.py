from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, Image, Frame
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus.flowables import PageBreak
from reportlab.platypus.doctemplate import PageTemplate, BaseDocTemplate
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime
from dateutil.relativedelta import relativedelta
from promotion_eligible_counter import get_promotion_eligibility
import os
from PyPDF2 import PdfMerger
import pandas as pd

# Register Calibri fonts
pdfmetrics.registerFont(TTFont('Calibri', 'Calibri.ttf'))
pdfmetrics.registerFont(TTFont('Calibri-Bold', 'Calibrib.ttf'))

promotion_map = {
    "SRA": "E5",
    "SSG": "E6",
    "TSG": "E7",
    "MSG": "E8",
    "SMS": "E9"
}

SCODs = {'SRA': f'31-MAR',
         'SSG': f'31-JAN',
         'TSG': f'30-NOV',
         'MSG': f'30-SEP',
         'SMS': f'31-JUL'
         }


def get_accounting_date(grade, year):
    scod = f'{SCODs.get(grade)}-{year}'
    formatted_scod_date = datetime.strptime(scod, "%d-%b-%Y")
    accounting_date = formatted_scod_date - relativedelta(days=120 - 1)
    adjusted_accounting_date = accounting_date.replace(day=3).replace(hour=23, minute=59, second=59)
    return adjusted_accounting_date.strftime("%d %b %Y")


def is_within_cycle(date, cycle, year):
    """Check if a date falls within the cycle period"""
    scod_str = f'{SCODs.get(cycle)}-{year}'
    scod_date = datetime.strptime(scod_str, "%d-%b-%Y")
    feb_first = datetime(year=scod_date.year, month=2, day=1)
    return date <= scod_date, date <= feb_first


class MilitaryRosterDocument(BaseDocTemplate):
    def __init__(self, filename, cycle=None, melYear=None, **kwargs):
        super().__init__(filename, **kwargs)
        self.page_width, self.page_height = landscape(letter)
        self.cycle = cycle
        self.melYear = melYear
        # Define the main content frame - start below header
        content_frame = Frame(
            x1=0.5 * inch,  # left margin
            y1=1.05 * inch,  # bottom margin (increased to make room for footer)
            width=self.page_width - inch,  # width with margins
            height=self.page_height - 2.735 * inch,  # reduced height to make room for header
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
        if os.path.exists(logo_path):
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

        if not doc.pas_info['pn'] == 'NA':
            # Promotion Key Title
            canvas.setFont('Calibri-Bold', 12)
            text_start_x = 5 * inch
            canvas.drawString(text_start_x, title_y, "Promotion Eligibility Data")

            # Promotion Key Information
            canvas.setFont('Calibri-Bold', 10)
            line_height = 0.2 * inch
            text_start_y = header_top - 0.1 * inch
            canvas.drawString(text_start_x, text_start_y, f"PROMOTE NOW: {doc.pas_info['pn']}")
            canvas.drawString(text_start_x, text_start_y - line_height, f"MUST PROMOTE: {doc.pas_info['mp']}")

            # Add unit size in the Promotion Eligibility Data section
            unit_size_txt = "SMALL" if doc.pas_info.get('is_small_unit', False) else "LARGE"
            canvas.drawString(text_start_x, text_start_y - 2 * line_height, f"UNIT SIZE: {unit_size_txt}")

        # Signature Block
        canvas.setFont('Calibri-Bold', 12)
        text_start_s = 7.5 * inch
        line_height_s = 0.2 * inch
        title_s = header_top - 0.5 * inch
        officer_name = doc.pas_info['fd name']
        rank = doc.pas_info['rank']
        title = doc.pas_info['title']
        canvas.drawString(text_start_s, title_s, f"{officer_name}, {rank}, USAF")
        canvas.drawString(text_start_s, title_s - line_height_s, title)

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
        canvas.drawString(x, bottom_y, datetime.now().strftime('%d %B %Y'))

        # Center: CUI and identifier
        cui_text = "CUI"
        cui_width = stringWidth(cui_text, 'Calibri-Bold', 12)
        cui_center_x = (self.page_width / 2) - (cui_width / 2)
        canvas.drawString(cui_center_x, bottom_y, cui_text)

        identifier_text = f"{str(self.melYear)[-2:]}{promotion_map[self.cycle]} - Initial MEL"
        identifier_width = stringWidth(identifier_text, 'Calibri-Bold', 12)
        identifier_center_x = (self.page_width / 2) - (identifier_width / 2)
        canvas.drawString(identifier_center_x, bottom_y - 18, identifier_text)

        # Right: Page numbers
        page_text = f"Accounting Date: {get_accounting_date(self.cycle, self.melYear)}"
        page_width = stringWidth(page_text, 'Calibri-Bold', 12)
        canvas.drawString(self.page_width - 0.5 * inch - page_width, bottom_y, page_text)


# Class for the combined small units document
class CombinedSmallUnitsDocument(MilitaryRosterDocument):
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
        if os.path.exists(logo_path):
            logo = Image(logo_path, width=1 * inch, height=1 * inch)
            logo.drawOn(canvas, 0.5 * inch, header_top - 0.8 * inch)

        # Title for combined small units document
        canvas.setFont('Calibri-Bold', 12)
        text_start_x = 2 * inch
        title_y = header_top + 0.1 * inch
        canvas.drawString(text_start_x, title_y, "Senior Rater Listing")

        # Information
        canvas.setFont('Calibri-Bold', 10)
        line_height = 0.2 * inch
        text_start_y = header_top - 0.1 * inch

        canvas.drawString(text_start_x, text_start_y, "Enlisted Force Distribution Panel")
        canvas.drawString(text_start_x, text_start_y - line_height, f"Cycle: {self.cycle} {self.melYear}")
        canvas.drawString(text_start_x, text_start_y - 2 * line_height,
                          f"PASCODE Count: {len(doc.pas_info.get('combined_units', []))}")

        # Promotion Key Title and Information
        canvas.setFont('Calibri-Bold', 12)
        promo_start_x = 5 * inch
        canvas.drawString(promo_start_x, title_y, "Promotion Eligibility Data")

        # Promotion Key Information
        canvas.setFont('Calibri-Bold', 10)
        canvas.drawString(promo_start_x, text_start_y, f"PROMOTE NOW: {doc.pas_info.get('sr_pn', 'NA')}")
        canvas.drawString(promo_start_x, text_start_y - line_height, f"MUST PROMOTE: {doc.pas_info.get('sr_mp', 'NA')}")

        # Signature Block
        canvas.setFont('Calibri-Bold', 12)
        text_start_s = 7.5 * inch
        line_height_s = 0.2 * inch
        title_s = header_top - 0.5 * inch

        if 'sr_name' in doc.pas_info and 'sr_rank' in doc.pas_info and 'sr_title' in doc.pas_info:
            officer_name = doc.pas_info['sr_name']
            rank = doc.pas_info['sr_rank']
            title = doc.pas_info['sr_title']
            canvas.drawString(text_start_s, title_s, f"{officer_name}, {rank}, USAF")
            canvas.drawString(text_start_s, title_s - line_height_s, title)


def calculate_projected_promotion_date(dor, months):
    """Calculate projected promotion date based on DOR and months required."""
    if isinstance(dor, str):
        try:
            dor = datetime.strptime(dor, "%d-%b-%Y")
        except ValueError:
            try:
                dor = pd.to_datetime(dor)
            except:
                return None
    return dor + relativedelta(months=months)


def filter_a1c_for_sra_cycle(eligible_df, ineligible_df, cycle, year):
    """
    Filter A1C members for SRA cycle based on special rules.
    Takes already separated eligible and ineligible DataFrames.
    Returns: updated eligible_df, ineligible_df, btz_validation_df
    """
    # Import necessary libraries
    import pandas as pd
    from datetime import datetime
    from dateutil.relativedelta import relativedelta

    # Only process for SRA cycle
    if cycle != "SRA":
        return eligible_df, ineligible_df, pd.DataFrame()

    # Print initial debugging information
    print("Starting A1C Filter Process")
    print(f"Eligible DataFrame Columns: {list(eligible_df.columns)}")
    print(f"Ineligible DataFrame Columns: {list(ineligible_df.columns)}")

    # Identify key columns with more robust column finding
    def find_column(df, keywords):
        for keyword in keywords:
            cols = [col for col in df.columns if keyword.upper() in col.upper()]
            if cols:
                return cols[0]
        return None

    # Find columns with more specific searches
    pascode_keywords = ['ASSIGNED_PAS']
    grade_keywords = ['GRADE']
    dor_keywords = ['DOR']

    pascode_col = find_column(eligible_df, pascode_keywords) or eligible_df.columns[-1]
    grade_col = find_column(eligible_df, grade_keywords) or eligible_df.columns[1]

    # Identify date columns more comprehensively
    date_cols = [col for col in eligible_df.columns if
                 any(key in col.upper() for key in ['DOR', 'DAS', 'TAFMSD', 'DATE'])]

    # Find Date of Rank column with preference
    dor_col = find_column(eligible_df, dor_keywords) or (date_cols[0] if date_cols else None)

    print(f"Identified Columns:")
    print(f"PASCODE Column: {pascode_col}")
    print(f"GRADE Column: {grade_col}")
    print(f"DOR Column: {dor_col}")
    print(f"Date Columns: {date_cols}")

    # Return early if critical columns are missing
    if dor_col is None or grade_col is None:
        print("Critical columns missing. Aborting A1C filtering.")
        return eligible_df, ineligible_df, pd.DataFrame()

    # Extract A1C members from both DataFrames
    a1c_eligible = eligible_df[eligible_df[grade_col] == 'A1C'].copy()
    a1c_ineligible = ineligible_df[ineligible_df[grade_col] == 'A1C'].copy()

    print(f"A1C Eligible Count: {len(a1c_eligible)}")
    print(f"A1C Ineligible Count: {len(a1c_ineligible)}")

    # Combine A1C members
    all_a1c_df = pd.concat([a1c_eligible, a1c_ineligible])

    # Remove A1C members from original DataFrames
    eligible_df = eligible_df[eligible_df[grade_col] != 'A1C']
    ineligible_df = ineligible_df[ineligible_df[grade_col] != 'A1C']

    # Drop rows with missing critical information
    all_a1c_df.dropna(subset=[pascode_col, dor_col], inplace=True)

    if all_a1c_df.empty:
        print("No A1C members found after filtering.")
        return eligible_df, ineligible_df, pd.DataFrame()

    # Robust date conversion
    for col in date_cols:
        try:
            # Try multiple date formats
            formats = ['%d-%b-%Y', '%Y-%m-%d', '%m/%d/%Y', '%d%b%Y']
            converted = False
            for fmt in formats:
                try:
                    all_a1c_df[col] = pd.to_datetime(all_a1c_df[col], format=fmt, errors='raise')
                    converted = True
                    break
                except:
                    continue

            if not converted:
                all_a1c_df[col] = pd.to_datetime(all_a1c_df[col], errors='coerce')
        except Exception as e:
            print(f"Error converting date column {col}: {e}")
            all_a1c_df[col] = pd.to_datetime(all_a1c_df[col], errors='coerce')

    # Calculate projected promotion dates
    def calculate_projected_promotion_date_safe(date, months):
        """Safe calculation of projected promotion date"""
        if pd.isnull(date):
            return None
        try:
            return date + relativedelta(months=months)
        except Exception as e:
            print(f"Error calculating projected date: {e}")
            return None

    all_a1c_df['PROJ_SRA_DATE'] = all_a1c_df[dor_col].apply(
        lambda x: calculate_projected_promotion_date_safe(x, 28)
    )
    all_a1c_df['PROJ_BTZ_DATE'] = all_a1c_df[dor_col].apply(
        lambda x: calculate_projected_promotion_date_safe(x, 22)
    )

    # Define cycle dates
    scod_date = datetime.strptime(f'31-MAR-{year}', "%d-%b-%Y")
    feb_first = datetime(year=scod_date.year, month=2, day=1)

    # Extra verbose logging for projection dates
    print("Projection Date Details:")
    print(f"SCOD Date: {scod_date}")
    print(f"February 1st: {feb_first}")

    # Filter A1C members
    a1c_eligible = all_a1c_df[
        (all_a1c_df['PROJ_SRA_DATE'].notna()) &
        (all_a1c_df['PROJ_SRA_DATE'] <= feb_first)
        ].copy()

    a1c_ineligible = all_a1c_df[
        (all_a1c_df['PROJ_SRA_DATE'].notna()) &
        (all_a1c_df['PROJ_SRA_DATE'] > feb_first) &
        (all_a1c_df['PROJ_SRA_DATE'] <= scod_date)
        ].copy()

    btz_validation = all_a1c_df[
        (all_a1c_df['PROJ_BTZ_DATE'].notna()) &
        (all_a1c_df['PROJ_BTZ_DATE'] <= scod_date)
        ].copy()

    print("Filtering Results:")
    print(f"A1C Eligible Count: {len(a1c_eligible)}")
    print(f"A1C Ineligible Count: {len(a1c_ineligible)}")
    print(f"BTZ Validation Count: {len(btz_validation)}")

    # Drop temporary columns
    drop_cols = ['PROJ_SRA_DATE', 'PROJ_BTZ_DATE']
    for df in [a1c_eligible, a1c_ineligible, btz_validation]:
        df.drop(columns=drop_cols, inplace=True, errors='ignore')

    # Formatting date columns
    for df in [a1c_eligible, a1c_ineligible, btz_validation, eligible_df, ineligible_df]:
        if not df.empty:
            for col in date_cols:
                if col in df.columns:
                    if pd.api.types.is_datetime64_dtype(df[col]):
                        df[col] = df[col].dt.strftime('%d-%b-%Y').fillna('')
                    else:
                        df[col] = df[col].astype(str).replace({'nan': '', 'NaT': '', 'None': ''})

    # Add filtered A1C members back to appropriate DataFrames
    eligible_df = pd.concat([eligible_df, a1c_eligible], ignore_index=True)
    ineligible_df = pd.concat([ineligible_df, a1c_ineligible], ignore_index=True)

    return eligible_df, ineligible_df, btz_validation


def create_ineligible_table(doc, data, header, table_type=None, count=None):
    """Create table for ineligible members with reason column"""
    table_width = doc.page_width - inch

    # Column widths for ineligible table
    # Format: [NAME, GRADE, PASCODE, DAFSC, UNIT, REASON NOT ELIGIBLE]
    col_widths = [table_width * x for x in [0.25, 0.08, 0.1, 0.1, 0.22, 0.25]]

    # Prepare table data
    table_data = [header] + data
    repeat_rows = 1

    # Add status row if provided
    if table_type and count is not None:
        status_row = [[table_type, "", "", "", "", f"Total: {count}"]]
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
        ('ALIGN', (2, repeat_rows - 1), (2, -1), 'CENTER'),  # PASCODE
        ('ALIGN', (3, repeat_rows - 1), (3, -1), 'LEFT'),  # DAFSC
        ('ALIGN', (4, repeat_rows - 1), (4, -1), 'LEFT'),  # UNIT
        ('ALIGN', (5, repeat_rows - 1), (5, -1), 'LEFT'),  # REASON NOT ELIGIBLE
    ]

    # Add status row styling if present
    if table_type:
        style.extend([
            ('SPAN', (0, 0), (4, 0)),  # Span heading across
            ('ALIGN', (5, 0), (5, 0), 'RIGHT'),  # Right align the count
        ])

    table.setStyle(TableStyle(style))
    return table

def create_table(doc, data, header, table_type=None, count=None):
    """Create table with optional status row"""
    # Type checking - ensure data is a list of lists
    if not isinstance(data, list):
        data = []
    elif data and not isinstance(data[0], list):
        if isinstance(data, list):
            data = [data]
        else:
            data = []

    # Filter out any rows with None/NaT values in the PASCODE column (last column)
    filtered_data = []
    for row in data:
        if len(row) > 7 and row[7] is not None and row[7] != 'None' and row[7] != 'NaT' and row[7] != 'nan':
            # Process name column if needed
            if len(row) > 0 and isinstance(row[0], str) and len(row[0]) > 30:
                row_copy = list(row)  # Create a copy of the row
                row_copy[0] = row[0][:27] + "..."  # Truncate name
                filtered_data.append(row_copy)
            else:
                filtered_data.append(row)

    # Use filtered data instead of original
    data = filtered_data

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


def generate_pascode_pdf(eligible_data, ineligible_data, cycle, melYear, pascode, pas_info,
                         output_filename, logo_path, btz_data=None):
    """Generate a PDF for a single pascode with optional BTZ validation section"""

    # Debugging print
    print(f"Generating PDF for PASCODE: {pascode}")
    print(f"Eligible Data Count: {len(eligible_data) if eligible_data else 0}")
    print(f"Ineligible Data Count: {len(ineligible_data) if ineligible_data else 0}")
    print(f"BTZ Data Count: {len(btz_data) if btz_data else 0}")

    name_idx = 0  # FULL_NAME
    grade_idx = 1  # GRADE
    dafsc_idx = 3  # DAFSC
    unit_idx = 4  # UNIT
    pascode_idx = 7  # ASSIGNED_PAS/PASCODE

    # Validate and normalize inputs
    def validate_data_list(data):
        if not isinstance(data, list):
            return []
        # Filter out None or empty rows
        return [row for row in data if row and any(cell is not None and str(cell).strip() for cell in row)]

    eligible_data = validate_data_list(eligible_data)
    ineligible_data = validate_data_list(ineligible_data)
    btz_data = validate_data_list(btz_data)

    # Detailed BTZ data processing
    if btz_data:
        processed_btz_data = []
        for row in btz_data:
            try:
                # Ensure row has enough elements
                if len(row) < max(name_idx, grade_idx, dafsc_idx, unit_idx, pascode_idx) + 1:
                    print(f"Skipping BTZ row due to insufficient data: {row}")
                    continue

                # Create a processed row with standard format
                processed_row = [
                    str(row[name_idx])[:30] if name_idx < len(row) else "Unknown",
                    str(row[grade_idx]) if grade_idx < len(row) else "Unknown",
                    row[2] if len(row) > 2 else "N/A",  # DAS
                    str(row[dafsc_idx]) if dafsc_idx < len(row) else "Unknown",
                    str(row[unit_idx]) if unit_idx < len(row) else "Unknown",
                    row[5] if len(row) > 5 else "N/A",  # DOR
                    row[6] if len(row) > 6 else "N/A",  # TAFMSD
                    str(row[pascode_idx]) if pascode_idx < len(row) else pascode
                ]
                processed_btz_data.append(processed_row)
            except Exception as e:
                print(f"Error processing BTZ row: {row}")
                print(f"Error details: {e}")

        # Update btz_data with processed rows
        btz_data = processed_btz_data

    # Determine if this is a small unit (10 or fewer eligible members)
    is_small_unit = len(eligible_data) <= 10

    # Update the pas_info with the small unit flag
    pas_info['is_small_unit'] = is_small_unit

    # Create document
    doc = MilitaryRosterDocument(
        output_filename,
        cycle=cycle,
        melYear=melYear,
        pagesize=landscape(letter),
        rightMargin=0.5 * inch,
        leftMargin=0.5 * inch,
        topMargin=0.5 * inch,
        bottomMargin=0.5 * inch
    )

    # Store additional information
    doc.logo_path = logo_path
    doc.pas_members = [pas_info]  # Only include this pascode's info
    doc.pas_info = pas_info  # Set directly

    elements = []
    header_row = ['FULL NAME', 'GRADE', 'DAS', 'DAFSC', 'UNIT', 'DOR', 'TAFMSD', 'PASCODE']
    ineligible_header_row = ['FULL NAME', 'GRADE', 'PASCODE', 'DAFSC', 'UNIT', 'REASON NOT ELIGIBLE']

    # Create eligible section if there are eligible records
    if eligible_data:
        table = create_table(
            doc,
            data=eligible_data,
            header=header_row,
            table_type="ELIGIBLE",
            count=len(eligible_data)
        )
        elements.append(table)

    # Process ineligible data
    processed_ineligible_data = []
    if ineligible_data:
        # Add page break before ineligible section if needed
        if elements:
            elements.append(PageBreak())

        for row in ineligible_data:
            # Make sure the row has enough elements
            if len(row) <= max(name_idx, grade_idx, dafsc_idx, unit_idx, pascode_idx):
                continue

            # Extract data safely with defaults for missing fields
            name = str(row[name_idx]) if name_idx < len(row) else "Unknown"
            if len(name) > 30:
                name = name[:27] + "..."

            grade = str(row[grade_idx]) if grade_idx < len(row) else "Unknown"
            dafsc = str(row[dafsc_idx]) if dafsc_idx < len(row) else "Unknown"
            unit = str(row[unit_idx]) if unit_idx < len(row) else "Unknown"
            pascode_val = str(row[pascode_idx]) if pascode_idx < len(row) else "Unknown"

            # Determine reason
            reason = "Ineligible"
            if isinstance(row, pd.Series) and 'REASON' in row.index and pd.notna(row['REASON']):
                # Use the specific reason we added
                reason = str(row['REASON'])
            elif len(row) > 8 and row[8] is not None:
                reason = str(row[8])
            elif len(row) > 5 and isinstance(row[5], str):
                reason = row[5]

            # Format: [NAME, GRADE, PASCODE, DAFSC, UNIT, REASON]
            new_row = [name, grade, pascode_val, dafsc, unit, reason]
            processed_ineligible_data.append(new_row)

        # Create ineligible table
        if processed_ineligible_data:
            table = create_ineligible_table(
                doc,
                data=processed_ineligible_data,
                header=ineligible_header_row,
                table_type="INELIGIBLE",
                count=len(processed_ineligible_data)
            )
            elements.append(table)

    # Add BTZ validation section if data provided
    if btz_data:
        # Always add a page break before BTZ section, regardless of previous content
        elements.append(PageBreak())

        print(f"Adding BTZ Validation Table with {len(btz_data)} rows")
        table = create_table(
            doc,
            data=btz_data,
            header=header_row,
            table_type="VERIFY MEMBER ELIGIBILITY",
            count=len(btz_data)
        )
        elements.append(table)

    # Debug print before building
    print(f"Total elements to build: {len(elements)}")

    # Build PDF for this pascode
    doc.build(elements)

    return output_filename, is_small_unit, eligible_data

def generate_combined_small_units_pdf(small_units_data, cycle, melYear, output_filename, logo_path, sr_info=None):
    """Generate a PDF combining all small units eligible personnel"""
    # Extract all eligible members from small units
    combined_eligible_data = []
    combined_pascode_info = []

    for pascode, data in small_units_data.items():
        eligible_members = data['eligible_data']
        pas_info = data['pas_info']
        combined_eligible_data.extend(eligible_members)
        combined_pascode_info.append(f"{pascode} ({pas_info['fd name']})")

    if not combined_eligible_data:
        return None  # No data to create PDF

    # Sort the combined data by FULL_NAME (index 0)
    combined_eligible_data.sort(key=lambda x: x[0] if x and len(x) > 0 and x[0] else "")

    # Calculate promotion eligibility for the combined count
    from promotion_eligible_counter import get_promotion_eligibility
    sr_mp, sr_pn = get_promotion_eligibility(len(combined_eligible_data), cycle)

    # If no sr_info is provided, use default values
    if sr_info is None:
        sr_info = {
            'sr_name': 'Senior Rater',
            'sr_rank': 'Col',
            'sr_title': 'Senior Rater'
        }

    # Create combined document
    doc = CombinedSmallUnitsDocument(
        output_filename,
        cycle=cycle,
        melYear=melYear,
        pagesize=landscape(letter),
        rightMargin=0.5 * inch,
        leftMargin=0.5 * inch,
        topMargin=0.5 * inch,
        bottomMargin=0.5 * inch
    )

    # Store additional information
    doc.logo_path = logo_path
    doc.pas_info = {
        'combined_units': combined_pascode_info,
        'fd name': 'Combined Small Units',
        'rank': 'N/A',
        'title': 'Combined Report',
        'srid': 'N/A',
        'fdid': 'COMBINED',
        'srid mpf': 'N/A',
        'mp': 'N/A',
        'pn': 'N/A',
        'sr_name': sr_info['sr_name'],
        'sr_rank': sr_info['sr_rank'],
        'sr_title': sr_info['sr_title'],
        'sr_mp': sr_mp,
        'sr_pn': sr_pn
    }

    elements = []
    header_row = ['FULL NAME', 'GRADE', 'DAS', 'DAFSC', 'UNIT', 'DOR', 'TAFMSD', 'PASCODE']

    # Create combined eligible section with updated table type
    table = create_table(
        doc,
        data=combined_eligible_data,
        header=header_row,
        table_type="ELIGIBLES (SENIOR RATER)",
        count=len(combined_eligible_data)
    )
    elements.append(table)

    # Build PDF
    doc.build(elements)
    return output_filename


def merge_pdfs(input_pdfs, output_pdf):
    """Merge multiple PDFs into a single PDF"""
    merger = PdfMerger()

    # Verify all files exist first
    input_pdfs = [pdf for pdf in input_pdfs if os.path.exists(pdf)]

    if not input_pdfs:
        return

    # Add each PDF to the merger in the provided order
    for pdf in input_pdfs:
        try:
            merger.append(pdf)
        except Exception as e:
            continue

    # Write the merged PDF to output file
    try:
        merger.write(output_pdf)
        merger.close()
    except Exception as e:
        raise


def generate_roster_pdf(eligible_df, ineligible_df, cycle, melYear, pascode_map,
                        output_filename="military_roster.pdf",
                        logo_path='images/Air_Force_Personnel_Center.png'):
    """Generate a military roster PDF with proper handling of invalid values and small units"""

    # Clean DataFrames of rows with NaT in PASCODE column (assumed to be the last column)
    pascode_col = eligible_df.columns[-1]
    eligible_df = eligible_df.dropna(subset=[pascode_col])
    ineligible_df = ineligible_df.dropna(subset=[pascode_col])

    # Apply A1C filtering logic if this is an SRA cycle
    btz_validation_df = pd.DataFrame()
    if cycle == "SRA":
        eligible_df, ineligible_df, btz_validation_df = filter_a1c_for_sra_cycle(
            eligible_df, ineligible_df, cycle, melYear)

    # Convert DataFrames to lists
    eligible_data = eligible_df.values.tolist()
    ineligible_data = ineligible_df.values.tolist()
    btz_data = btz_validation_df.values.tolist() if not btz_validation_df.empty else []

    # Final filter to remove any rows with problematic PASCODE values
    eligible_data = [row for row in eligible_data if
                     row and len(row) > 7 and row[7] and str(row[7]) != 'nan' and str(row[7]) != 'NaT']
    ineligible_data = [row for row in ineligible_data if
                       row and len(row) > 7 and row[7] and str(row[7]) != 'nan' and str(row[7]) != 'NaT']
    btz_data = [row for row in btz_data if
                row and len(row) > 7 and row[7] and str(row[7]) != 'nan' and str(row[7]) != 'NaT']

    # Get unique PASCODEs from data
    unique_pascodes = set()
    for row in eligible_data + ineligible_data + btz_data:
        if row and len(row) > 7 and row[7]:
            unique_pascodes.add(str(row[7]))
    unique_pascodes = sorted(list(unique_pascodes))

    # Dictionary to store temporary PDFs by PASCODE to ensure order
    temp_pdf_dict = {}

    # Store data for small units (10 or fewer eligible members)
    small_units_data = {}

    # Generate a separate PDF for each pascode
    for pascode in unique_pascodes:
        # Skip if this pascode is not in the pascode_map
        if pascode not in pascode_map:
            continue

        # Filter data for current pascode
        pascode_eligible = [row for row in eligible_data if str(row[7]) == pascode]
        pascode_ineligible = [row for row in ineligible_data if str(row[7]) == pascode]
        pascode_btz = [row for row in btz_data if str(row[7]) == pascode]

        # Skip if there's no data for this pascode
        if not pascode_eligible and not pascode_ineligible and not pascode_btz:
            continue

        # Create PAS info for this pascode
        eligible_candidates = len(pascode_eligible)
        must_promote, promote_now = get_promotion_eligibility(eligible_candidates, cycle)
        pas_info = {
            'srid': pascode_map[pascode][3],
            'fd name': pascode_map[pascode][0],
            'rank': pascode_map[pascode][1],
            'title': pascode_map[pascode][2],
            'fdid': f'{pascode_map[pascode][3]}{pascode[-4:]}',
            'srid mpf': pascode[:2],
            'mp': must_promote,
            'pn': promote_now
        }

        # Create temporary filename
        temp_filename = f"temp_{pascode}.pdf"

        # Generate PDF for this pascode - returns tuple with filename, is_small_unit flag, and eligible data
        pdf_result, is_small_unit, eligible_members = generate_pascode_pdf(
            pascode_eligible,
            pascode_ineligible,
            cycle,
            melYear,
            pascode,
            pas_info,
            temp_filename,
            logo_path,
            pascode_btz
        )

        # Store in dictionary using pascode as key to maintain order
        temp_pdf_dict[pascode] = pdf_result

        # If this is a small unit, store its data for the combined report
        if is_small_unit and eligible_members:
            small_units_data[pascode] = {
                'eligible_data': eligible_members,
                'pas_info': pas_info
            }

    # Create combined small units PDF if there are any small units
    combined_small_units_pdf = None
    if small_units_data:
        # Prompt for Senior Rater information
        print("\nEnter Senior Rater information for Combined Small Units Report:")
        sr_name = input("Senior Rater Name: ")
        sr_rank = input("Senior Rater Rank: ")
        sr_title = input("Senior Rater Title: ")

        sr_info = {
            'sr_name': sr_name,
            'sr_rank': sr_rank,
            'sr_title': sr_title
        }

        combined_small_units_pdf = generate_combined_small_units_pdf(
            small_units_data,
            cycle,
            melYear,
            f"temp_combined_small_units.pdf",
            logo_path,
            sr_info
        )

        # Add the combined PDF to the list of PDFs to merge, if it was created
        if combined_small_units_pdf and os.path.exists(combined_small_units_pdf):
            temp_pdf_dict['COMBINED'] = combined_small_units_pdf

    # Convert dictionary to ordered list based on sorted pascodes
    # Ensure 'COMBINED' small units PDF is last if it exists
    temp_pdfs = []
    for pascode in unique_pascodes:
        if pascode in temp_pdf_dict:
            temp_pdfs.append(temp_pdf_dict[pascode])

    # Add combined small units PDF at the end if it exists
    if 'COMBINED' in temp_pdf_dict:
        temp_pdfs.append(temp_pdf_dict['COMBINED'])

    # Merge all the temporary PDFs into the final output file
    if temp_pdfs:
        merge_pdfs(temp_pdfs, output_filename)

        # Clean up temporary files
        for pdf in temp_pdfs:
            try:
                if os.path.exists(pdf):
                    os.remove(pdf)
            except:
                pass

    return output_filename














from datetime import datetime
from dateutil.relativedelta import relativedelta

# Static closeout date = annual report due date
SCODs = {'SRA': f'31-MAR',
         'SSG': f'31-JAN',
         'TSG': f'30-NOV',
         'MSG': f'30-SEP',
         'SMS': f'31-JUL'
         }

# Time in grade = how long you've been in a rank
TIG = {'SRA': f'01-AUG',
       'SSG': f'01-JUL',
       'TSG': f'01-MAY',
       'MSG': f'01-MAR',
       'SMS': f'01-DEC'
       }

tig_months_required = {
    'SRA': 6,
    'SSG': 23,
    'TSG': 24,
    'MSG': 20,
    'SMS': 21
}

# Total active federal military service date = time in military (years)
TAFMSD = {
    'SRA': 3,
    'SSG': 5,
    'TSG': 8,
    'MSG': 11,
    'SMS': 14
}

# manditory date of separation = the day you have to exit the military
mdos = {
    'SRA': f'01-SEP',
    'SSG': f'01-AUG',
    'TSG': f'01-AUG',
    'MSG': f'01-APR',
    'SMS': f'01-JAN'
}

main_higher_tenure = {
    'SRA': 10,
    'SSG': 20,
    'TSG': 22,
    'MSG': 24,
    'SMS': 26
}

exception_higher_tenure = {
    'SRA': 12,
    'SSG': 22,
    'TSG': 24,
    'MSG': 26,
    'SMS': 28
}

# reenlistment codes
re_codes = {
    "2A": "HQ AFPC denied reenlistment for quality reasons.",
    "2B": "Discharged under general conditions.",
    "2C": "Involuntary separation with honorable discharge.",
    "2F": "Undergoing rehabilitation in a DOD facility.",
    "2G": "Failed Substance Abuse Treatment for drugs.",
    "2H": "Failed Substance Abuse Treatment for alcohol.",
    "2J": "Under investigation, may result in discharge.",
    "2K": "Notified of involuntary separation.",
    "2M": "Serving or separated while under sentence.",
    "2P": "AWOL; deserter.",
    "2W": "Retired and recalled to active duty.",
    "2X": "Not selected for reenlistment.",
    "4H": "Ineligible due to Article 15.",
    "4I": "Ineligible due to Control Roster.",
    "4J": "Ineligible due to AF Weight Management Program.",
    "4K": "Medically disqualified or pending evaluation.",
    "4L": "Separated from a commissioning program.",
    "4M": "Breach of enlistment agreement.",
    "4N": "Convicted by civil authority."
}

exception_hyt_start_date = datetime(2023, 12, 8)
exception_hyt_end_date = datetime(2025, 9, 30)


def board_filter(grade, year, date_of_rank, uif_code, uif_disposition_date, tafmsd, re_status, cafsc):
    formatted_date_of_rank = datetime.strptime(date_of_rank, "%d-%b-%Y")
    formatted_tafmsd = datetime.strptime(tafmsd, "%d-%b-%Y")
    scod = f'{SCODs.get(grade)}-{year}'
    formatted_scod = datetime.strptime(scod, "%d-%b-%Y")
    tig_selection_month = f'{TIG.get(grade)}-{year}'
    formatted_tig_selection_month = datetime.strptime(tig_selection_month, "%d-%b-%Y")
    tig_eligibility_month = formatted_tig_selection_month - relativedelta(months=tig_months_required.get(grade))
    tafmsd_required_date = formatted_tig_selection_month - relativedelta(years=TAFMSD.get(grade) - 1)
    hyt_date = datetime.strptime(tafmsd, "%d-%b-%Y") + relativedelta(years=main_higher_tenure.get(grade))
    mdos_date = formatted_tig_selection_month + relativedelta(months=1)

    # Add safe datetime conversion for UIF disposition date
    try:
        if uif_disposition_date:
            formatted_uif_disposition_date = datetime.strptime(uif_disposition_date, "%d-%b-%Y")
        else:
            formatted_uif_disposition_date = None
    except (ValueError, TypeError):
        formatted_uif_disposition_date = None

    # Check eligibility criteria and return specific reason if not eligible
    if formatted_date_of_rank > tig_eligibility_month:
        return f"TIG: <{tig_months_required.get(grade)} months"

    if formatted_tafmsd > tafmsd_required_date:
        return f"TIS: <{TAFMSD.get(grade)} years"

    if exception_hyt_start_date < hyt_date < exception_hyt_end_date:
        hyt_date += relativedelta(years=2)

    if hyt_date < mdos_date:
        return "HYT: Mandatory DOS"

    if uif_code and uif_code > 1 and formatted_uif_disposition_date and formatted_uif_disposition_date < formatted_scod:
        return f"UIF: Code {uif_code}"

    if re_status and re_status in re_codes:
        # Get the description and truncate it if needed
        description = re_codes[re_status]
        short_desc = description[:30] + "..." if len(description) > 30 else description
        return f"RE {re_status}: {short_desc}"

    # Check CAFSC skill level requirements
    if cafsc and isinstance(cafsc, str) and len(cafsc) >= 5:
        # Skip 8X000 career fields (exempt from the skill level requirements)
        if not cafsc.startswith('8'):
            # Get the skill level (5th position)
            skill_level = cafsc[4] if len(cafsc) > 4 else '0'

            # A1C requires 5-level for SrA
            if grade == 'A1C' and skill_level < '5':
                return "AFSC: Requires 5-skill level"

            # SSG requires 7-level for TSG
            if grade == 'SSG' and skill_level < '7':
                return "AFSC: Requires 7-skill level"

            # MSG requires 9-level for SMS
            if grade == 'MSG' and skill_level < '9':
                return "AFSC: Requires 9-skill level"

    # If all checks pass, return True for eligible
    return True