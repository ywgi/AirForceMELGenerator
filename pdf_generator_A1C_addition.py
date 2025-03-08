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
    # Only process for SRA cycle
    if cycle != "SRA":
        return eligible_df, ineligible_df, pd.DataFrame()

    # Make deep copies to avoid modifying originals
    eligible_df = eligible_df.copy(deep=True)
    ineligible_df = ineligible_df.copy(deep=True)

    # Identify key columns
    pascode_col = next((col for col in eligible_df.columns if 'PAS' in col.upper() or 'CODE' in col.upper()),
                       eligible_df.columns[-1])
    grade_col = next((col for col in eligible_df.columns if 'GRADE' in col.upper() or 'RANK' in col.upper()),
                     eligible_df.columns[1])

    # Identify date columns
    date_cols = [col for col in eligible_df.columns if
                 any(key in col.upper() for key in ['DOR', 'DAS', 'TAFMSD', 'DATE'])]
    dor_col = next((col for col in date_cols if 'DOR' in col.upper()), date_cols[0] if date_cols else None)

    # Return early if no DOR column found
    if dor_col is None:
        return eligible_df, ineligible_df, pd.DataFrame()

    # Extract A1C members from both DataFrames
    a1c_eligible = eligible_df[eligible_df[grade_col] == 'A1C'].copy()
    a1c_ineligible = ineligible_df[ineligible_df[grade_col] == 'A1C'].copy()
    all_a1c_df = pd.concat([a1c_eligible, a1c_ineligible])

    # Remove A1C members from original DataFrames
    eligible_df = eligible_df[eligible_df[grade_col] != 'A1C']
    ineligible_df = ineligible_df[ineligible_df[grade_col] != 'A1C']

    # Drop rows with missing PASCODE or DOR
    all_a1c_df.dropna(subset=[pascode_col, dor_col], inplace=True)

    if all_a1c_df.empty:
        return eligible_df, ineligible_df, pd.DataFrame()

    # Convert date columns to datetime using safe approach
    for col in date_cols:
        try:
            all_a1c_df[col] = pd.to_datetime(all_a1c_df[col], format='%d-%b-%Y', errors='coerce')
        except:
            all_a1c_df[col] = pd.to_datetime(all_a1c_df[col], errors='coerce')

    # Calculate projected promotion dates
    all_a1c_df['PROJ_SRA_DATE'] = all_a1c_df[dor_col].apply(
        lambda x: calculate_projected_promotion_date(x, 28) if pd.notna(x) else None
    )
    all_a1c_df['PROJ_BTZ_DATE'] = all_a1c_df[dor_col].apply(
        lambda x: calculate_projected_promotion_date(x, 22) if pd.notna(x) else None
    )

    # Define cycle dates
    scod_date = datetime.strptime(f'31-MAR-{year}', "%d-%b-%Y")
    feb_first = datetime(year=scod_date.year, month=2, day=1)

    # Filter A1C members based on projected promotion dates
    a1c_eligible = all_a1c_df[all_a1c_df['PROJ_SRA_DATE'].notna() &
                              (all_a1c_df['PROJ_SRA_DATE'] <= feb_first)].copy()

    a1c_ineligible = all_a1c_df[all_a1c_df['PROJ_SRA_DATE'].notna() &
                                (all_a1c_df['PROJ_SRA_DATE'] > feb_first) &
                                (all_a1c_df['PROJ_SRA_DATE'] <= scod_date)].copy()

    btz_validation = all_a1c_df[all_a1c_df['PROJ_BTZ_DATE'].notna() &
                                (all_a1c_df['PROJ_BTZ_DATE'] <= scod_date)].copy()

    # Drop temporary columns before merging back
    drop_cols = ['PROJ_SRA_DATE', 'PROJ_BTZ_DATE']
    a1c_eligible.drop(columns=drop_cols, inplace=True, errors='ignore')
    a1c_ineligible.drop(columns=drop_cols, inplace=True, errors='ignore')
    btz_validation.drop(columns=drop_cols, inplace=True, errors='ignore')

    # Format date columns safely - check type first
    for df in [a1c_eligible, a1c_ineligible, btz_validation]:
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

    # Format date columns in the combined DataFrames as well
    for df in [eligible_df, ineligible_df]:
        if not df.empty:
            for col in date_cols:
                if col in df.columns:
                    if pd.api.types.is_datetime64_dtype(df[col]):
                        df[col] = df[col].dt.strftime('%d-%b-%Y').fillna('')
                    else:
                        df[col] = df[col].astype(str).replace({'nan': '', 'NaT': '', 'None': ''})

    return eligible_df, ineligible_df, btz_validation


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
    # Validate inputs are proper lists
    if not isinstance(eligible_data, list):
        eligible_data = []
    if not isinstance(ineligible_data, list):
        ineligible_data = []
    if btz_data is not None and not isinstance(btz_data, list):
        btz_data = []

    # Determine if this is a small unit (10 or fewer eligible members)
    is_small_unit = len(eligible_data) <= 10

    # Update the pas_info with the small unit flag
    pas_info['is_small_unit'] = is_small_unit

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

    # Add page break before ineligible section
    if ineligible_data:
        # Process ineligible data to ensure exactly 8 columns
        processed_ineligible_data = [
            row[:8] if len(row) >= 8 else row + [''] * (8 - len(row))
            for row in ineligible_data
        ]

        if eligible_data:  # Only add page break if there was an eligible section
            elements.append(PageBreak())

        table = create_table(
            doc,
            data=processed_ineligible_data,
            header=header_row,
            table_type="INELIGIBLE",
            count=len(processed_ineligible_data)
        )
        elements.append(table)

    # Add BTZ validation section if data provided
    if btz_data:
        if eligible_data or ineligible_data:  # Add page break if there was prior content
            elements.append(PageBreak())
        table = create_table(
            doc,
            data=btz_data,
            header=header_row,
            table_type="VERIFY MEMBER ELIGIBILITY",
            count=len(btz_data)
        )
        elements.append(table)

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