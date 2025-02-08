from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image, Spacer
from reportlab.lib.units import inch


def create_military_roster(output_filename):
    doc = SimpleDocTemplate(
        output_filename,
        pagesize=letter,
        rightMargin=0.5 * inch,
        leftMargin=0.5 * inch,
        topMargin=0.5 * inch,
        bottomMargin=0.5 * inch
    )

    # Container for elements
    elements = []

    # Styles
    styles = getSampleStyleSheet()
    header_style = ParagraphStyle(
        'CustomHeader',
        parent=styles['Normal'],
        fontSize=12,
        spaceAfter=20
    )

    # Header data
    header_data = [
        ['PAS: RF17FFSJ', '', 'FOR OFFICIAL USE ONLY'],
        ['SRID : 1M-ICA', '', ''],
        ['SRID Name: BG TREVINO, ALICE W', '', ''],
        ['PAS MPS: RF', '', '']
    ]

    # Create header table
    header_table = Table(header_data, colWidths=[4 * inch, 2 * inch, 3 * inch])
    header_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
    ]))

    # Main table headers
    column_headers = ['NAME', 'GRADE', 'LAST 4', 'SCOD AD Unit', 'DAFSC', 'PROM ELIG STATUS']

    # Sample data row
    data = [
        ['TIG TIS: ELIGIBLE', '', '', '', '', 'ASSIGNED TO MPS: 1'],
        ['JOHNSON MICHAEL G JR', '(E8) SMS', 'XXX-XX-0314', 'AF INSTAL CONTRACT CENTER (OL AFE)', 'D6C091',
         'ELIG - ELIGIBLE FOR SELECTION/PROMOTION'],
    ]

    # Combine headers and data
    table_data = [column_headers] + data

    # Create main table
    main_table = Table(table_data, repeatRows=1)
    main_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#000080')),  # Navy blue header
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#E6E6FA')),  # Light blue for TIG TIS row
    ]))

    # Footer text
    footer_text = Paragraph(
        "The information herein is FOR OFFICIAL USE ONLY (FOUO) information which must be protected under the Freedom of Information Act (5 U.S.C. 552) and/or the Privacy Act of 1974 (5 U.S.C. 552a). Unauthorized disclosure or misuse of this PERSONAL INFORMATION may result in disciplinary action, criminal and/or civil penalties.",
        styles['Normal']
    )

    # Add elements to document
    elements.append(header_table)
    elements.append(Spacer(1, 20))
    elements.append(main_table)
    elements.append(Spacer(1, 20))
    elements.append(footer_text)

    # Build document
    doc.build(elements)


create_military_roster('military_roster.pdf')