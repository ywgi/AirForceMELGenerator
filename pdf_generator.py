# from tkinter import Image
#
# import pandas as pd
# from reportlab.lib.pagesizes import landscape, letter
# import pandas as pd
# from reportlab.lib.styles import getSampleStyleSheet
# from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph
# from reportlab.platypus import Image as RLImage
# from reportlab.lib import colors
# import os
#
#
# def pdf_generator(dataframe):
#     pdf = SimpleDocTemplate("dataframe_example.pdf", pagesize=landscape(letter))
#     elements = []
#
#     styles = getSampleStyleSheet()
#     header_style = styles["Title"]
#     subheader_style = styles["Normal"]
#
#     image_path = 'images/air-force-logo.png'
#     logo = RLImage(image_path, 100, 100)
#     # elements.append()
#     logo_table = Table([[logo]], colWidths=[100])
#     logo_table.setStyle(
#         TableStyle([
#             ("ALIGN", (0, 0), (-1, -1), "LEFT"),
#             ("VALIGN", (0, 0), (-1, -1), "TOP"),
#         ])
#     )
#
#     header_text = Paragraph("Testing", header_style)
# # Convert DataFrame to a list of lists
#     table_data = [dataframe.columns.tolist()] + dataframe.values.tolist()
#
#     # Create PDF
#
#     # Create a Table
#     table = Table(table_data)
#     spacer = Spacer(1, 20)
#     # Add Style to the Table
#     style = TableStyle([
#         ('BACKGROUND', (0, 0), (-1, 0), colors.navy),  # Header background color
#         ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),  # Header text color
#         ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Center align text
#         ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Header font
#         ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  # Padding for header
#         ('BACKGROUND', (0, 1), (-1, -1), colors.lightgrey),  # Cell background color
#         ('GRID', (0, 0), (-1, -1), 1, colors.black)  # Grid lines
#     ])
#     table.setStyle(style)
#
#     elements.append(logo_table)
#     elements.append(spacer)
#     elements.append(table)
#
#     # Build the PDF
#     pdf.build(elements)


from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, Image
from reportlab.lib import colors
from reportlab.lib.units import inch
import pandas as pd
import os


def create_header_section(logo_path, pas_info):
    """Create the header section with logo and PAS information"""
    # Create the 'CONTROLLED UNCLASSIFIED INFORMATION' header
    cui_data = [['CUI// CONTROLLED UNCLASSIFIED INFORMATION']]
    cui_table = Table(cui_data, colWidths=[10.2 * inch])  # Adjust width to match full page width
    cui_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
    ]))

    # Header data without cui
    header_data = [
        ['PAS: ' + pas_info.get('pas', ''), '', ''],
        ['SRID: ' + pas_info.get('srid', ''), '', ''],
        ['SRID Name: ' + pas_info.get('srid_name', ''), '', ''],
        ['PAS MPS: ' + pas_info.get('pas_mps', ''), '', '']
    ]

    # Create logo table
    logo = Image(logo_path, width=1 * inch, height=1 * inch)
    logo_table = Table([[logo]], colWidths=[1.2 * inch])
    logo_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))

    # Create header info table
    header_table = Table(header_data, colWidths=[3 * inch, 3 * inch, 3 * inch])
    header_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
    ]))

    # Combine logo and header info
    combined_header = Table([[logo_table, header_table]], colWidths=[1.2 * inch, 9 * inch])
    combined_header.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))

    return [cui_table, Spacer(1, 10), combined_header]


def create_footer():
    """Create the footer with cui text"""
    styles = getSampleStyleSheet()
    footer_style = ParagraphStyle(
        'Footer',
        parent=styles['Normal'],
        fontSize=8,
        leading=10
    )

    footer_text = ("The information herein is FOR OFFICIAL USE ONLY (cui) information which must be protected under "
                   "the Freedom of Information Act (5 U.S.C. 552) and/or the Privacy Act of 1974 (5 U.S.C. 552a). "
                   "Unauthorized disclosure or misuse of this PERSONAL INFORMATION may result in disciplinary action, "
                   "criminal and/or civil penalties.")

    return Paragraph(footer_text, footer_style)


def pdf_generator(dataframe, output_filename="military_roster.pdf", logo_path='images/air-force-logo.png'):
    """Generate a military roster PDF from a DataFrame"""
    # Create PDF document
    doc = SimpleDocTemplate(
        output_filename,
        pagesize=landscape(letter),
        rightMargin=0.5 * inch,
        leftMargin=0.5 * inch,
        topMargin=0.1 * inch,
        bottomMargin=0.1 * inch
    )

    elements = []

    # Sample PAS info - replace with actual data
    pas_info = {
        'pas': 'RF17FFSJ',
        'srid': '1M-ICA',
        'srid_name': 'BG TREVINO, ALICE W',
        'pas_mps': 'RF'
    }

    # Add header elements
    header_elements = create_header_section(logo_path, pas_info)
    elements.extend(header_elements)
    elements.append(Spacer(1, 20))

    # Convert DataFrame to table data
    table_data = [dataframe.columns.tolist()] + dataframe.values.tolist()

    # Create main table
    table = Table(table_data)
    table.setStyle(TableStyle([
        # Header styling
        ('BACKGROUND', (0, 0), (-1, 0), colors.navy),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),

        # Data rows styling
        ('BACKGROUND', (0, 1), (-1, -1), colors.Color(0.95, 0.95, 0.95)),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.Color(0.95, 0.95, 0.95)]),
    ]))

    elements.append(table)
    elements.append(Spacer(1, 20))

    # Add footer
    elements.append(create_footer())

    # Build PDF
    doc.build(elements)

