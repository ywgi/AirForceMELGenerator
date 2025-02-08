def add_page_elements(self, canvas, doc):
    """Add header and footer to each page"""
    canvas.saveState()
    self.add_header(canvas, doc)
    # add header border bottom
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