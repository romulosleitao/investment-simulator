from openpyxl.styles import Font, PatternFill


class StylesManager:
    def __init__(self):
        self.header_fill = PatternFill(
            start_color="1F4E78", end_color="1F4E78", fill_type="solid"
        )
        self.header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

        self.accent_fill = PatternFill(
            start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
        )
        self.accent_font = Font(name="Calibri", size=11, bold=True, color="000000")

        self.bold_font = Font(name="Calibri", size=11, bold=True)
        self.regular_font = Font(name="Calibri", size=11)
        self.title_font = Font(name="Calibri", size=16, bold=True, color="1F4E78")
