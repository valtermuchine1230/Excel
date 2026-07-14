"""
VERISCOPE EDGE - Styling Module
Centralized style presets for consistent branding across all sheets.
"""

from openpyxl.styles import PatternFill, Font, Border, Side, Alignment, Protection, numbers

# ============================================================================
# COLOR PALETTE (4-Color Rule - Non-Negotiable)
# ============================================================================

COLORS = {
    'bg_dark': '0B1120',           # Dark base (background/headers)
    'text_light': 'F8FAFC',         # White/near-white (surface, text on dark)
    'accent_gold': 'D4AF37',        # Gold (primary accent)
    'accent_green': '22D3A6',       # Profit green
    'accent_red': 'E5484D',         # Loss red
    'text_dark': '1A1A1A',          # Dark text
    'bg_light': 'F5F5F5',           # Light background
}

# ============================================================================
# STYLE PRESETS (Reusable across all sheets)
# ============================================================================

class StylePresets:
    """Collection of reusable style definitions."""
    
    @staticmethod
    def title_band():
        """Style for main sheet title bands."""
        return {
            'fill': PatternFill(start_color=COLORS['bg_dark'], end_color=COLORS['bg_dark'], fill_type='solid'),
            'font': Font(name='Montserrat', size=18, bold=True, color=COLORS['accent_gold']),
            'alignment': Alignment(horizontal='center', vertical='center'),
            'border': Border(
                bottom=Side(style='thin', color=COLORS['accent_gold']),
            ),
        }
    
    @staticmethod
    def section_header():
        """Style for section headers within sheets."""
        return {
            'fill': PatternFill(start_color=COLORS['accent_gold'], end_color=COLORS['accent_gold'], fill_type='solid'),
            'font': Font(size=12, bold=True, color=COLORS['bg_dark']),
            'alignment': Alignment(horizontal='left', vertical='center'),
        }
    
    @staticmethod
    def table_header():
        """Style for table column headers."""
        return {
            'fill': PatternFill(start_color=COLORS['bg_dark'], end_color=COLORS['bg_dark'], fill_type='solid'),
            'font': Font(size=10, bold=True, color=COLORS['text_light']),
            'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True),
            'border': Border(
                left=Side(style='thin', color=COLORS['accent_gold']),
                right=Side(style='thin', color=COLORS['accent_gold']),
                top=Side(style='thin', color=COLORS['accent_gold']),
                bottom=Side(style='thin', color=COLORS['accent_gold']),
            ),
        }
    
    @staticmethod
    def input_cell():
        """Style for user input cells (unlocked)."""
        return {
            'fill': PatternFill(start_color='E8F4F8', end_color='E8F4F8', fill_type='solid'),
            'font': Font(size=10, color=COLORS['text_dark']),
            'alignment': Alignment(horizontal='left', vertical='center'),
            'border': Border(
                left=Side(style='thin', color='CCCCCC'),
                right=Side(style='thin', color='CCCCCC'),
                top=Side(style='thin', color='CCCCCC'),
                bottom=Side(style='thin', color='CCCCCC'),
            ),
            'protection': Protection(locked=False),
        }
    
    @staticmethod
    def locked_cell():
        """Style for formula/locked cells."""
        return {
            'fill': PatternFill(start_color='F0F0F0', end_color='F0F0F0', fill_type='solid'),
            'font': Font(size=10, color=COLORS['text_dark']),
            'alignment': Alignment(horizontal='right', vertical='center'),
            'border': Border(
                left=Side(style='thin', color='CCCCCC'),
                right=Side(style='thin', color='CCCCCC'),
                top=Side(style='thin', color='CCCCCC'),
                bottom=Side(style='thin', color='CCCCCC'),
            ),
            'protection': Protection(locked=True),
        }
    
    @staticmethod
    def kpi_label():
        """Style for KPI card labels."""
        return {
            'fill': PatternFill(start_color=COLORS['bg_light'], end_color=COLORS['bg_light'], fill_type='solid'),
            'font': Font(size=9, bold=False, color=COLORS['text_dark']),
            'alignment': Alignment(horizontal='center', vertical='bottom'),
            'protection': Protection(locked=True),
        }
    
    @staticmethod
    def kpi_value(color=None):
        """Style for KPI card values."""
        if color is None:
            color = COLORS['accent_gold']
        return {
            'fill': PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid'),
            'font': Font(size=20, bold=True, color=color),
            'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True),
            'border': Border(
                left=Side(style='medium', color=color),
                right=Side(style='medium', color=color),
                top=Side(style='medium', color=color),
                bottom=Side(style='medium', color=color),
            ),
            'protection': Protection(locked=True),
            'number_format': '#,##0.00',
        }
    
    @staticmethod
    def warning():
        """Style for warning/disclaimer text."""
        return {
            'fill': PatternFill(start_color='FFF3CD', end_color='FFF3CD', fill_type='solid'),
            'font': Font(size=9, italic=True, color=COLORS['text_dark']),
            'alignment': Alignment(horizontal='left', vertical='top', wrap_text=True),
            'border': Border(
                left=Side(style='thin', color=COLORS['accent_red']),
                right=Side(style='thin', color=COLORS['accent_red']),
                top=Side(style='thin', color=COLORS['accent_red']),
                bottom=Side(style='thin', color=COLORS['accent_red']),
            ),
        }
    
    @staticmethod
    def percent_format():
        """Percentage number format."""
        return '0.00%'
    
    @staticmethod
    def currency_format():
        """Currency number format."""
        return '"$"#,##0.00'


def apply_style(cell, style_dict):
    """Apply a style dictionary to a cell."""
    if isinstance(style_dict, dict):
        for key, value in style_dict.items():
            if key == 'fill':
                cell.fill = value
            elif key == 'font':
                cell.font = value
            elif key == 'alignment':
                cell.alignment = value
            elif key == 'border':
                cell.border = value
            elif key == 'protection':
                cell.protection = value
            elif key == 'number_format':
                cell.number_format = value
