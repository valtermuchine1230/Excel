"""
VERISCOPE EDGE - Style Definitions Module
Centralized style presets for consistent visual design across all sheets.
"""

from openpyxl.styles import (
    PatternFill, Font, Border, Side, Alignment, Protection, numbers
)

# ============================================================================
# COLOR PALETTE (Exact HEX values per spec)
# ============================================================================

COLORS = {
    'bg_dark': '0B1120',           # Dark base
    'surface': '111827',            # Card background
    'surface_light': '1A2333',      # Secondary card
    'accent_gold': 'D4AF37',        # Primary accent
    'accent_teal': '22D3A6',        # Positive/bullish
    'accent_red': 'E5484D',         # Negative/bearish
    'accent_amber': 'F5A524',       # Warning/caution
    'text_light': 'F8FAFC',         # Text on dark
    'text_dark': '1E293B',          # Text on light
    'bg_light': 'F4F6FA',           # Light background
    'border_grid': '2A3550',        # Border replacement
    'white': 'FFFFFF',
    'black': '000000',
}

# ============================================================================
# TYPOGRAPHY DEFAULTS
# ============================================================================

FONTS = {
    'header_bold': 'Montserrat',
    'body': 'Calibri',
    'fallback': 'Segoe UI',
}

# ============================================================================
# PREDEFINED STYLE PRESETS (reusable throughout the workbook)
# ============================================================================

class StylePresets:
    """Centralized style definitions for consistent application."""
    
    # Title Band Style (dark background, gold text, bold, large)
    @staticmethod
    def title_band():
        return {
            'fill': PatternFill(start_color=COLORS['bg_dark'], end_color=COLORS['bg_dark'], fill_type='solid'),
            'font': Font(name=FONTS['header_bold'], size=16, bold=True, color=COLORS['accent_gold']),
            'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True),
            'border': Border(
                bottom=Side(style='thin', color=COLORS['accent_gold']),
                top=Side(style='thin', color=COLORS['accent_gold']),
                left=Side(style='thin', color=COLORS['accent_gold']),
                right=Side(style='thin', color=COLORS['accent_gold']),
            ),
        }
    
    # KPI Card Label (small gray text)
    @staticmethod
    def kpi_label():
        return {
            'font': Font(name=FONTS['body'], size=9, color=COLORS['text_dark'], bold=False),
            'alignment': Alignment(horizontal='center', vertical='center'),
            'fill': PatternFill(start_color=COLORS['surface_light'], end_color=COLORS['surface_light'], fill_type='solid'),
        }
    
    # KPI Card Value (large bold number, colored)
    @staticmethod
    def kpi_value(color=None):
        if color is None:
            color = COLORS['accent_gold']
        return {
            'font': Font(name=FONTS['body'], size=22, bold=True, color=color),
            'alignment': Alignment(horizontal='center', vertical='center'),
            'fill': PatternFill(start_color=COLORS['surface'], end_color=COLORS['surface'], fill_type='solid'),
            'border': Border(
                top=Side(style='thin', color=COLORS['border_grid']),
                bottom=Side(style='thin', color=COLORS['border_grid']),
                left=Side(style='thin', color=COLORS['border_grid']),
                right=Side(style='thin', color=COLORS['border_grid']),
            ),
        }
    
    # Section Header (dark background, white text, bold)
    @staticmethod
    def section_header():
        return {
            'fill': PatternFill(start_color=COLORS['surface'], end_color=COLORS['surface'], fill_type='solid'),
            'font': Font(name=FONTS['body'], size=11, bold=True, color=COLORS['text_light']),
            'alignment': Alignment(horizontal='left', vertical='center'),
            'border': Border(
                bottom=Side(style='thin', color=COLORS['border_grid']),
            ),
        }
    
    # Data Table Header (light background, dark text, bold)
    @staticmethod
    def table_header():
        return {
            'fill': PatternFill(start_color=COLORS['bg_light'], end_color=COLORS['bg_light'], fill_type='solid'),
            'font': Font(name=FONTS['body'], size=10, bold=True, color=COLORS['text_dark']),
            'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True),
            'border': Border(
                bottom=Side(style='thin', color=COLORS['border_grid']),
                top=Side(style='thin', color=COLORS['border_grid']),
                left=Side(style='thin', color=COLORS['border_grid']),
                right=Side(style='thin', color=COLORS['border_grid']),
            ),
        }
    
    # Input Cell (light background, unlocked)
    @staticmethod
    def input_cell():
        return {
            'fill': PatternFill(start_color=COLORS['bg_light'], end_color=COLORS['bg_light'], fill_type='solid'),
            'font': Font(name=FONTS['body'], size=10, color=COLORS['text_dark']),
            'alignment': Alignment(horizontal='left', vertical='center'),
            'border': Border(
                bottom=Side(style='thin', color=COLORS['border_grid']),
                top=Side(style='thin', color=COLORS['border_grid']),
                left=Side(style='thin', color=COLORS['border_grid']),
                right=Side(style='thin', color=COLORS['border_grid']),
            ),
            'protection': Protection(locked=False),
        }
    
    # Locked Cell (darker background, locked, formula output)
    @staticmethod
    def locked_cell():
        return {
            'fill': PatternFill(start_color=COLORS['surface_light'], end_color=COLORS['surface_light'], fill_type='solid'),
            'font': Font(name=FONTS['body'], size=10, color=COLORS['text_light']),
            'alignment': Alignment(horizontal='center', vertical='center'),
            'border': Border(
                bottom=Side(style='thin', color=COLORS['border_grid']),
                top=Side(style='thin', color=COLORS['border_grid']),
                left=Side(style='thin', color=COLORS['border_grid']),
                right=Side(style='thin', color=COLORS['border_grid']),
            ),
            'protection': Protection(locked=True),
        }
    
    # Positive Value (green/teal)
    @staticmethod
    def positive_value():
        return {
            'fill': PatternFill(start_color=COLORS['surface_light'], end_color=COLORS['surface_light'], fill_type='solid'),
            'font': Font(name=FONTS['body'], size=10, bold=True, color=COLORS['accent_teal']),
            'alignment': Alignment(horizontal='right', vertical='center'),
            'protection': Protection(locked=True),
        }
    
    # Negative Value (red)
    @staticmethod
    def negative_value():
        return {
            'fill': PatternFill(start_color=COLORS['surface_light'], end_color=COLORS['surface_light'], fill_type='solid'),
            'font': Font(name=FONTS['body'], size=10, bold=True, color=COLORS['accent_red']),
            'alignment': Alignment(horizontal='right', vertical='center'),
            'protection': Protection(locked=True),
        }
    
    # Warning/Alert (amber background)
    @staticmethod
    def warning():
        return {
            'fill': PatternFill(start_color=COLORS['accent_amber'], end_color=COLORS['accent_amber'], fill_type='solid'),
            'font': Font(name=FONTS['body'], size=9, color=COLORS['text_dark'], bold=True),
            'alignment': Alignment(horizontal='left', vertical='center', wrap_text=True),
        }
    
    # Badge Style (small colored background, centered)
    @staticmethod
    def badge(bg_color):
        return {
            'fill': PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid'),
            'font': Font(name=FONTS['body'], size=10, bold=True, color=COLORS['white']),
            'alignment': Alignment(horizontal='center', vertical='center'),
            'protection': Protection(locked=False),
        }


def apply_style(cell, style_dict):
    """Apply a style dictionary to a cell."""
    if 'fill' in style_dict:
        cell.fill = style_dict['fill']
    if 'font' in style_dict:
        cell.font = style_dict['font']
    if 'alignment' in style_dict:
        cell.alignment = style_dict['alignment']
    if 'border' in style_dict:
        cell.border = style_dict['border']
    if 'number_format' in style_dict:
        cell.number_format = style_dict['number_format']
    if 'protection' in style_dict:
        cell.protection = style_dict['protection']
