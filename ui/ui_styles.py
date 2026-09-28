import tkinter as tk

# Centralized UI styling constants and small helpers for the app

COLOR_PRIMARY = "#000000"
COLOR_SECONDARY = "#3498db"
COLOR_ACCENT = "#e74c3c"
COLOR_SUCCESS = "#27ae60"
COLOR_BG = "#f0f0f0"
COLOR_LIGHT = "#ecf0f1"
COLOR_TEXT = "#2c3e50"
COLOR_TEXT_LIGHT = "#7f8c8d"


# Colors
ROW_BG_1 = "#eeeeee"
ROW_BG_2 = "#e0e0e0"
CELL_BORDER = "#b3b3b3"

DARK_GREY = "#727272"
LIGHT_GREY = "#eeeeee"

PRIMARY_BG = "#4d4d4d"
SECONDARY_BG = "#b3b3b3"

LIKERT_SELECTED_BG = "#4d4d4d"
LIKERT_DEFAULT_BG = "#d9d9d9"

# Fonts
TEXT_FONT = ("Segoe UI", 11)
FONTS = {
    "text": TEXT_FONT,
    "small": ("Segoe UI", 10),
    "small_normal": ("Segoe UI", 10, "normal"),
    "small_bold": ("Segoe UI", 10, "bold"),
    "medium": ("Segoe UI", 11),
    "bold": ("Segoe UI", 11, "bold"),
    "medium_bold": ("Segoe UI", 11, "bold"),
    "title": ("Segoe UI", 14, "bold"),
}


def style_label(widget: tk.Widget, variant: str = "text", bg: str | None = None, fg: str | None = None, **kwargs) -> None:
    """Apply common label styles."""
    font = FONTS.get(variant, TEXT_FONT)
    widget.config(font=font)
    if bg is not None:
        widget.config(bg=bg)
    if fg is not None:
        widget.config(fg=fg)
    if kwargs:
        widget.config(**kwargs)


def style_button(widget: tk.Widget, variant: str = "primary") -> None:
    """Apply common button styles (works for `tk.Button`)."""
    if variant == "primary":
        widget.config(bg=PRIMARY_BG, fg="white", font=FONTS["bold"], padx=20, pady=5)
    else:
        widget.config(bg=SECONDARY_BG, fg="white", font=FONTS["bold"], padx=20, pady=5)


def make_likert_label(parent: tk.Widget, text: str, width: int = 3, font_key: str = "small") -> tk.Label:
    """Create a likert-style label used as a clickable option."""
    lbl = tk.Label(
        parent,
        text=text,
        width=width,
        bd=1,
        relief="raised",
        bg=LIKERT_DEFAULT_BG,
        fg="black",
        font=FONTS.get(font_key, TEXT_FONT),
    )
    return lbl


# Backwards-compatible style mapping used in older phase files (S = {...})
S = {
    "bg": "#ffffff",
    "dark": DARK_GREY,
    "header": "#807C7D",
    "odd": ROW_BG_1,
    "even": ROW_BG_2,
    "yellow": "#000000", #dit is elke kleur geel in de hele site eerdere naam was *yellow*
    "btn": LIKERT_DEFAULT_BG,
    "btn_on": LIKERT_SELECTED_BG,
    "f_title": FONTS.get("title"),
    "f_sub": FONTS.get("small"),
    "f": FONTS.get("small"),
    "f_b": FONTS.get("small_bold"),
    "f_small": ("Segoe UI", 9),
}

"""
Centralized Tkinter UI constants for colors, fonts, spacing, and layout styling.
"""

# --- Colors ---
# -- Base colors --
COLOR_BLACK = "#000000"
COLOR_WHITE = "#FFFFFF"

# -- Page colors --
PAGE_BACKGROUND_COLOR = "#FFFFFF"
PANEL_BACKGROUND_COLOR = "#F5F5F5"

# -- Sidebar colors --
SIDEBAR_BACKGROUND_COLOR = "#000000"

# -- Table colors --
TABLE_HEADER_BACKGROUND_COLOR = "#363636"
TABLE_ROW_PRIMARY_BACKGROUND_COLOR = "#EEEEEE"
TABLE_ROW_SECONDARY_BACKGROUND_COLOR = "#E0E0E0"

# -- Control colors --
# - Primary buttons -
CONTROL_PRIMARY_BACKGROUND_COLOR = "#D9D9D9"
CONTROL_PRIMARY_BACKGROUND_ACTIVE_COLOR = "#C0C0C0"

# - Secondary buttons -
CONTROL_SECONDARY_BACKGROUND_COLOR = "#363636"
CONTROL_SECONDARY_BACKGROUND_ACTIVE_COLOR = "#F0F0F0"

# -- Text colors ---
TEXT_PRIMARY_COLOR = "#000000"
TEXT_SECONDARY_COLOR = "#FFFFFF"

# -- Accent colors --
ACCENT_PRIMARY_COLOR = "#000000"
YELLOW_ACCENT = "#000000"

# -- Card/Background colors --
CARD_LIGHT_BG = "#F5F5F5"
WHITE_BG = "#FFFFFF"

# --- Fonts ---
TITLE_FONT = ("Segoe UI", 14, "bold")
SUBTITLE_FONT = ("Segoe UI", 10, "normal")
TABLE_HEADER_FONT = ("Segoe UI", 10, "bold")
TEXT_FONT = ("Segoe UI", 10, "normal")
NUMBER_FONT = ("Segoe UI", 10, "bold")
BUTTON_FONT = ("Segoe UI", 10, "normal")

# --- Layout ---
# -- Padding --
# - Title padding -
TITLE_PADX = 20
TITLE_PADY_TOP = 15
TITLE_PADY_BOTTOM = 5

# - Subtitle padding -
SUBTITLE_PADX = 20
SUBTITLE_PADY_TOP = 0
SUBTITLE_PADY_BOTTOM = 15

# - Section padding -
TABLE_HEADER_PADX = 10
TABLE_HEADER_PADY = 6
TABLE_PADX = 20
TABLE_PADY = (0, 10)

# - Statement padding -
STATEMENT_PADX = 10
STATEMENT_PADY = 8

# - Row padding -
ROW_PADY = 1
ROW_INNER_PADX = 10
ROW_INNER_PADY = 8

# - Button padding -
BUTTON_PADX = 20
BUTTON_PADY = 5
BUTTON_CONTAINER_PADX = 30
BUTTON_FRAME_PADY = 20

# -- Question row --
NUMBER_WIDTH = 4
STATEMENT_WRAP = 650
LIKERT_BUTTON_WIDTH = 3
LIKERT_BUTTON_PADX = 4

# -- Separator between rows
SEPARATOR_HEIGHT = 1
SEPARATOR_COLOR = COLOR_WHITE
