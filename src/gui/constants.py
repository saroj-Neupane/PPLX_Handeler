"""GUI constants and configuration."""

# Modern minimal theme - purple accent on clean neutrals
THEME = {
    "bg": "#f5f4f8",              # main window background
    "bg_dark": "#f5f4f8",         # alias for bg (backward compat)
    "bg_panel": "#ffffff",        # card / panel surfaces
    "bg_card": "#f0eef5",         # input fields, listboxes
    "border": "#e2dced",          # subtle borders
    "purple": "#6b4c9a",          # primary accent
    "purple_light": "#8b6fbf",    # hover / active
    "purple_dim": "#d4cceb",      # disabled / placeholder
    "text": "#1e1b2e",            # primary text - near black
    "text_secondary": "#5e5875",  # secondary text
    "text_muted": "#928da5",      # muted / hint text
    "accent": "#6b4c9a",          # alias for purple
}

# Aux fields with auto-fill from Excel
POLE_TAG_BLANK = "NO TAG"
PLACEHOLDER_AUX1 = "(Will auto-fill from Excel)"
PLACEHOLDER_AUX2 = "(Will fill from sheet)"

# (index, config_key, checkbox_label, placeholder, default_checked)
AUX_AUTO_FILL_CONFIG = [
    (0, "auto_fill_aux1", "Auto Fill", PLACEHOLDER_AUX1, False),
    (1, "auto_fill_aux2", "Auto Fill", PLACEHOLDER_AUX2, False),
]
