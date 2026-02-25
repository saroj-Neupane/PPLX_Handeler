"""GUI constants and configuration."""

# White & Purple theme
THEME = {
    "bg_dark": "#f5f3ff",
    "bg_panel": "#ffffff",
    "bg_card": "#f8f7fc",
    "purple": "#6b4c9a",
    "purple_light": "#9b7bb8",
    "purple_dim": "#b8a9d4",
    "text": "#2d2a3e",
    "text_muted": "#6b6b7b",
    "accent": "#6b4c9a",
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
