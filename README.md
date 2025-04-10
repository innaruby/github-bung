 def is_yellow_or_green(cell):
    fill = cell.fill
    if fill is None or fill.fill_type is None:
        return False

    color = fill.start_color

    if color.type == "rgb" and color.rgb:
        rgb = color.rgb.upper()
        return rgb.startswith("FFFF00") or rgb.startswith("FF00FF00") or rgb.startswith("FFFFFF99") or rgb.startswith("FFFFE135")
    
    # You can also match against indexed values (yellow ~ 6, light yellow ~ 22 etc.)
    if color.type == "indexed" and color.indexed in [6, 22, 44, 3]:  # common yellow and green indexes
        return True

    return False
