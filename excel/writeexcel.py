def set_style(fontname, height, color_index, bold=False):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = fontname
    font.bold = bold
    font.colour_index = color_index
    font.height = height
    style.font = font
    return style


def write_to_exl(sheet, row_value, rowindex, ncols, color_index):
    style = set_style('宋体', 220, color_index)
    for i in range(ncols):
        sheet.write(rowindex, i, row_value[i], style)
