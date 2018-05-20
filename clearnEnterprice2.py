import xlrd
import xlwt


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


old_ent_path = "excel\\定表二轮清洗 - 0520(1).xlsx"
result_path = 'excel\\result.xls'
add_path = 'excel\\21.xls'

old_ent_book = xlrd.open_workbook(old_ent_path)
old_ent_sheet = old_ent_book.sheets()[0]
add_book = xlrd.open_workbook(add_path)
add_sheet = add_book.sheets()[0]

old_ent_ncols = old_ent_sheet.ncols
old_ent_nrows = old_ent_sheet.nrows
address_index = 9

old_ent_names = old_ent_sheet.col_values(3)
add_names = add_sheet.col_values(3)
add_address = add_sheet.col_values(address_index)

w = xlwt.Workbook()
result_sheet = w.add_sheet('二次清洗底单')

for indx in range(old_ent_nrows):
    name = old_ent_names[indx]
    row_val = old_ent_sheet.row_values(indx)
    if indx == 0:
        write_to_exl(result_sheet, row_val, indx, old_ent_ncols, 0)
        continue
    if name == '无':
        write_to_exl(result_sheet, row_val, indx, old_ent_ncols, 0)
        continue
    if name in add_names:
        add_name_index = add_names.index(name)
        if add_address[add_name_index] == '':
            write_to_exl(result_sheet, row_val, indx, old_ent_ncols, 0)
            continue

        if row_val[address_index] == add_address[add_name_index]:
            write_to_exl(result_sheet, row_val, indx, old_ent_ncols, 0)
        else:
            print(row_val[address_index])
            print(add_address[add_name_index])
            if add_address[add_name_index].count(row_val[address_index]):
                row_val[address_index] = add_address[add_name_index]
                write_to_exl(result_sheet, row_val, indx, old_ent_ncols, 4)
                print('已修改')
            else:
                row_val[address_index] = add_address[add_name_index]
                write_to_exl(result_sheet, row_val, indx, old_ent_ncols, 4)
            print('--------------------------')
    else:
        write_to_exl(result_sheet, row_val, indx, old_ent_ncols, 0)

    w.save(result_path)
