# -*- coding:utf-8 -*-
import xlrd
import xlwt

noexist_ent_path = "C:\\Users\\Administrator\\PycharmProjects\\execl\\31.xls"
origin_ent_path = "C:\\Users\\Administrator\\PycharmProjects\\execl\\二轮清洗名单.xlsx"
red_row_index = 104

result_file = '二次清洗（删除注销企业）.xls'

noexist_ent_book = xlrd.open_workbook(noexist_ent_path)
origin_ent_book = xlrd.open_workbook(origin_ent_path)
# print(noexist_ent_book)

noexist_sheet = noexist_ent_book.sheets()[2]
origin_sheet = origin_ent_book.sheets()[0]
noexist_ent_names = noexist_sheet.col_values(1)
origin_ent_names = origin_sheet.col_values(3)
# print(noexist_ent_names)

origin_ncols = origin_sheet.ncols
origin_nrows = origin_sheet.nrows
w = xlwt.Workbook()
sheet_has_been_deleted = w.add_sheet('已删除注销企业')
sheet_been_deleted = w.add_sheet('删除的企业')


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


chongfu = []
tt = 0
yy = 0
for indx in range(origin_nrows):
    name = origin_ent_names[indx]
    row_val = origin_sheet.row_values(indx)
    if indx == 0:
        write_to_exl(sheet_has_been_deleted, row_val, indx, origin_ncols, 0)
        tt = +1
        write_to_exl(sheet_been_deleted, row_val, indx, origin_ncols, 0)
        yy = +1
        continue
    if indx <= red_row_index:
        write_to_exl(sheet_has_been_deleted, row_val, indx, origin_ncols, 2)
    else:
        if name in noexist_ent_names:
            # print(str(indx) + name)
            chongfu.append(name)
            write_to_exl(sheet_been_deleted, row_val, indx, origin_ncols, 0)
            continue
        if name == '无':
            write_to_exl(sheet_has_been_deleted, row_val, indx, origin_ncols, 4)
        else:
            write_to_exl(sheet_has_been_deleted, row_val, indx, origin_ncols, 0)

    # origin_sheet.cell(row_index, 26).value
print(type(origin_ent_names))
#origin_ent_names.pop('无')
print(len(origin_ent_names))
print(origin_ent_names.count('无'))
print(len(list(set(origin_ent_names))))
w.save(result_file)
print(len(chongfu))
