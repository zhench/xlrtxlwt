import xlrd
import os
import shutil

rootdir = '..\data\下仓'
ent_path = '..\data\\下仓镇0529.xlsx'
dstdir = '..\data\下仓镇筛选'
dst_xls_dir = os.path.join(dstdir, 'excel')
if os.path.exists(dstdir):
    print('删除' + dstdir)
    shutil.rmtree(dstdir)
os.mkdir(dstdir)
if os.path.exists(dst_xls_dir):
    print('删除' + dst_xls_dir)
    shutil.rmtree(dst_xls_dir)
os.mkdir(dst_xls_dir)
ent_boot = xlrd.open_workbook(ent_path)
ent_sheet = ent_boot.sheets()[0]

nrows = ent_sheet.nrows
ent_names = ent_sheet.col_values(30)

print('行数为' + str(nrows))
listdir = os.listdir(rootdir)

print('开始复制...')
index = 0
for i in range(0, len(listdir)):
    path = os.path.join(rootdir, listdir[i])
    if not os.path.isfile(path) and listdir[i] in ent_names:
        ent_copy_path = os.path.join(dstdir, listdir[i])
        shutil.copytree(path, ent_copy_path)
        index += 1
        print(str(index) + listdir[i])
        list_copy_ent_sub = os.listdir(ent_copy_path)
        for j in range(0, len(list_copy_ent_sub)):
            if 'xlsx' in list_copy_ent_sub[j] or 'xls' in list_copy_ent_sub[j]:
                xls_path = os.path.join(ent_copy_path, list_copy_ent_sub[j])
                shutil.copyfile(xls_path, os.path.join(dst_xls_dir, list_copy_ent_sub[j]))
print('结束复制...')
