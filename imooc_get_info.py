#coding:utf-8
import xlrd
import xlwt
def floattostr(val):
    if isinstance(val,float):
        val = str(int(val))
    return val
def get_usr_info(path):
    list = []
    list_title = ['usrname','password']
    table = xlrd.open_workbook(path) #打开excel文件
    usr_sheet = table.sheets()[0] #打开第一张表单，即用户信息表
    nrows = usr_sheet.nrows
    ncols = usr_sheet.ncols
    for i in range(1,nrows):
        tmp = [floattostr(val) for val in usr_sheet.row_values(i)]
        list.append(dict(zip(list_title,tmp)))
    return list
if __name__ == '__main__':
    usr_list = get_usr_info(r'F:\PyProjects\imooc\imooc_info.xlsx')
    #print (usr_list)
    for arg in usr_list:
        print(arg)