# python 办公自动化
## 安装python。。。。。。
## 安装pycharm
## python基础
### 数字和字符串可以相乘。  
数个字符串可以相加，不能相减  
字符串替换：字符串.replace('想替换','被替换')
字符串格式化：  
1.‘含有%s的字符串’%‘要插入的字符串’
2.‘含有{}的字符串'.format('要插入的字符串')  
### 切片  
print(a[0:4])
取后三位：[-3:]  
### 数据结构：  
1.列表：list.append();  list.remove();    
2.字典:	  
3.元组：  
4.集合：  
### 比较运算符：==/!=/...
### 循环
### 函数，调用第三方库:type();len();round(1.5456444,3);input();def.....;
##  办公自动化：
### xlsx文件：工作簿，工作表(sheet)，单元格  
### xlrd  xlwt  
  
    import xlrd
    xlsx = xlrd.open_workbook('文件目录+文件名')    
    table = xlsx.sheet_by_index()
    #table = xlsx.sheet_by_name('sheet名')  
    print(table.cell_value(1,4))  
    print(table.cell(1,4).value)
    print(table.row(1)[4].value)  
	
    #import xlwt
    #new_workbook = xlwt.Workbook()  
    #worksheet = new_workbook.add_sheet('sheet_test')  
    #worksheet.write(0,0,'test')  
    #nwe_workbook.save('路径+文件名')  
  
  
### xlutils   
  
    from xlutils.copy import copy
    import xlrd
    import xlwt
    
    tem_excel = xlrd.open_workbook('/Users/tnjmytuu/Documents/test.xls',formatting_info=True)
    tem_sheet = tem_excel.sheet_by_name('sheet_test')
    
    new_excel = copy(tem_excel)
    new_sheet = new_excel.get_sheet(0)
    
    #设置格式
    style = xlwt.XFStyle()
    
    #设置字体
    font = xlwt.Font()
    font.name = '微软雅黑'
    font.bold = True
    font.height = 360  #  字号22乘以80
    style.font = font
    
    borders = xlwt.Borders()
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    style.borders = borders
    
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER
    style.alignment = alignment
    
    new_sheet.write(2,1,12.style)
    
    new_excel.save('/Users/tnjmytuu/Documents/test_1.xls') 
    
