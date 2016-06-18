# -*- coding: utf-8 -*-
import xlrd
from xlwt import *
from xlutils.copy import copy

'''
Excel 增删改查 
应用3个库文件
具体参照： http://www.cnblogs.com/BeginMan/p/3657805.html
'''


def cell_modify(sheet,row,col,ctype,value,xf=0):
    '''
    @类型 0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
    @ctype = 1 
    @value = '单元格的值'
    @xf = 0 # 扩展的格式化 
    '''   
    sheet.put_cell(row, col, ctype, value, xf)

def cell_read(sheet,row,col):
    '''
    @cell_A1 = table.cell(0,0).value
    @cell_C4 = table.cell(2,3).value
    
    '''
    return  sheet.cell(row,col).value

def workbook_rb(path):
    '''
    @param path: r'filepath'
    @return: workbook fd   
    '''
    try:
        print '[workbook_rb] opening file ' + path
        rb = xlrd.open_workbook(path,encoding_override='utf-8')
        return rb
    except:
        print '[workbook_rb] Error occur!'
        raise
    
def excel_close():
    return
def sheet_rows_num(sheet):
    return sheet.nrows
    
def sheet_cols_num(sheet):
    return sheet.ncols


def workbook_copy(rb):
    '''
    @return: workbook which is write enabled
    '''
    try:
        wb = copy(rb)
        return wb
    except:
        print '[workbook_copy] Error occur!'
        raise
    
def workbook_cell_write(wb,row,col,value = ''):
    try:
        if value == '' :
            return
        ws = wb.get_sheet(0)
        ws.write(row,col,value)   
    except:
        print '[cell_write] Error occur!'
        raise
def workbook_save(wb,path):
    try:
        wb.save(path)
    except:
        print '[workbook_save] Error occur!'
        raise


def read_vin_list(path):
    '''
    B: 1 = VIN
    Q:16 = 运输公司
    W:22 = TSS省
    X:23 = TSS市
    Y:24 = 实际到达时间
    @return: list of VIN where 运输公司='重庆博宇'
    '''
    lst = []
    try:
        rb = workbook_rb(path)
        rs = rb.sheet_by_index(0)
        for row in range(rs.nrows):
            if rs.cell(row,16).value == '重庆博宇':
                lst.append(str(rs.cell(row,1).value))
        return lst   
                
    except:
        print '[read_vin_list] Error occur!'
        raise
    
def excel_update(from_path,to_path,result_list):
    '''
    @summary: update excel 
    @param path: excel file path
    @param result_lsit:query available record   
    1:VIN22:当前位置省份23:当前位置城市 24:实际到达时间
    result cols info:
        0: 序号
        1: 验证结果
        2: 校验未通过原因
        3: 标准距离
        4: 扫描距离
        6: 扫描操作时间  str[0:18]
        9: VIN码
        10: 操作名称
        24: TSS省
        25: TSS市
        44: 运输公司 
    '''
    try:
        rb = workbook_rb(from_path)
        rs = rb.sheet_by_index(0)
        wb = workbook_copy(rb)
        ws = wb.get_sheet(0)
        
        font0 = Font()
        #font0.name = 'Times New Roman'
        #font0.struck_out = True
        font0.bold = True
        style0 = XFStyle()
        style0.font = font0
        
        font1 = Font()
        font1.bold = True
        style1 = XFStyle()
        #style1.font = font1
        style1.num_format_str = 'YYYY-MM-DD hh:mm:ss'
        
        #Excel Header Row Style
        style2 = XFStyle()
        font2 = Font()
        font2.bold = True
        style2.font = font2
       
        
        #Excel Header CellStyle
        style3 = XFStyle()
        font3 = Font()
        font3.bold = True
        font3.colour_index = 2 #red
        style3.font = font3
        print '[update_record] handle excel ' + from_path + ' start ...'
        #Update Excel Header Style
        for c in range(rs.ncols):
            if c in (1,2,7,19,20,21,22,23): 
                ws.write(0,c,rs.cell(0,c).value,style3)
            else:
                ws.write(0,c,rs.cell(0,c).value,style2)
        for lst in result_list:
            for r in range(rs.nrows):
                if lst[9] == rs.cell(r,1).value:
                    
                    ws.write(r,22,lst[24],style0)
                    print '[excel_update] update province from ' + rs.cell(r,22).value + ' to '+ lst[24]  
                    
                    if str(lst[10])[0:2] == '05':
                        if str(lst[1]) == '校验通过':
                            ws.write(r,23,lst[25],style0) #当前市-已交车
                            ws.write(r,24,lst[6],style1) #实际时间-时间
                            print '[excel_update] 05已交车校验通过: update city from ' + rs.cell(r,23).value + ' to '+ lst[25]
                            print '[excel_update] 05已交车校验通过: update datetime from ' + str(rs.cell(r,24).value) + ' to '+ lst[6]
                        else:
                            ws.write(r,23,lst[25]+'-'+lst[1]+'-'+lst[2],style0)
                            ws.write(r,24,lst[6],style1) #实际时间-时间
                            print '[excel_update] 05已交车校验未通过: update city from ' + rs.cell(r,23).value + ' to '+ lst[25]+'-'+lst[2]
                            print '[excel_update] 05已交车校验未通过: update datetime from ' + str(rs.cell(r,24).value) + ' to '+ lst[6]
                    else:
                        ws.write(r,23,lst[25],style0)#当前市-已交车
                        print '[excel_update] 03在途：update city from ' + rs.cell(r,23).value + ' to '+ lst[25]
        wb.save(to_path)                    
        print '[update_record] handle excel finish.'
    except:
        print '[update_record] Error occur!'
        raise
    
if __name__=="__main__":
    '''
    @keyword testcase: 
        获取整行和整列的值（数组) 　　
         table.row_values(i)
         table.col_values(i)
        获取行数和列数
        nrows = table.nrows
        ncols = table.ncols
        循环行列表数据
        for i in range(nrows ):
      print table.row_values(i)
    '''
    rb = workbook_rb(r'D:\20160606.xlsx');
    fd = rb.sheet_by_index(0)
    print sheet_rows_num(fd);
    print sheet_cols_num(fd);
    print cell_read(fd,0,0);
    print cell_read(fd,1,1);
    print fd.cell(1,1)
    cell_modify(fd,1,1,1,'testtest');
    print cell_read(fd, 1, 1)
    
    #####read_vin_list test case
    lst = read_vin_list(r'D:\20160606.xlsx')
    print lst;
    print len(lst)

    ###write testcase
    #wb = workbook_copy(rb)
    #workbook_cell_write(wb,1,1,'testwritevalue')
    #workbook_save(wb,r'D:\write.xlsx')
