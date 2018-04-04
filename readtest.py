# -*- coding:utf-8 -*-
from XlsxTool import XlsxHandler
from openpyxl.styles import Font, Border, Side, Fill, colors, PatternFill


def test():
    '''
    read excel file, sheetname specify which sheet to use,
    if it is None, then the first sheet will be used
    '''
    x = XlsxHandler(open_file_name="readtest.xlsx", sheetname="Sheet2")
    '''
    we can get single cell value specifying row and column number
    we can also get a number of values using row_values and column_values
    '''
    print x.value(1, 2)
    print x.row_values(row=1, start_column_num=1)
    print x.column_values(column=3, start_row_num=1)
    '''
    we can set also set value with some style
    '''

    style_dict = {"color": colors.RED, "background": "7CCD7C", "bold": True}
    x.set_value(4, 1, "测试", style=style_dict)
    x.set_style(1, 2, style_dict)
    x.set_row_values([1, 4, "彩色"], 4, 1, style=style_dict)
    x.set_column_values(["a"], 6, 1, style=style_dict)
    x.save("new.xlsx")
    

if __name__ == "__main__":
    test()

