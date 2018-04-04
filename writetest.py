# -*- coding:utf-8 -*-
from XlsxTool import XlsxHandler
from openpyxl.styles import Font, Border, Side, Fill, colors, PatternFill


def test():
    x = XlsxHandler()
    sheet = x.create_sheet("表1")

    sheet.set_column_values([3, "黑"], 1, start_row_index=5,
                        style={"color": colors.RED, "background": "7CCD7C", "bold": True})
    x.change_sheet("Sheet")
    x.set_value(1, 2, "表")
    print x.sheet_names()
    x.save("写入测试.xlsx")


if __name__ == "__main__":
    test()