#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time   : 3/7/2018
# @Author : Sun Zhang
# @Email  : root@meowwoo.com

import xlwings as xw

# excel file path
fn = r"C:\Users\filename.xlsx"
# split number
num = 10
# Maximum column letter (fix bug)
letter = 'W'
# Attention: If you want to define column format, please find the corresponding code (two place).

# The output file is in the output folder.


class ExportNewData(object):

    def __init__(self, fn, num, letter):
        super(ExportNewData, self).__init__()
        self.OldListImport = list()
        self.NewListExport = list()
        self.fn = fn
        self.num = num
        self.letter = letter

    def import_excel(self):
        app = xw.App(visible=True, add_book=False)

        wb1 = app.books.open(self.fn)
        sht1 = wb1.sheets['Sheet1']

        range1 = sht1.range("A1").expand("down")

        self.OldListImport = range1.value
        total = len(self.OldListImport)
        part = (total - 1) // self.num
        count = 1

        raw_name = fn.split('\\')[-1].split('.')[0]

        while count <= num:
            wb2 = app.books.add()
            sht2 = wb2.sheets['Sheet1']
            # Need to define the format of the column
            sht2.range('A1:F' + str(part + 1)).number_format = '@'
            # first part contains the original title.
            if count == 1:
                # Calculate the split range
                field = "$A$1:$" + self.letter + "$" + str(part + 1)
                # fill in data
                sht2.range('A1').options(expand='table').value = sht1.range(field).value
                sht2.autofit()
                wb2.save(r"output/" + raw_name + "_" + str(count) + ".xlsx")
                wb2.close()
                count = count + 1
            # last part contains the remaining part
            elif count == num:
                # fill in title
                sht2.range('A1').expand("right").value = sht1.range('A1').expand("right").value
                field = "$A$" + str((count - 1) * part + 2) + ":$" + self.letter + "$" + str(total)
                # Need to define the format of the column
                sht2.range('A1:F' + str(total - (count - 1) * part + 2)).number_format = '@'
                # fill in data
                sht2.range('A2').options(expand='table').value = sht1.range(field).value
                sht2.autofit()
                wb2.save(r"output/" + raw_name + "_" + str(count) + ".xlsx")
                wb2.close()
                count = count + 1
            else:
                # fill in title
                sht2.range('A1').expand("right").value = sht1.range('A1').expand("right").value
                field = "$A$" + str((count - 1) * part + 2) + ":$" + self.letter + "$" + str(count * part + 1)
                # fill in data
                sht2.range('A2').options(expand='table').value = sht1.range(field).value
                sht2.autofit()
                wb2.save(r"output/" + raw_name + "_" + str(count) + ".xlsx")
                wb2.close()
                count = count + 1

        wb1.app.quit()


if __name__ == '__main__':
    export = ExportNewData(fn, num, letter)
    export.import_excel()
