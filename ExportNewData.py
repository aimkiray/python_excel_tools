#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time   : 3/6/2018
# @Author : Sun Zhang
# @Email  : root@meowwoo.com

import xlwings as xw

# old excel file path
old_fn = r"C:\Users\filename.xlsx"
# new excel file path
new_fn = r"C:\Users\filename.xlsx"
# Maximum column letter (fix bug)
letter = 'W'
# Attention: If you want to define column format, please find the corresponding code.

# The output file is in the output folder.


class ExportNewData(object):

    def __init__(self, old_fn, new_fn, letter):
        super(ExportNewData, self).__init__()
        self.ExistSet = set()
        self.NewListExport = list()
        self.NewListAddress = list()
        self.CompareList = list()
        self.old_fn = old_fn
        self.new_fn = new_fn
        self.letter = letter

    def import_excel(self):
        app = xw.App(visible=True, add_book=False)

        wb1 = app.books.open(self.old_fn)
        wb2 = app.books.open(self.new_fn)
        wb3 = app.books.add()

        sht1 = wb1.sheets['Sheet1']
        sht2 = wb2.sheets['Sheet1']
        sht3 = wb3.sheets['Sheet1']

        range1 = sht1.range("A1").expand("down")
        range2 = sht2.range("A1").expand("down")

        # Existing data is saved in the set to increase the speed of comparison.
        for name in range1.value:
            self.ExistSet.add(str(name).strip())

        # Need to compare the data
        self.CompareList = range2.value
        for index in range(len(self.CompareList)):
            if str(self.CompareList[index]).strip() not in self.ExistSet:
                self.NewListAddress.append("A" + str(index + 1) + ":" + self.letter + str(index + 1))

        while self.NewListAddress:
            td = self.NewListAddress.pop()
            self.NewListExport.append(sht2.range(td).value)

        # fill in title
        sht3.range('A1:' + self.letter + '1').value = sht2.range('A1:' + self.letter + '1').value
        # Need to define the format of the column
        sht3.range('A1:F' + str(len(self.NewListExport) + 1)).number_format = '@'
        # fill in data
        sht3.range('A2').options(expand='table').value = self.NewListExport
        # save
        sht3.autofit()
        wb3.save(r"output/new.xlsx")
        wb3.app.quit()


if __name__ == '__main__':
    export = ExportNewData(old_fn, new_fn, letter)
    export.import_excel()
