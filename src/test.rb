require "spreadsheet"
# require_relative 'spreadhelper'

# book = Spreadsheet.open("./test.xls")
# sheet1 = book.worksheet("Sheet1")
# p sheet1.class
# p sheet1["A1"]

bh = Spreadsheet::Workbook.new
# bh = BookHelper.open("./test.xls")
sheet1 = bh.create_worksheet
sheet1.name = "Sheet1"
# sheet1["A1:D5"] = [[1, 2, 3, 4], [2, 4, 6, 8], [3, 6, 9, 12], [4, 8, 12, 16], [5, 10, 15, 20, 25]]
# p sheet1["A1:D5"]
sheet1[1, 1] = 111

bh.write("test2.xls")
