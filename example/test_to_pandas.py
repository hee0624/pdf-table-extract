#coding: utf-8

import pandas as pd
import pdftableextract as pdf

pages = ["1"]
#pages = ["23"]
#cells = [pdf.process_page("example.pdf", p) for p in pages]
#cells = [pdf.process_page("./2015_002594.pdf", p) for p in pages]
cells = [pdf.process_page("table.pdf", p) for p in pages]
#print "type(cells): ", type(cells)	#<type 'list'> list of list
#print "cells: ", cells	#<col, row, ?, ?, ?, "content of the cell">, the last three "?" of useful information are <1, 1, 1>?
#cells: [[(col, row, ?, ?, ?, "content of the cell"), (col, row, ?, ?, ?, "content of the cell"),..., (col, row, ?, ?, ?, "content of the cell")]]
#lxw NOTE: check whether to deal with the multiple tables in the same pages here?

#flatten the cells structure
cells = [item for sublist in cells for item in sublist]
#print "type(cells): ", type(cells)	#<type 'list'> list
#print "cells: ", cells
#cells: [(col, row, ?, ?, ?, "content of the cell"), (col, row, ?, ?, ?, "content of the cell"),..., (col, row, ?, ?, ?, "content of the cell")]
#一页中单个表格的情况下,上面的两个cells基本一致，一页中多个表格呢？
#------------------------------------------------------------------------------------------------------------------------------

#without any options, process_page picks up a blank table at the top of the page.
#so choose table '1'
li = pdf.table_to_list(cells, pages)[-1]
#print "type(li): ", type(li)	#<type 'list'> list of list
#print li
"""
li:
[
["content of cell 0 in row 0", "content of cell 1 in row 0",... , "content of cell n in row 0"], 
["content of cell 0 in row 0", "content of cell 1 in row 0",... , "content of cell n in row 1"],
...
["content of cell 0 in row 0", "content of cell 1 in row 0",... , "content of cell n in row m"]
]
"""

#li is a list of lists, the first line is the header, last is the footer (for this table only!)
#column '0' contains store names
#row '1' contains column headings
#data is row '2' through '-1'

data =pd.DataFrame(li[2:-1], columns=li[1], index=[l[0] for l in li[2:-1]])
#print type(data) #<class 'pandas.core.frame.DataFrame'>
print data
data.to_excel("./result.xlsx", sheet_name="Sheet1")
