#coding: utf-8

import pandas as pd
import pdftableextract as pdf

pages = ["95"]
#pages = ["84"]

#cells = [pdf.process_page("example.pdf", p) for p in pages]
#cells = [pdf.process_page("./2015_601628.pdf", p) for p in pages]
cells = [pdf.process_page("./2015_002594.pdf", p) for p in pages]
#cells = [pdf.process_page("./2015_601766.pdf", p) for p in pages]
#cells = [pdf.process_page("table.pdf", p) for p in pages]
#cells = [pdf.process_page("table_up.pdf", p) for p in pages]
#cells = [pdf.process_page("table_down.pdf", p) for p in pages]
#cells = [pdf.process_page("table_down_1.pdf", p) for p in pages]
#exit(0)

#cells: [[(col, row, ?, ?, ?, "content of the cell"), (col, row, ?, ?, ?, "content of the cell"),..., (col, row, ?, ?, ?, "content of the cell")]]
#print "type(cells): ", type(cells)	#<type 'list'> list of list
#print "cells: ", cells	#lxw NOTE: <col, row, colspan?, ?, ?, "content of the cell">, what do the last three "?" mean? t

"""
print "cells:"
for cell in cells:	#cells: list of list
	print "len(cells) == 1? ", len(cells)
	for item in cell:	#cell: list of tuple
		tempStr = ""
		for element in item: #item: tuple
			tempStr += str(element) + ", "
		print tempStr
"""

#flatten the cells structure
cells = [item for sublist in cells for item in sublist]
#cells: [(col, row, ?, ?, ?, "content of the cell"), (col, row, ?, ?, ?, "content of the cell"),..., (col, row, ?, ?, ?, "content of the cell")]
#print "cells: ", cells	#type(cells)	#<type 'list'> list

"""
print "cells:"
for cell in cells:	#cells: list of tuple
	tempStr = ""
	for item in cell: #cell: tuple
		tempStr += str(item) + ", "
	print tempStr
"""

#check whether to deal with the multiple tables in the same pages here? SEEMED NOT, no obvious differences between two different tables.(通过横跨所有列来判断，也不保险，因为有些表格的中间可能存在一行横跨所有列的数据)

#------------------------------------------------------------------------------------------------------------------------------
#without any options, process_page picks up a blank table at the top of the page.
#so choose table '1'
li = pdf.table_to_list(cells, pages)[-1]
print "li: " 	#type(li)	#<type 'list'> list of list
for line in li:
	print ", ".join(line)
"""
li:
[
["content of cell 0 in row 0", "content of cell 1 in row 0",... , "content of cell n in row 0"], 
["content of cell 0 in row 1", "content of cell 1 in row 1",... , "content of cell n in row 1"],
...
["content of cell 0 in row m", "content of cell 1 in row m",... , "content of cell n in row m"]
]
"""

#li is a list of lists, the first line is the header, last is the footer (for this table only!)
#column '0' contains store names
#row '1' contains column headings
#data is row '2' through '-1'

#data =pd.DataFrame(li[2:-1], columns=li[1], index=[l[0] for l in li[2:-1]])
#data =pd.DataFrame(li[2:-1], columns=li[1])
#lxw NOTE: [2:-1] the last row is not included.
data =pd.DataFrame(li[1:-1])

#print type(data) #<class 'pandas.core.frame.DataFrame'>
print "data: ", data
data.to_excel("./result.xlsx", sheet_name="Sheet1")
