'''
Created on May 13, 2013
@author: Tom Eaton

Overview:
This program generates Musician's Friend and Guitar Center's
weekly shipment pull sheet for the warehouse.  This should
be run every Monday morning and printed to the warehouse.

Workflow:
1.  F Read backorder spreadsheet
2.    Verify column headings
3.  F Remove extraneous customer lines (keep MF & GC)
4.  F Remove extraneous columns
5.  F Remove non-due items *leave early ship items in
6.  F Except committed items, remove items with zero stock
7.  F Remove items with PO's not received
8.  F Remove ZZ- items
9.    Determine shippable quantity in ship field.  Assign available stock to newest orders first.
10. F Make master list tab and item order tab
11.   Make pretty & printable
12.   Print both tabs
13. F Add SLMTools to Git
14.   Web interface

---------------
'''
# Function that takes a list and a row number to write data to row on new_ws

import xlrd
import xlwt
import datetime
import os

''' Functions
-----------------------
'''

def find_ship_date():
	now = datetime.datetime.now()
	shift_fri = 5 - int(now.strftime('%w'))

	if int(now.strftime('%w')) < 3:
		shift = datetime.timedelta(days = 7)
		return_date = now + datetime.timedelta(days=shift_fri)
	else:
		shift = datetime.timedelta(days = 7)
		return_date = now + datetime.timedelta(days=shift_fri) + datetime.timedelta(days = 7)
	return return_date

'''
-----------------------
'''


''' Styles
-----------------------
header_style = xlwt.XFStyle()
font = xlwt.Font()
font.name = 'MF Sans Serif'
font.bold = False
font.size = 10
-----------------------
'''

# 1. Read backorder spreadsheet:
orig_wb = xlrd.open_workbook(os.path.dirname(__file__) + '\Back Orders 730AM.xls', logfile=open(os.devnull, 'w'))
orig_ws = orig_wb.sheet_by_index(0)

# Column headings:
col_head = ['PO', 'SO', 'Deliver', 'Early', 'Line', 'Code', 'Description', 'BO', 'Committed', 'Stock', 'Level II']

# Necessary columns from original:
keep_col = [0, 5, 6, 10, 12, 13, 14, 20, 22, 23, 24]

# Caculate shipping dates
ship_date = find_ship_date()
last_date = ship_date + datetime.timedelta(days = 11)

# Load desired rows into 2D array
saved_data = []
for cur_row in range(orig_ws.nrows):     
	new_row = []
	if (     (orig_ws.cell(cur_row,3).value == 'MUSICIA1') 
		 and (datetime.datetime.strptime(orig_ws.cell(cur_row,22).value, '%m/%d/%y') < last_date or orig_ws.cell(cur_row,20).value == 'T')
	     and (orig_ws.cell(cur_row,5).value[:3] != 'ZZ-')
		 and (orig_ws.cell(cur_row,12).value != 0 or orig_ws.cell(cur_row,13).value != 0)
		 and (orig_ws.cell(cur_row,18).value[len(orig_ws.cell(cur_row,18).value) - 5:len(orig_ws.cell(cur_row,18).value) - 3] != 'NR')
	   ):
		for cur_col in range(orig_ws.ncols):
			if cur_col in keep_col:
				new_row.append(orig_ws.cell(cur_row,cur_col).value)
		saved_data.append(new_row)

# New column order
new_order = [6, 0, 8, 7, 9, 1, 2, 3, 4, 5, 10]

final_data = []
for i, cur_row in enumerate(saved_data):
	new_row = []
	for x in new_order:
		new_row.append(cur_row[x])
	final_data.append(new_row)
		
'''
#ship_master = sorted(saved_data, key=lambda so_num: (so_num[0], so_num[9]))
ship_master = sorted(saved_data, key=lambda so_num: (so_num[4]), reverse=True)
ship_item_order = sorted(saved_data, key=lambda so_num: (so_num[1], datetime.datetime.strptime(so_num[8], '%m/%d/%y')))
'''

#Determine shippable quantity
'''
Sort by: Committed (Descending), Code (Ascending), Deliver (Ascending)
Copy Committed quantities over first
For item, find total needed, then assign stock to oldest order until out
'''

ship_master = sorted(final_data, key=lambda so_num: (so_num[8]), reverse=True)
ship_item_order = sorted(final_data, key=lambda so_num: (so_num[5], datetime.datetime.strptime(so_num[2], '%m/%d/%y')))

# Write to workbook
new_wb = xlwt.Workbook(encoding = 'ascii')
master_ws = new_wb.add_sheet('Master List')
itemord_ws = new_wb.add_sheet('Item Order')

# Add column headings
ship_master.insert(0, col_head)
ship_item_order.insert(0, col_head)

cur_row = 0
while cur_row < len(ship_master):
	cur_col = 0
	while cur_col < len(ship_master[0]):
		master_ws.write(cur_row, cur_col, label = ship_master[cur_row][cur_col])
		itemord_ws.write(cur_row, cur_col, label = ship_item_order[cur_row][cur_col])
		cur_col += 1
	cur_row += 1

# Save new worksheet		
file_name = os.path.dirname(__file__) + '\output\MF Weekly Shipment ' + datetime.datetime.now().strftime('%m-%d-%Y %I%M%p') + '.xls'
new_wb.save(file_name)


print
print '-----------------------------------------------------------------------------'
print
print 'New file generated: '
print file_name
print
print str(len(ship_master) - 1) + ' rows of data saved.'
print 'Ship date: ' + str(ship_date)[:10]
print '-----------------------------------------------------------------------------'