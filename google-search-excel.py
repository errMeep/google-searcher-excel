import xlsxwriter
import xlrd

loc = (r"PATH")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 1)

out_wb = xlsxwriter.Workbook("out_book.xlsx")
out_sheet = out_wb.add_worksheet("out_sheet")

for i in range(7):
    print(sheet.cell_value(i+1, 0))
    sname=sheet.cell_value(i+1, 0)
    sstate=sheet.cell_value(i+1, 1)
    try:
        from googlesearch import search
    except ImportError:
        print("No module named 'google' found")
	 
	# to search
    query = sname +" "+sstate
    for j in search(query, tld="co.in", num=1, stop=1, pause=1.5):
	    out_sheet.write('A'+str(i),str(sname))
	    out_sheet.write('B'+str(i),j)
out_wb.close()
