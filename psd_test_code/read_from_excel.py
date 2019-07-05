# Reading an excel file using Python 
import xlrd 
from pprint import pprint
import datetime
from time import time

# def setSuperScript(day):
# 	if day[1] == '1':
# 		if day== '11':
# 			supScrpt = 'th'
# 			return supScrpt
# 		else:
# 			supScrpt = 'st'
# 			return supScrpt
# 	elif day[1] == '2':
# 		if day== '12':
# 			supScrpt = 'th'
# 			return supScrpt
# 		else:
# 			supScrpt = 'nd'
# 			return supScrpt
# 	elif day[1] == '3':
# 		if day== '13':
# 			supScrpt = 'th'
# 			return supScrpt
# 		else:
# 			supScrpt = 'rd'
# 			return supScrpt
# 	else:
# 		supScrpt = 'th'
# 		return supScrpt


superscript_lst = {1:"st",2:"nd",3:"rd",4:"th",5:"th",6:"th",7:"th",8:"th",9:"th",10:"th",11:"th",12:"th",
                   13:"th",14:"th",15:"th",16:"th",17:"th",18:"th",19:"th",20:"th",21:"st",22:"nd",23:"rd",
                   24:"th",25:"th",26:"th",27:"th",28:"th",29:"th",30:"th",31:"st" }

def full_data_list(sheet):
    header = [ sheet.cell_value(0,col) for col in range(sheet.ncols)]
    final_list = list(map(lambda a:list(),header))
    print(final_list)
    dicto = dict(zip(header,final_list))
    for row in range(1,sheet.nrows):
        for col in range(sheet.ncols):
            dicto[header[col]].append(sheet.cell_value(row,col))
        
            # header.append(sheet.cell_value(0,col))
            # print(sheet.cell_value(row,col), end=" ")
        for each in dicto.keys():
            if each.strip() == "Visit Date":
                # print(dicto['Visit Date '])
                day=[]
                supscript = []
                year=[]
                for i in dicto['Visit Date ']:
                    date_list = i.split('.')
                    day.append(date_list[0])
                    supscript.append(superscript_lst[int(date_list[0])])
                    month=datetime.date(2019, int(date_list[1]), 1).strftime('%B')[0:3]
                    year.append(month+','+date_list[2])
                    # print(i)
    # print(type(each))
# pprint(dicto) 
    lst = []
    lst.append(day)
    lst.append(supscript)
    lst.append(year)
    dicto['Visit Date ']=lst
    print(dicto)


if __name__ =="__main__":
    loc = ("./MCKVIE.xls") 
    # To open Workbook 
    wb = xlrd.open_workbook(loc) 
    sheet = wb.sheet_by_index(0)
    start = time() 
    full_data_list(sheet)
    end = time()
    print("the time requrede is {}".format(end-start))
