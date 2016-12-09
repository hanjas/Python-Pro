import xlrd, datetime, json, csv, time
from itertools import product
from datetime import date, timedelta
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive


mimetypes = {
    'application/vnd.google-apps.document': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'application/vnd.google-apps.spreadsheet': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
}

def listFiles():
	gauth = GoogleAuth('settings.yaml')
	gauth.LocalWebserverAuth()
	drive = GoogleDrive(gauth)
	file_list = drive.ListFile({'q': "'root' in parents"}).GetList()
	for file1 in file_list:
	  print('title: %s, id: %s' % (file1['title'], file1['id']))

	# Paginate file lists by specifying number of max results
	for file_list in drive.ListFile({'maxResults': 10}):
	  print 'Received %s files from Files.list()' % len(file_list) # <= 10
	  for file1 in file_list:
	    print('title: %s, id: %s' % (file1['title'], file1['id']))


def downloadFile():
	gauth = GoogleAuth('out/settings.yaml')
	gauth.LocalWebserverAuth()
	drive = GoogleDrive(gauth)
	myfile = drive.CreateFile({'id': '1CLYrVDTixpPRCnrQG6uzoboW6c8IfI-TXY-Q1pFH5jM'})
	today = str(time.strftime("%d-%b-%y"))
	if myfile['title'] == 'OLA: Inventory & Fitment Tracker':
		download_mimetype = mimetypes[myfile['mimeType']]
		myfile.GetContentFile(today+'.xlsx', mimetype=download_mimetype)
		return today
	return False

def csv_from_excel(filename):
	try:
		book = xlrd.open_workbook(filename+'.xlsx')
		sheet = book.sheet_by_name('Fitments')
		csvfile = open(filename+'.csv', 'wb')
		wr = csv.writer(csvfile, quoting=csv.QUOTE_ALL)

		for rownum in range(sheet.nrows):
		    date = sheet.row_values(rownum)[1]
		    if isinstance( date, float) or isinstance( date, long ):
		        year, month, day, hour, minute, sec = xlrd.xldate_as_tuple(date, book.datemode)
		        py_date = "%02d-%02d-%04d" % (day, month, year)
		        wr.writerow(sheet.row_values(rownum)[0:1] + [py_date] + sheet.row_values(rownum)[2:])
		    else:
		    	# print sheet.row_values(rownum)
		        wr.writerow(sheet.row_values(rownum))
		csvfile.close()
		return True
	except Exception as e:
		print e
		return False

def strtConver(yesterday, today):
	if yesterday != today:
		f1 = file(yesterday+'.csv', 'r')
		f2 = file(today+'.csv', 'r')

		c1 = csv.reader(f1)
		c2 = csv.reader(f2)

		fst = []
		fst2 = []
		scnd = []
		scnd2 = []

		for fstrow in c1:
			for data in fstrow:
				fst.append(data)
			fst2.append(fst)
			fst = []
		for scndrow in c2:
			for data in scndrow:
				scnd.append(data)
			scnd2.append(scnd)
			scnd = []
		changes = {}
		for datas in scnd2:
			if not datas in fst2:
				##############
				for row in range(len(scnd2)):
					if scnd2[row] == datas:
						for col in range(len(scnd2[row])):
							if(datas[col] != fst2[row][col]):
								changes[row] = {}
								changes[row][ scnd2[1][col] ] = {}
								changes[row][ scnd2[1][col] ]['new'] = scnd2[row][col]
								changes[row][ scnd2[1][col] ]['old'] = fst2[row][col]
								
				############## 
		# print changes
		f1.close()
		f2.close()
		return changes
	else:
		return False

def getDetails(columnName, value):
	if(columnName != ''):
		today = str(time.strftime("%d-%b-%y"))
		wb = xlrd.open_workbook(today+'.xlsx')
		details = []
		for sheet in wb.sheets():
			if(sheet.name == 'Fitments'):
				for row, col in product(range(sheet.nrows), range(sheet.ncols)):
					if sheet.cell(row, col).value == columnName:
						colNames = sheet.row(1)
						for row_index in range(2, sheet.nrows):
							col_values = {}
							if( str(sheet.cell(row_index, 0).value) !='' and str(sheet.cell(row_index, 1).value) !=''):
								if str(sheet.cell(row_index, col).value).replace('.0', '') == value:
									for name , col_index in zip(colNames, range(sheet.ncols)):
										if (str(name.value)) !='' and str(sheet.cell(row_index, col_index).value) !='':
											if str(name.value) == 'DATE':
												date =  datetime.datetime(*xlrd.xldate_as_tuple(sheet.cell(row_index, col_index).value, wb.datemode)).strftime('%d-%b-%y')
												col_values[(str(name.value))] = date
											else:
												col_values[(str(name.value))] = str(sheet.cell(row_index, col_index).value).replace('.0', '')
								if col_values != {}:
									details.append(col_values)
				return details
	else:
		details.append('Something is wrong')
		return details


if __name__ == '__main__':
	print '\n trying to download file'
	today = downloadFile()
	yesterday = str((date.today() - timedelta(1)).strftime('%d-%b-%y'))
	if today:
		print '\n downloading successfull'
		print '\n now trying to convert xlsx to csv'
		if csv_from_excel(today):
			if csv_from_excel(yesterday):
				print '\n conversion successfull\n\n now comparing two files'
				strtConver(yesterday, today)
	# print strtConver(yesterday, today)
				print '\n open results.csv to check difference'
			else:
				strtConver(today, today)
		else:
			print 'something wrong while converting to csv'
	else:
		print 'no file downloaded'
