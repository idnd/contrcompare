#!/usr/bin/env python 
# -*- coding: utf-8 -*- 
from xml.dom.minidom import *
from os import listdir
from os.path import isfile, join
import xlrd


def loadFilesInDir(dir_,workbookLoader):
	#enumirate file names in directory and load it
	onlyfiles = [ f for f in listdir(dir_) if isfile(join(dir_,f)) ]
	sapFiles = {}
	for fname in onlyfiles:
		sapFiles[fname] = workbookLoader(dir_ + fname)
	return sapFiles


def loadSapFiles(dir_):
	def loadWorkBookSap(fname):
		dom = parse(fname)
		wb = {}
		wsNum = 1
		Topic=dom.getElementsByTagName('Workbook')
		for node in Topic:
			for book in node.childNodes:
				if (book.nodeName == 'Worksheet'):
					ws = {}
					for worksheet in book.childNodes:
						if (worksheet.nodeName == 'Table'):
							rowNum = 1
							for row in worksheet.childNodes:
								if((row.nodeName) == 'Row'):
									cellNum = 1
									ws[rowNum] = {}
									for cell in row.childNodes:
										if((cell.nodeName) == 'Cell'):
											try:
												ws[rowNum][cellNum] = cell.getElementsByTagName("Data")[0].childNodes[0].data
												#print(cell.getElementsByTagName("Data")[0].childNodes[0].data)
											except:
												ws[rowNum][cellNum] = ''
											finally:
												cellNum += 1
									rowNum += 1
					wb[wsNum] = ws
		return wb

	return loadFilesInDir(dir_, loadWorkBookSap)


def loadIasFiles(dir_):
	# load ias files
	def loadWorkbookIas(fname):
		rb = xlrd.open_workbook(fname,formatting_info=True)
		wb = {}
		shNum = 1
		for s in rb.sheets():
			wb[shNum] = {}
			for rownum in range(s.nrows):
				wb[shNum][rowNum] = {}
				cellnum = 1
				row = s.row_values(rownum)
				for c_el in row:
					wb[shNum][rowNum][cellnum] = c_el
					cellnum += 1
		return wb

	return loadFilesInDir(dir_, loadWorkbookIas)


def filesToContracts(rawFileData, format):
	#convert raw file data to contract
	def fillContract(wb, typeInfo):
		''' convert file data to Contract '''
		def parseYear(date_):
			return  date[6:10]
		def parseMonth(date_):
			return  date_[3:5]
		def parseDay(date_):
			return date_[0:2]

		c = {}
		c['outnum'] = wb[format['outnum']['row']][format['outnum']['col']]
		c['flows'] = {}

		endOfFlowsTable = 0
		for row in range(typeInfo['first_row'],):
			if wb[row][typeInfo['date']] == '':
				endOfFlowsTable += 1
			else:
				endOfFlowsTable = 0
			if endOfFlowsTable == 2:
				break
			year = parseYear(wb[row][typeInfo['date']])
			month = parseYear(wb[row][typeInfo['date']])
			day = parseDay(wb[row][typeInfo['date']])
			if year not in c['flows']:
				c['flows'][year] = {}
			if month  not in c['flows'][year]:
				c['flows'][year][month] = {}
			if 'beginDate' not in c:
				c['beginDate'] = wb[row][typeInfo['date']]
				c['beginDateYear'] = year
				c['beginDateMonth'] = month
			c['endDate'] = wb[row][typeInfo['date']]
			c['endDateYear'] = year
			c['endDateMonth'] = month
			c['flows'][year][month][day]['mdti'] = wb[row][typeInfo['MDT_I']]
			c['flows'][year][month][day]['mdtd'] = wb[row][typeInfo['MDT_D']]
			c['flows'][year][month][day]['inti'] = wb[row][typeInfo['INT_CHRG']]
			c['flows'][year][month][day]['intd'] = wb[row][typeInfo['INT_PAY']]
			c['flows'][year][month][day]['comi'] = wb[row][typeInfo['COM_INC']]
			c['flows'][year][month][day]['comd'] = wb[row][typeInfo['COM_DEC']]
		return c
	#for each file in files fillContract
	contracts = {}
	for fname in rawFileData:
		tmp = fillContract[rawFileData[fname][1]]  # use only first sheet
		if tmp['outnum'] in contract:
			print('Contract outnum duplicate!')
			raise
		contract[tmp['outnum']] = tmp
	return contracts


def compare(contrSap, contrIas):
	def removeMissingOtherSide(sap,ias):
		#find missing docs
		#remove them (sync arrays)
		diff = set(sap.keys()) - set(ias.keys())
		print('Ias missed outnums:')
		for d in diff:
			print(d)
		diff = set(ias.keys()) - set(sap.keys())
		print('Sap missed outnums:')
		for d in diff:
			print(d)

		generalKeys = set(ias.keys()) - (set(ias.keys()) - set(sap.keys()))
		generalSap = {}
		generalIas = {}		
		for k in generalKeys:
			generalSap[k] = contractsSap[k]
			generalIas[k] = contractsIas[k]
		contractsSap = generalSap
		contractsIas = generalIas

	def checkCommonInfo(contrSap, contrIas):
		#check contract date etc
		pass

	def checkFlows(contrSap, contrIas):
		#check flows
		def calcIncDec(contr, year, month,incId,decId):
			inc = 0
			dec = 0
			for d in contr['flows']['year']['month']:
				inc += contr['flows']['year']['month'][d][incId]
				dec += contr['flows']['year']['month'][d][decId]
			return inc,dec

		for y in range(contrSap['beginDateYear'], contractsSap['endDateYear']):
			endMonth = contractsSap['endDateMonth'] if y == contractsSap['endDateYear'] else 12
			beginMonth = contractsSap['beginMonth'] if y == contractsSap['endDateYear'] else 1
			for m in range(beginMonth, endMonth):
				if calcIncDec(contractsIas, y, m, 'mdti', 'mdtd') != calcIncDec(contractsSap, y, m, 'mdti', 'mdtd'):  #  main debt inc, dec
					print(str(y) + '-' + str(m) + 'main debt is not equal')
				if calcIncDec(contractsIas, y, m, 'inti', 'intd') != calcIncDec(contractsSap, y, m, 'inti', 'intd'):  # sap interest inc, dec
					print(str(y) + '-' + str(m) + 'interest is not equal')
				if calcIncDec(contractsSap, y, m, 'comi', 'comd') != calcIncDec(contractsSap, y, m, 'comi', 'comd'):  # sap commissi inc, dec
					print(str(y) + '-' + str(m) + 'commission is not equal')
				
	removeMissingOtherSide(contrSap, contrIas)
	checkCommonInfo(contrSap, contrIas)
	checkFlows(contrSap, contrIas)


def main():
	sapFiles = loadSapFiles('C:\\trash\\ias\\')
	iasFiles = loadIasFiles('C:\\trash\\sap\\')
	sapFormat = {
		'outnum': {'row': 5, 'col': 7},
		'first_row': 11, 
		'date': 4, 
		'MDT_I': 8,
		'MDT_D': 12,
		'INT_CHRG': 17,
		'INT_PAY': 18,
		'COM_INC': 23,
		'COM_DEC': 24}

	iasFormat = {
		'outnum': {'row': 3, 'col': 5},
		'first_row': 11, 
		'date': 1, 
		'MDT_I': 4,
		'MDT_D': 8,
		'INT_CHRG': 13,
		'INT_PAY': 14,
		'COM_INC': 18,
		'COM_DEC': 19}

	contractsSap = filesToContracts(sapFiles, sapFormat)
	contractsIas = filesToContracts(iasFiles, iasFormat)

	compare(contractsSap, contractsIas)


if __name__ == '__main__':
	main()