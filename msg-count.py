import json
import xlsxwriter
from datetime import datetime 

input_file = open('skype-messages.json')
json_array = json.load(input_file)
finalArr = []
allDate = []
allName = []
expectedDAte =  "2020-03"

def getDisplayNames(messages):
	nameArr = []
	for msg in messages:
		if expectedDAte in msg["originalarrivaltime"][0:7]:
			nameArr.append(msg["displayName"])
	nameArr = list(set(nameArr))
	return nameArr

def getUniqueDate(newMessages):
	dateArr = []
	for msg in newMessages:
		if expectedDAte in msg["originalarrivaltime"][0:7]:
			dateArr.append(msg["originalarrivaltime"][0:10])
	dateArr = list(set(dateArr))
	return dateArr

def search(messages, name):
	newMsgArr = []
	newDateArr = []
	for msg in messages:
		if msg["displayName"] == name:
			newMsgArr.append(msg)
			newDateArr.append(msg["originalarrivaltime"][0:10])

	dates = getUniqueDate(newMsgArr)

	for currentDate in dates:
		finalArr.append([name, currentDate, newDateArr.count(currentDate)])

def searchAll(messages):
	nameArr = getDisplayNames(messages)
	for name in nameArr:
		search(messages, name)

def workingWithExcel(array, sheetname):
	worksheet = workbook.add_worksheet(sheetname)
	
	finalTuple = tuple(finalArr)
	col = 0
	for date in allDate:
		col=col+1
		worksheet.write(0, col, date)

	row = 0
	for name in allName:
		row = row + 1
		worksheet.write(row, 0, name)

	for name, date, count in (finalTuple):
		worksheet.write(allName.index(name)+1, allDate.index(date)+1, count)


workbook = xlsxwriter.Workbook('Counter.xlsx')
for item in json_array:
    searchAll(item['messages'])
    allDate = getUniqueDate(item['messages'])
    allDate.sort(key = lambda date: datetime.strptime(date, '%Y-%m-%d'))
    allName = getDisplayNames(item['messages'])
    workingWithExcel(finalArr, item['group'])
    finalArr = []

workbook.close()
