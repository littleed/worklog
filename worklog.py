import requests
import json
import getpass
import xlwt
from requests.auth import HTTPBasicAuth

def callForJSON(url, field):
	response = requests.get(url, headers=headers, auth=HTTPBasicAuth(username, password))
	if(response.ok):
		result = json.loads(response.content)[field]
		response.close
		return result;
	else:
		response.raise_for_status()

##Initial values
site='https://littleed.atlassian.net'
headers = {'Content-type': 'application/json'}

##Get credentials
username=input("Username:")
password=getpass.getpass()

##Get Project
project=input("Project Key:") or "TES"
##Get Board
url=site + '/rest/agile/1.0/board?type=scrum&projectKeyOrId='+project
print("ID\tBoard Name")
for board in callForJSON(url, "values"):
	print(str(board["id"]) + "\t" + board["name"])
board=input("BoardID:")

##Get Sprint
url=site + '/rest/agile/1.0/board/' + board + '/sprint'
print("ID\tSprint Name")
for sprint in callForJSON(url, "values"):
	print(str(sprint["id"]) + "\t" + sprint["name"])
sprint=input("SprintID:")

##Prep Spreadsheet
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Worklog")
print("KEY\tUSER\tTIME_SPENT(seconds)")
i=0
sheet1.write(i,0, "Issue")
sheet1.write(i,1, "Developer")
sheet1.write(i,2, "Time Spent")

##Get Issues in sprint
url = site + '/rest/agile/1.0/sprint/' + sprint + '/issue'
for issue in callForJSON(url, "issues"):
	key = issue["key"]
	##Get Worklogs for issue
	url = site + '/rest/api/2/issue/' + key + '/worklog'
	for log in callForJSON(url, "worklogs"):
		i=i+1
		sheet1.write(i,0, key)
		sheet1.write(i,1, log["author"]["key"] )
		sheet1.write(i,2, log["timeSpentSeconds"])
		print(key + "\t" + log["author"]["key"] + "\t" + str(log["timeSpentSeconds"]))

##Save Spreadsheet
book.save("worklog.xls")
print("Written to worklog.xls")

	
	
