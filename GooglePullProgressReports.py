#!/usr/bin/python

try:
	AdminEmail = '<!--#4DHTMLVAR <>tGoogleEmail-->'
	AdminPassword = '<!--#4DHTMLVAR <>tGooglePassword-->'
	TeacherDocName = <!--#4DHTMLVAR tGoogleDocName-->
except:
	exit("-99")  # data issue

try:
	import gdata.docs.service
	import gdata.spreadsheet
	import gdata.spreadsheet.service

except:
	exit("-100")  # required module not present

try:
	gd_client = gdata.docs.service.DocsService()
	gd_client.ClientLogin(AdminEmail, AdminPassword)
except:
	exit("-101")
	
#===================================================	
#        Progress report
#===================================================	
# Test for existence of the progress report
query = gdata.docs.service.DocumentQuery(categories=['spreadsheet', 'mine'], params={'title-exact': 'true'})
query['title'] = TeacherDocName
# query.AddNamedFolder(AdminEmail, TeacherFolderName)   #specifies to look for the course folder in the course folder
feed = gd_client.Query(query.ToUri())

try:
	if not feed.entry:
		exit("-102")
	else:
		progress_report_entry = feed.entry[0]                                   
except:
	exit("-103")

# Populate progress report data
try:
    gd_client = gdata.spreadsheet.service.SpreadsheetsService()
    gd_client.email = AdminEmail
    gd_client.password = AdminPassword
    gd_client.source = 'TestTakers Progress Report'
    gd_client.ProgrammaticLogin()
except:
	exit("-104")

spreadsheet_key = progress_report_entry.id.text.rsplit('/', 1)[1]
spreadsheet_key = spreadsheet_key.split('%3A', 1)[1]
#print spreadsheet_key
worksheets_feed = gd_client.GetWorksheetsFeed(spreadsheet_key)
worksheet_key = worksheets_feed.entry[0].id.text.rsplit('/', 1)[1]
#print worksheet_key

#===================================================	
#        Retreival Code
#===================================================	
try:
	query = gdata.spreadsheet.service.CellQuery()
	query['min-col'] = '1'
	query['max-col'] = '2'
	query['min-row'] = '2'
	query['return-empty'] = 'true'
	feed = gd_client.GetCellsFeed(spreadsheet_key, worksheet_key, query=query)
	#print feed.row_count.text 
	for entry in feed.entry:
		if entry.content.text is None:
			print entry.content.text
		else:
			print entry.content.text.replace('\n','<lf>')		
except:
	exit("-105")
