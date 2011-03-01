#!/usr/bin/python

try:
   AdminEmail = '<!--#4DHTMLVAR <>tGoogleEmail-->'
   AdminPassword = '<!--#4DHTMLVAR <>tGooglePassword-->'
   TeacherEmail = '<!--#4DHTMLVAR tGoogleShareEmail-->'
   TeacherFolderName = <!--#4DHTMLVAR tGoogleTeacherFolder-->
   TeacherDocName = <!--#4DHTMLVAR tGoogleDocName-->
   LocalPRTemplatePath = '<!--#4DHTMLVAR tGoogleLocalTemplatePath-->'
   CellValuesList = <!--#4DHTMLVAR tGoogleCellValues-->
   ScoreReportName = ''
   ScoreReportPath = ''
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

scope = gdata.docs.Scope(value=TeacherEmail, type='user')
role = gdata.docs.Role(value='writer')
acl_entry = gdata.docs.DocumentListAclEntry(scope=scope, role=role)

#===================================================	
#         Teacher folder
#===================================================	
#Test for existence of teacher folder
query = gdata.docs.service.DocumentQuery(categories=['folder', 'mine'], params={'showfolders': 'true', 'title-exact': 'true'})
query['title'] = TeacherFolderName
feed = gd_client.Query(query.ToUri())

try:
	if not feed.entry:
		#create the folder
		teacher_folder_entry = gd_client.CreateFolder(TeacherFolderName)
	else:
		#check the sharing
		teacher_folder_entry = feed.entry[0]
except:
	exit("-102")

#update the ACL on the teacher folder
try:
	created_acl_entry = gd_client.Post(acl_entry, teacher_folder_entry.GetAclLink().href,converter=gdata.docs.DocumentListAclEntryFromString)
except:
	pass
	
#===================================================	
#        Progress report template
#===================================================	
# Test for existence of the progress report
query = gdata.docs.service.DocumentQuery(categories=['spreadsheet', 'mine'], params={'title-exact': 'true'})
query['title'] = TeacherDocName
query.AddNamedFolder(AdminEmail, TeacherFolderName)   #specifies to look for the course folder in the course folder
feed = gd_client.Query(query.ToUri())

try:
	ms = gdata.MediaSource(file_path=LocalPRTemplatePath, content_type=gdata.docs.service.SUPPORTED_FILETYPES['XLS'])
	if not feed.entry:
		#upload the template doc  
		progress_report_entry = gd_client.Upload(ms, TeacherDocName, folder_or_uri=teacher_folder_entry)
	else:
		#progress report exists, overwrite
		progress_report_entry = gd_client.Put(ms, feed.entry[0].GetEditMediaLink().href)                                   
except:
	exit("-103")

#update the ACL on the progress report
try:
	created_acl_entry = gd_client.Post(acl_entry, progress_report_entry.GetAclLink().href,converter=gdata.docs.DocumentListAclEntryFromString)
except:
	pass

#===================================================	
#        Score report pdf
#===================================================	
if ScoreReportPath != '':
	# Test for existence of the score report
	query = gdata.docs.service.DocumentQuery(categories=['PDF', 'mine'], params={'title-exact': 'true'})
	query['title'] = ScoreReportName
	query.AddNamedFolder(AdminEmail, TeacherFolderName)   #specifies to look for the course folder in the course folder
	feed = gd_client.Query(query.ToUri())

	try:
		msPDF = gdata.MediaSource(file_path=ScoreReportPath, content_type=gdata.docs.service.SUPPORTED_FILETYPES['PDF'])
		if not feed.entry:
			#upload the score report doc  
			score_report_entry = gd_client.Upload(msPDF, ScoreReportName, folder_or_uri=teacher_folder_entry)
		else:
			#score report exists, overwrite
			score_report_entry = gd_client.Put(msPDF, feed.entry[0].GetEditMediaLink().href)                                   
	except:
		exit("-110")

	#update the ACL on the progress report
	try:
		created_acl_entry = gd_client.Post(acl_entry, score_report_entry.GetAclLink().href,converter=gdata.docs.DocumentListAclEntryFromString)
	except:
		pass
else:
	pass

#print 'Spreadsheet now accessible online at:', progress_report_entry.GetAlternateLink().href                       
#print 'Edit link:', progress_report_entry.GetEditLink().href                       

#===================================================	
# Populate progress report data
#===================================================	
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

LinkMe = progress_report_entry.GetAlternateLink().href
try:
	for CellValues in CellValuesList:
		entry = gd_client.InsertRow(CellValues, spreadsheet_key, worksheet_key)
#if isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
# 			LinkMe = progress_report_entry.GetAlternateLink().href
except:
	exit("-105")

print LinkMe
