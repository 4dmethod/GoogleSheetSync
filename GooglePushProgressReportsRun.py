#!/usr/bin/python

try:
   AdminEmail = 'brento.is@gmail.com'
   AdminPassword = 'mypassword'
   TeacherEmail = 'bcorridan@gmail.com'
   TeacherFolderName = '''Corridan, Brian'''
   TeacherDocName = '''Corridan, Brian - S10SUMGNE - Progress Reports'''
   LocalPRTemplatePath = '/0_objsys/TestTakers/TestTakers v12/Resources/ProgressReportTemplate.xls'
   CellValuesList = [dict({'studentname':'''Doshi, Nirmita (Mita)''', 'comment':'''Mita's a great student, but her diagnostic scores are not fully reflecting the comprehension she shows in class. I truly believe that Mita can (and will!) get a higher SAT score, but to do so, she needs to be more careful with the strategies we've been reviewing in the course. In Math, for example, Mita made several careless errors on easier questions by either misreading the question or making simple math mistakes. In Reading, missing 6 of the 19 vocab-based questions cost her an additional 70 points. Though we will continue to practice all of these areas in class, I would like Mita to practice at home in her blue SAT book and to email me any questions she has! I am very confident that she can raise her score significantly in the coming weeks, but it will take diligence and practice. Any time she needs help, please have her contact me and I'll be glad to help!''', 'sentencecompletions':'X', 'rdg.compeliminate':'X', 'readentirechoice':'X', 'functionsgraphs':'X', 'targetquestions':'X', 'joebloggs':'X', 'workingcarefully':'X', 'erroridentification':'X', 'parallelism':'X', 'lengthok':'=if(LEN(R[0]C[-1])<885, "OK!", "TOO LONG!")'}), dict({'studentname':'''George, Merit''', 'comment':'''Merit is truly a fantastic student, and I'm happy to report that not only did he break 700 in Math, but he hit an 800 on Writing on Diag 2! With such strong success in those two subjects, Merit should be focusing his at-home practice on the Reading, which, while still very strong, can be raised even further. In particular, I'd like to see Merit work on finding answers WITHIN the passage itself, treating the SAT like the open-book test that it is. This will help him with both accuracy and timing (which appeared to be a bit of a problem on Diag 2, as he left 6 questions blank). Though we'll continue to practice in class, Merit would be wise to do at least one practice passage per week at home in his blue book. If he ever has questions, he should feel free to email me anytime. Keep up the great work, Merit!''', 'sentencecompletions':'X', 'rdg.compfind':'X', 'answeringassignedquestions':'X', 'geometry':'X', 'functionsgraphs':'X', 'targetquestions':'X', 'erroridentification':'X', 'lengthok':'=if(LEN(R[0]C[-1])<885, "OK!", "TOO LONG!")'}), dict({'studentname':'''Giarraputo, Brianna''', 'comment':'''After teaching Brianna at a couple of Make-Up and Extra Help classes last semester, it's nice to get the chance to be her formal teacher this time around! While she's clearly a very smart student (congrats on the 800 in Writing on Diag 2!), she made several careless errors in math that robbed her of nearly 50 points. She needs to work on slowing down and checking her work during the "easier" portion of the math problems (for example, on Diag 2, she missed 4 questions in the first half of the section, including a #1--the easiest question in the section). By taking the time to ensure that she hasn't made any simple mistakes, Brianna can make a big difference in her math score. In Reading, I'd like to see her continue to practice the Sentence Completions (vocab) portion by studying vocab and keeping track of the words we review that she didn't know. ''', 'sentencecompletions':'X', 'readentirechoice':'X', 'answeringassignedquestions':'X', 'geometry':'X', 'targetquestions':'X', 'plug-in':'X', 'erroridentification':'X', 'lengthok':'=if(LEN(R[0]C[-1])<885, "OK!", "TOO LONG!")'}), dict({'studentname':'''Gil, Michelle''', 'comment':'''Michelle has been a terrific student so far! Although her overall score has already improved significantly, her Reading score has been dropping, so I'd like to see her focus her at-home practice on this subject. In particular, the vocab seems to be Michelle's main area of weakness, as she missed 7 of the 19 vocab-based questions on Diag 2. (By contrast, she missed only 6 of the 48 passage-based questions.) The good news is that vocab is the easiest area to improve with practice; Michelle should simply continue to study her vocab box and keep track of all the words we discuss in class that she didn't already know. And if she ever needs help coming up with a mnemonic device, she should shoot me an email or ask in class and I'll be happy to help out. With more focus on the Sentence Completions, Michelle can raise that Reading score significantly!''', 'attendanceissue':'X', 'missedclass3':'X', 'sentencecompletions':'X', 'rdg.compeliminate':'X', 'answeringassignedquestions':'X', 'geometry':'X', 'erroridentification':'X', 'lengthok':'=if(LEN(R[0]C[-1])<885, "OK!", "TOO LONG!")'}), dict({'studentname':'''Kim, Jina''', 'comment':'''Though Jina's overall score has already started to improve, I know she has more in her! In particular, I'd like to see her practice the reading passages at home in her blue SAT book. As we've been discussing in class, answers are always located in the passages, so by going back into the passage to find answers that match something explicitly stated, Jina can gain back nearly 60 points in Reading! Vocabulary is also an area Jina should work on, as she missed 6 of the 19 vocab-based questions. Even getting just 2 or 3 more of those correct can raise her score 20-30 points. We'll continue to practice in class, but if Jina can come to the Tuesday night Extra Help sessions, I think that would benefit her as well. Overall, she's doing very well and should keep up the great work!''', 'attendanceissue':'X', 'missedclass3':'X', 'rdg.compfind':'X', 'usemta':'X', 'answeringassignedquestions':'X', 'arithmetic':'X', 'algebra':'X', 'targetquestions':'X', 'plug-in':'X', 'joebloggs':'X', 'workingcarefully':'X', 'erroridentification':'X', 'pronounagreement':'X', 'lengthok':'=if(LEN(R[0]C[-1])<885, "OK!", "TOO LONG!")'}), dict({'studentname':'''Sabbatini, Carla''', 'comment':'''In order to raise her score, Carla needs to implement all of the techniques we've been teaching her in class. In Math, for example, she missed 6 questions that could have been answered with Plug-In, Backsolve, or Guesstimate, costing her about 80 points. I'd like to see her practice these techniques at home in her blue SAT book to become comfortable enough with them to be able to use them properly on Diags 3 and 4. In Reading, vocab is a relative strong area, but the reading passages are costing her a lot of points. In particular, she missed 10 questions that had answers explicitly stated in the passage itself. By going back into the passage to find answers that match, Carla has the potential to see a 100+ point gain in Reading! If she is free on Tuesday nights, I'd recommend that she attend the Extra Help classes to get even more practice.''', 'attendanceissue':'X', 'misseddiag1':'X', 'rdg.compfind':'X', 'rdg.compeliminate':'X', 'readformainideas':'X', 'usemta':'X', 'readentirechoice':'X', 'dontuseoutsidebeliefs':'X', 'answeringassignedquestions':'X', 'geometry':'X', 'targetquestions':'X', 'plug-in':'X', 'workingcarefully':'X', 'erroridentification':'X', 'sentenceimprovement':'X', 'paragraphimprovement':'X', 'lengthok':'=if(LEN(R[0]C[-1])<885, "OK!", "TOO LONG!")'}), dict({'studentname':'''Sharma, Radhika''', 'comment':'''Radhika is a fantastic student, and I'm happy to report that she's already raised her score nearly 200 points! As her greatest improvements have been in Math and Writing so far, I'd like to see her concentrate her at-home practice on the Reading section. In particular, Radhika should keep practicing vocab, as she missed 6 of the 19 Sentence Completion (vocab-based) questions on Diag 2. To practice, Radhika should continue to learn all of the words in her box of flash cards, and she should also keep careful track of all of the words we discuss in class that she didn't already know. For the passages, she should remember to always go back into the passage to find answers that match, treating the SAT like the open-book test that it is. With continued practice at home in her blue SAT book, Radhika can improve her score significantly. ''', 'attendanceissue':'X', 'missedclass3':'X', 'sentencecompletions':'X', 'answeringassignedquestions':'X', 'geometry':'X', 'functionsgraphs':'X', 'targetquestions':'X', 'sentenceimprovement':'X', 'lengthok':'=if(LEN(R[0]C[-1])<885, "OK!", "TOO LONG!")'}), dict({'studentname':'''Sherman, Dylan''', 'comment':'''After seeing a bit of a drop on Diag 1, Dylan has brought his score back up above 2000. From here, though, I'd like to see additional improvement, particularly in the Reading section. As vocab is a relative strong area (having missed only 3 of the 19 questions), Dylan should practice the passages at home in his blue SAT book. On Diag 2, 8 of the 10 passage-based questions that he missed had answers explicitly stated within the passage itself! By going back into the passage to find answers that match, Dylan has the potential to raise his Reading score nearly 100 points. We'll continue to practice this in class, but to maximize his chances of acing the test, Dylan should try to do at least one passage a week at home as practice. If he ever has questions or needs help, he should email me and I'll be happy to help out!''', 'rdg.compfind':'X', 'readformainideas':'X', 'readentirechoice':'X', 'functionsgraphs':'X', 'targetquestions':'X', 'backsolve':'X', 'joebloggs':'X', 'erroridentification':'X', 'parallelism':'X', 'lengthok':'=if(LEN(R[0]C[-1])<885, "OK!", "TOO LONG!")'}), dict({'studentname':'''Su, Howard''', 'comment':'''It's really a pleasure to get the chance to teach Howard again, as he is a truly bright student. He continues to do well in class and on the diags, but I believe he should be getting an even higher score. In particular, Howard should focus his at-home practice on the Reading Comprehension, which is his single greatest area of weakness on the SAT, having missed 8 of the 48 questions on Diag 2. (By comparison, he answered all 19 vocab questions correctly.) By being more aggressive about finding answers that match INSIDE the passages, Howard can raise his Reading score significantly. We'll continue to practice this in class, but Howard should use his blue SAT book at home, as well (doing one passage a week is good practice). If he needs any help or has any questions, he should email me anytime!''', 'attendanceissue':'X', 'misseddiag1':'X', 'missedclass3':'X', 'rdg.compfind':'X', 'readformainideas':'X', 'readentirechoice':'X', 'dontuseoutsidebeliefs':'X', 'algebra':'X', 'targetquestions':'X', 'workingcarefully':'X', 'lengthok':'=if(LEN(R[0]C[-1])<885, "OK!", "TOO LONG!")'}), dict({'studentname':'''Yan, Winnie''', 'comment':'''Winnie is a fantastic student and it's so nice to see Bronx Science represented well! (In addition to being Winnie's Math and Reading instructor, I am the director of the TestTakers program at Bronx Science.) She's been doing a good job of steadily raising her Math score, but she still has room for improvement in her Reading and Writing scores. To raise her score, Winnie should focus on vocabulary, as she missed 5 of the 19 Sentence Completion (vocab-based) questions on Diag 2. To practice, Winnie must continue to study her vocab box and she should also keep a list of all the words we review in class that she didn't already know. We will continue to practice in all areas in class, but Winnie should also practice diligently at home in her blue SAT book. The more familiar she becomes with the Reading format, the better she will do on the real SAT!''', 'attendanceissue':'X', 'missedclass3':'X', 'sentencecompletions':'X', 'rdg.compeliminate':'X', 'usemta':'X', 'functionsgraphs':'X', 'targetquestions':'X', 'plug-in':'X', 'joebloggs':'X', 'erroridentification':'X', 'conciseexpression':'X', 'lengthok':'=if(LEN(R[0]C[-1])<885, "OK!", "TOO LONG!")'})]
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
