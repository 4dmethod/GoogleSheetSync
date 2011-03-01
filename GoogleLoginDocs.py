#!/usr/bin/python

try:
	import gdata.docs.service
except:
	exit("-100")  # required module not present

try:
	# Create a client class which will make HTTP requests with Google Docs server.
	gd_client = gdata.docs.service.DocsService()
	# Authenticate using your Google Docs email address and password.
	gd_client.ClientLogin('<!--#4DVAR <>tGoogleEmail-->', '<!--#4DVAR <>tGooglePassword-->')
except:
	exit("-101")
