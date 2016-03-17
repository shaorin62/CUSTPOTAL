<%
	dim uploadform : set uploadform = server.createobject("dext.fileupload")
	uploadform.defaultpath = "f:\wwwhome\eventfolder_dasfprx\pds\media"

	'response.write isobject(uploadform)
	response.write uploadform("txtfile")
%>