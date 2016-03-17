<%
	dim uploadform : Set uploadform = Server.CreateObject("DEXT.FileUpload") 
	uploadform.defaultpath = "f:\wwwhome\eventfolder_dasfprx\pds\media"

	'response.write isobject(uploadform)
	response.write uploadform("txtfile")
%>