<!--#include virtual="/inc/getdbcon.asp" -->
<%

'	dim item 
'	for each item in request.querystring
'		response.write item & " : " & request.querystring(item) & "<br>"
'	next
'	response.end
	dim gotopage : gotopage = request("gotopage")
	if gotopage = "" then gotopage = 1

	dim commentsidx : commentsidx = request("commentsidx")
	dim menuidx : menuidx = request("menuidx")
	dim idx : idx = request("idx")
	dim filename : filename = request("filename")
	
	dim attachFile : attachFile = server.mappath("..")&"\pds\file" & "\"&filename	
	dim objrs : set objrs = server.CreateObject("adodb.recordset")
	objrs.activeconnection = application("connectionstring")
	objrs.cursorlocation = aduseclient
	objrs.cursortype = adopenstatic
	objrs.locktype = adlockoptimistic
	objrs.source = "dbo.WEB_BOARD_COMMENT"
	objrs.open

	objrs.find = "IDX = " & commentsidx 
	dim fso : set fso = server.createobject("scripting.filesystemobject")
	if fso.fileexists(attachFile) then 
		fso.deletefile(attachFile)
	end if
	objrs.delete()
	objrs.update
	objrs.close
	Set objrs = Nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	location.href = "view.asp?idx=<%=idx%>&menuidx=<%=menuidx%>&gotopage=<%=gotopage%>";
//-->
</SCRIPT>