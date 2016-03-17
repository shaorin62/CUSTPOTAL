<!--#include virtual="/inc/getdbcon.asp" -->
<%
	dim idx : idx = request.form("idx")
	dim menuidx : menuidx = request.form("menuidx")
	dim gotopage : gotopage = request.form("gotopage")

	dim objrs : Set objrs = server.createobject("adodb.recordset")
	objrs.activeconnection = application("connectionstring")
	objrs.cursorlocation = aduseclient
	objrs.cursortype = adopenstatic
	objrs.locktype = adlockoptimistic
	objrs.source = "dbo.WEB_BOARD"
	objrs.open

	objrs.find = "IDX = " & idx

	dim attachFile : attachFile = server.mappath("..")&"\pds\file" & "\"& objrs("FILENAME")	
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
	location.href = "list.asp?menuidx=<%=menuidx%>&gotopage=<%=gotopage%>"
//-->
</SCRIPT>