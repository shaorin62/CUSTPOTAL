<!--#Include virtual="/inc/getdbcon.asp"-->
<%
	dim uploadform : set uploadform = server.createobject("DEXT.FileUpload")
	uploadform.defaultPath = server.mappath("..")&"\pds\file"

	dim idx : idx = uploadform("idx")
	dim menuidx : menuidx = uploadform("menuidx")
	dim gotopage : gotopage = uploadform("gotopage")
	dim subject : subject =  uploadform("txtsubject")
	dim contents : contents = uploadform("txtcontents")
	dim u_date : u_date = Date
	dim u_user : u_user = Request.cookies("userid")
	dim email : email = uploadform("txtemail")
	dim attachFile : attachFile = uploadform("txtattach")
	attachFile = uploadform.defaultPath & "\"&attachFile
	dim filename

	dim objrs : Set objrs = server.createobject("adodb.recordset")
	objrs.activeconnection = application("connectionstring")
	objrs.cursorlocation = aduseclient
	objrs.cursortype = adopenstatic
	objrs.locktype = adlockoptimistic
	objrs.source = "dbo.WEB_BOARD"
	objrs.open

	If uploadform("txtfile") <> "" Then
		dim fso : set fso = server.createobject("scripting.filesystemobject")
		if fso.fileexists(attachFile) then
			fso.deletefile(attachFile)
		end if
		set fso = nothing
		filename = uploadform("txtfile").save (, false)
		filename = right(filename, len(filename)-InStrRev(filename, "\"))
		objrs.fields("FILENAME").value = filename
	End If

	If email = "" Then
		email = null
	End if

	objrs.find = "IDX = " & idx

	objrs.fields("SUBJECT").value = Replace(subject, "'", Chr(39))
	objrs.fields("CONTENTS").value = Replace(contents, "'", Chr(39))
	objrs.fields("UDATE").value = u_date
	objrs.fields("UUSER").value = u_user
	if uploadform("txtemail") <> "" then objrs.fields("EMAIL").value = email end if
	objrs.update

	objrs.close
	Set objrs = Nothing

%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	location.href = "view.asp?idx=<%=idx%>&menuidx=<%=menuidx%>&gotopage=<%=gotopage%>"
//-->
</SCRIPT>