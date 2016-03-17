<!--#Include virtual="/inc/getdbcon.asp"-->
<%
	dim uploadform : set uploadform = server.createobject("DEXT.FileUpload")
	uploadform.defaultPath = server.mappath("..")&"\pds\file"

	dim boardidx : boardidx = uploadform("txtboardidx")
	dim comments : comments = uploadform("txtcomments")
	dim filename 

	Dim objrs : Set objrs = server.createobject("adodb.recordset")
	objrs.activeconnection = application("connectionstring")
	objrs.cursorlocation = aduseclient
	objrs.cursortype = adopenstatic
	objrs.locktype = adlockoptimistic
	objrs.source = "dbo.WEB_BOARD_COMMENT"
	objrs.open

	If uploadform("txtfile") <> "" Then 
		uploadform("txtfile").save
		filename = uploadform("txtfile").filename
	Else
		filename = null
	End If

	objrs.addnew
	objrs.fields("BOARDIDX").value = boardidx
	objrs.fields("COMMENTS").value = Replace(comments, "'", Chr(39))
	objrs.fields("FILENAME").value = filename
	objrs.fields("CUSER").value = request.cookies("userid")
	objrs.fields("CDATE").value = now
	objrs.update

	objrs.close
	Set objrs = Nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	window.opener.location.reload();
	this.close();
//-->
</SCRIPT>