<!--#include virtual="/inc/getdbcon.asp" -->
<%
'	Dim item
'	For Each item In request.Form
'		response.write item &  " :" & request.Form(item) & "<br>"
'	Next
'	response.end

	Dim menuname : menuname = request.Form("txtmenuname")
	Dim highmenuidx : highmenuidx = request.Form("selhighmenuidx")
	Dim deptcode : deptcode = request.Form("txtdeptcode")
	Dim isfile : isfile = request.Form("chkfile")
	Dim isemail : isemail = request.Form("chkmail")
	Dim iscomment : iscomment = request.Form("chkcomment")
	Dim isuse : isuse = request.Form("isuse")
	if highmenuidx = "" then highmenuidx = null
	if isfile = "" then isfile = false else isfile = true
	if isemail = "" then isemail = false else isemail = true
	if iscomment = "" then iscomment = false else iscomment = true

	Dim objrs : Set objrs = server.CreateObject("adodb.recordset")
	objrs.activeconnection = application("connectionstring")
	objrs.cursorlocation = aduseclient
	objrs.cursortype = adopenstatic
	objrs.locktype = adlockoptimistic
	objrs.source = "dbo.WEB_BOARD_MENU"
	objrs.open

	objrs.addnew
	objrs.fields("MENUNAME").value = menuname
	objrs.fields("HIGHMENUIDX").value = highmenuidx
	objrs.fields("CUSTCODE").value = deptcode
	objrs.fields("ISFILE").value = isfile
	objrs.fields("ISEMAIL").value = isemail
	objrs.fields("ISCOMMENT").value = iscomment
	objrs.fields("ISUSE").value = isuse
	objrs.fields("CDATE").value = date
	objrs.fields("CUSER").value = request.cookies("userid")
	objrs.update

	objrs.close
	Set objrs = Nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	location.href = "mnu_list.asp";
//-->
</SCRIPT>