<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '보안관련라이브러리 %>
<!--#include virtual="/inc/head.asp"-->			<% '초기 설정 페이지(에러 메세지 미출력) %>

<%
'	Dim item
'	For Each item In request.Form
'		response.write item &  " :" & request.Form(item) & "<br>"
'	Next
'	response.end
	Dim atag
	dim midx : midx = request("midx")
	dim custcode : custcode = request("custcode")
	dim title : title = request("txtTitle")
	dim file : file = request("chkfile")
	if file = "" then file = 0 else file = 1
	dim comment : comment = request("chkcomment")
	if comment = "" then comment = 0 else comment = 1
	dim email : email = request("chkemail")
	if email = "" then email = 0 else email = 1

	dim objrs, sql, mp

	sql = "select isnull(mp,0) mp from dbo.wb_menu_mst where highmenu=1 and  midx= " & midx

	Call get_recordset(objrs, sql)

	if not objrs.eof then
		mp = obrs("mp")
	end if


	sql  = "select  midx, title, custcode, lvl, isfile, iscomment, isemail, isuse, cuser, cdate, uuser, udate, ref, highmenu, mp from dbo.wb_menu_mst where midx =" &midx
	call set_recordset(objrs, sql)

	objrs.addnew
	objrs.fields("title").value = clearXSS( title, atag)
	objrs.fields("custcode").value = custcode
	objrs.fields("lvl").value = 2
	objrs.fields("isfile").value = file
	objrs.fields("iscomment").value = comment
	objrs.fields("isemail").value = email
	objrs.fields("mp").value = mp
	objrs.fields("isuse").value = "Y"
	objrs.fields("cuser").value = request.cookies("userid")
	objrs.fields("cdate").value = date
	objrs.fields("uuser").value = request.cookies("userid")
	objrs.fields("udate").value = date
	objrs.fields("ref").value = midx

	objrs.update
	objrs.close

	Set objrs = Nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	this.close();
//-->
</SCRIPT>