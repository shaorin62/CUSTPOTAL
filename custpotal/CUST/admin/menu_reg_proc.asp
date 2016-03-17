<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '보안관련라이브러리 %>
<!--#include virtual="/inc/head.asp"-->			<% '초기 설정 페이지(에러 메세지 미출력) %>

<%
	Dim atag
	dim FLAG : FLAG = request("FLAG")
	dim custcode : custcode = request("custcode")
	dim title : title = request("txtTitle")
	dim file : file = request("chkfile")
	dim comment : comment = request("chkcomment")
	dim email : email = request("chkemail")

	if FLAG = "" then FLAG = null
	if custcode = "" then custcode = Null
	if file = "" then file = 0 else file = 1
	if comment = "" then comment = 0 else comment = 1
	if email = "" then email = 0 else email = 1

	dim objrs, sql
	sql  = "select midx, title, custcode, lvl, isfile, iscomment, isemail, isuse, cuser, cdate, uuser, udate, ref, highmenu from dbo.wb_menu_mst"
	call set_recordset(objrs, sql)

	objrs.addnew
	objrs.fields("title").value = clearXSS( title, atag)
	objrs.fields("custcode").value = custcode
	objrs.fields("lvl").value = 1
	objrs.fields("isfile").value = file
	objrs.fields("iscomment").value = comment
	objrs.fields("isemail").value = email
	objrs.fields("isuse").value = "Y"
	objrs.fields("cuser").value = request.cookies("userid")
	objrs.fields("cdate").value = date
	objrs.fields("uuser").value = request.cookies("userid")
	objrs.fields("udate").value = date
	objrs.fields("ref").value = null
	objrs.fields("highmenu").value = 0

	objrs.update
	dim midx : midx = objrs("midx")
	objrs.close
	sql  = "select midx, title, custcode, lvl, isfile, iscomment, isemail, isuse, cuser, cdate, uuser, udate, ref, highmenu from dbo.wb_menu_mst where midx=" & midx
	call set_recordset(objrs, sql)

	objrs("ref") = midx
	objrs.update
	objrs.close

	Set objrs = Nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	//window.opener.scriptFrame.location.reload();
	window.opener.document.location.href = window.opener.document.URL;
	this.close();
//-->
</SCRIPT>