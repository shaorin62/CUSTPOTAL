<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '보안관련라이브러리 %>
<!--#include virtual="/inc/head.asp"-->			<% '초기 설정 페이지(에러 메세지 미출력) %>

<%
	Dim atag
	dim midx : midx = request("midx")
	dim title : title = request("txtTitle")
	dim file : file = request("chkfile")
	dim comment : comment = request("chkcomment")
	dim email : email = request("chkemail")
	dim mp : mp = request("chkMP")

	if file = "" then file = 0 else file = 1
	if comment = "" then comment = 0 else comment = 1
	if email = "" then email = 0 else email = 1
	if mp = "" then mp = 0 else mp = 1

	dim objrs, sql
	sql  = "select midx, title, custcode, lvl, isfile, iscomment, isemail, mp, isuse, cuser, cdate, uuser, udate, ref, highmenu from dbo.wb_menu_mst where midx = " & midx
	call set_recordset(objrs, sql)

	objrs.fields("title").value = clearXSS( title, atag)
	objrs.fields("isfile").value = file
	objrs.fields("iscomment").value = comment
	objrs.fields("isemail").value = email
	objrs.fields("mp").value = mp
	objrs.fields("uuser").value = request.cookies("userid")
	objrs.fields("udate").value = date

	objrs.update
	objrs.close
	Set objrs = Nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	//window.opener.location.reload();
	window.opener.document.location.href = window.opener.document.URL;
	location.href="pop_menu_view.asp?midx=<%=midx%>";
//-->
</SCRIPT>