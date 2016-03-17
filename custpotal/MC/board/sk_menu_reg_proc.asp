<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%

	dim custcode : custcode = request("custcode")
	dim custflag : custflag = request("FLAG")
	If custflag = "HIGHCUSTCODE" Then
		custflag = null
	Else
		custflag = 1
	End If

	dim title : title = request("txtTitle")
	dim file : file = request("chkfile")
	dim comment : comment = request("chkcomment")
	dim email : email = request("chkemail")

	if custcode = "" then custcode = null

	if file = "" then file = 0 else file = 1
	if comment = "" then comment = 0 else comment = 1
	if email = "" then email = 0 else email = 1

	dim objrs, sql

	sql  = "select max(midx)+1 midx from dbo.wb_menu_mst "
	call set_recordset(objrs, sql)
	dim midx : midx = objrs("midx")
	objrs.close

	sql  = "select midx, title, custcode, lvl, isfile, iscomment, isemail, isuse, cuser, cdate, uuser, udate, ref, highmenu, attr01 from dbo.wb_menu_mst"
	call set_recordset(objrs, sql)

	objrs.addnew
	objrs.fields("midx").value = midx
	objrs.fields("title").value = title
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
	objrs.fields("ref").value = midx
	objrs.fields("highmenu").value = 0
	objrs.fields("attr01").value = custflag
	objrs.update
	objrs.close


	Set objrs = Nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	//window.opener.location.reload();
	window.opener.document.location.href = window.opener.document.URL;
	this.close();
//-->
</SCRIPT>