<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim categoryname : categoryname = request("txtcategoryname")
	dim categoryidx : categoryidx = request("txtcategoryidx")

	dim objrs, sql
	sql = "select categoryidx, categoryname, categorylvl, highcategoryidx from dbo.wb_category where categoryidx = " & categoryidx
	call set_recordset(objrs, sql)

	dim categorylvl : categorylvl = objrs("categorylvl")

	objrs.addnew
	objrs("categoryname") = categoryname
	objrs("categorylvl") = categorylvl+1
	objrs("highcategoryidx") =categoryidx

	objrs.update

	objrs.close

	set objrs = nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	window.opener.location.href ="/od/outdoor/category_list.asp";
	this.close();
//-->
</SCRIPT>