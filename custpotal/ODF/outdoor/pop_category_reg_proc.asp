<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
'	dim item
'
'	for each item in request.form
'		response.write item & " : " & request.form(item) & "<br>"
'	next

	dim gcategoryidx : gcategoryidx = request("selgcategory")
	if gcategoryidx = "" then gcategoryidx = null
	dim mcategoryidx : mcategoryidx = request("selmcategory")
	if mcategoryidx = "" then mcategoryidx = null
	dim scategoryidx : scategoryidx = request("selscategory")
	if scategoryidx = "" then scategoryidx = null
	dim dcategoryidx : dcategoryidx = request("seldcategory")
	if dcategoryidx = "" then dcategoryidx = null
	dim categoryname  : categoryname = request("txtcategoryname")

	dim categoryidx, categorylvl
	categoryidx = null
	if not isnull(gcategoryidx) then
		categoryidx = gcategoryidx
		categorylvl = 1
	end if
	if not isnull(mcategoryidx) then
		categoryidx = mcategoryidx
		categorylvl = 2
	end if
	if not isnull(scategoryidx) then
		categoryidx = scategoryidx
		categorylvl = 3
	end if

	dim objrs, sql
	sql = "select categoryidx, categoryname, categorylvl , highcategoryidx from dbo.wb_category where highcategoryidx = " & categoryidx
	response.write sql
	call set_recordset(objrs, sql)

	objrs.addnew
	objrs("categoryname").value = categoryname
	objrs("categorylvl").value = categorylvl
	objrs("highcategoryidx").value = categoryidx
	objrs.update
	objrs.close

	set objrs = nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	window.opener.location.href = "category_list.asp";
	this.close();
//-->
</SCRIPT>