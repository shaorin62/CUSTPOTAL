<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
'	dim item
'
'	for each item in request.form
'		response.write item & " : " & request.form(item) & "<br>"
'	next
'	response.write

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
	if not isnull(gcategoryidx) then 		categoryidx = gcategoryidx
	if not isnull(mcategoryidx) then		categoryidx = mcategoryidx
	if not isnull(scategoryidx) then		categoryidx = scategoryidx
	if not isnull(dcategoryidx) then		categoryidx = dcategoryidx

	dim objrs, sql
	sql = "select categoryidx, categoryname, categorylvl , highcategoryidx from dbo.wb_category where categoryidx = " & categoryidx
	response.write sql
	call set_recordset(objrs, sql)

	objrs("categoryname").value = categoryname
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