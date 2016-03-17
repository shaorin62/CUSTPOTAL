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
	sql = "select * from dbo.wb_category where highcategoryidx = " & categoryidx
	call get_recordset(objrs, sql)

	if not objrs.eof then
		response.write "<script type='text/javascript'> alert('하위메뉴가 존재하는 분류명은 삭제할수 없습니다.'); this.close(); </script>"
		response.end
	end if
	objrs.close

	sql = "select * from dbo.wb_medium_mst where categoryidx = " & categoryidx
	call get_recordset(objrs, sql)

	if not objrs.eof then
		response.write "<script type='text/javascript'> alert('선택하신 분류명은 매체관리에서 사용중입니다.\n\n사용중인 분류명은 삭제할수 없습니다. .'); this.close(); </script>"
		response.end
	end if
	objrs.close

	sql = "select categoryidx, categoryname, categorylvl , highcategoryidx from dbo.wb_category where categoryidx = " & categoryidx
	response.write sql
	call set_recordset(objrs, sql)

	objrs.delete
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