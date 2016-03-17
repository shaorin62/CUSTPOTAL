<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim categoryidx : categoryidx = request("categoryidx")

	dim objrs, sql
	sql = "select categoryidx from dbo.wb_contact_md where categoryidx = " & categoryidx
	call get_recordset(objrs, sql)

	if not objrs.eof then
		response.write "<script> alert('매체분류 코드로 등록되어 있는 분류명입니다.\n\n삭제하시려면 등록된 코드를 모두 삭제하세요'); history.back(); </script>"
		response.end
	end if
	sql = "select categoryidx, categoryname, categorylvl, highcategoryidx from dbo.wb_category where categoryidx = " & categoryidx
	call set_recordset(objrs, sql)

	objrs.delete

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