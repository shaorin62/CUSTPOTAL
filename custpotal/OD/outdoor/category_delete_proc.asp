<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim categoryidx : categoryidx = request("categoryidx")

	dim objrs, sql
	sql = "select categoryidx from dbo.wb_contact_md where categoryidx = " & categoryidx
	call get_recordset(objrs, sql)

	if not objrs.eof then
		response.write "<script> alert('��ü�з� �ڵ�� ��ϵǾ� �ִ� �з����Դϴ�.\n\n�����Ͻ÷��� ��ϵ� �ڵ带 ��� �����ϼ���'); history.back(); </script>"
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