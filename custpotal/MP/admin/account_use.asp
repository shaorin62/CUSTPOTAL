<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<%
	dim userid : userid = request("userid")

	dim objrs, sql
	sql = "select isuse from dbo.wb_account where userid = '" & userid & "' "
	call set_recordset(objrs, sql)

	if not objrs.eof then
		objrs.fields("isuse").value= "Y"
		objrs.update
		response.write "<script> alert('��������� �����Ǿ����ϴ�.'); </script>"
	end if
	objrs.close
	set objrs = nothing
%>
<script language="JavaScript">
<!--
	parent.location.reload();
//-->
</script>