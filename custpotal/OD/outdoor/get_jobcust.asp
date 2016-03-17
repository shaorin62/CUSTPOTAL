<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
	dim seqno : seqno = request("seqno")
	dim jobidx : jobidx = request("jobidx")

	response.write jobidx

	dim objrs, sql
	sql = "select jobidx, thema from dbo.wb_jobcust where seqno = '" & seqno & "' "
	call get_recordset(objrs, sql)

	dim str
	str = "<select name='selsubject' style='width:207px;'>"
	str = str & "<option value=''> 소재를 선택하세요. </option>"
	do until objrs.eof
	str = str & "<option value='"&objrs("jobidx")&"' "
		if int(jobidx) = int(objrs("jobidx")) then str = str & " selected"
	str = str & ">" & objrs("thema") & "</option>"
	objrs.movenext
	loop
	str = str & "</select>"
%>
<script language="JavaScript">
<!--
	var thema = top.document.getElementById("thema");
	thema.innerHTML = "<%=str%>";
//-->
</script>