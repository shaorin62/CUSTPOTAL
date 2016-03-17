<!--#include virtual="/inc/getdbcon.asp"-->
<!--#include virtual="/inc/func.asp"-->
<%
'	dim item
'	for each item in request.form
'		response.write item & " :" & request.form(item) & "<BR>"
'	next
'	response.end

	dim seqno : seqno = request("selseqno")
	dim thema : thema = request("txtthema")


	dim objrs
	dim sql : sql = "select top 1 jobidx, seqno, thema, cuser, cdate, uuser, udate from dbo.wb_jobcust"
	call set_recordset(objrs, sql)

	objrs.addnew
	objrs.fields("seqno").value = seqno
	objrs.fields("thema").value = thema
	objrs.fields("cuser").value = request.cookies("userid")
	objrs.fields("cdate").value = date
	objrs.fields("uuser").value = request.cookies("userid")
	objrs.fields("udate").value = date
	objrs.update

	dim jobidx : jobidx = objrs("jobidx")

	objrs.close
	set objrs = nothing

%>
<script language="JavaScript">
<!--
	window.opener.location.reload();
	this.close();
//-->
</script>