<!--#include virtual="/inc/getdbcon.asp"-->
<!--#include virtual="/inc/func.asp"-->
<%
'	dim item
'	for each item in request.form
'		response.write item & " :" & request.form(item) & "<BR>"
'	next
'	response.end

	dim jobidx : jobidx = request("jobidx")
	dim gotopage : gotopage = request("gotopage")
	dim searchstring : searchstring = request("searchstring")
	dim seqno : seqno = request("selseqno")
	dim thema : thema = request("txtthema")


	dim objrs
	dim sql : sql = "select seqno, thema, uuser, udate from dbo.wb_jobcust where jobidx = " & jobidx
	call set_recordset(objrs, sql)

	objrs.fields("thema").value = thema
	objrs.fields("uuser").value = request.cookies("userid")
	objrs.fields("udate").value = date
	objrs.update

	objrs.close
	set objrs = nothing

%>
<script language="JavaScript">
<!--
	window.opener.location.reload();
	this.close();
//-->
</script>