<!--#include virtual="/inc/getdbcon.asp"-->
<!--#include virtual="/inc/func.asp"-->
<%

	dim sidx : sidx = request("sidx")

	dim objrs, sql
	sql = "select sidx, mdidx, side, standard, quality, unitprice from dbo.WB_MEDIUM_DTL where sidx = " & sidx
	call set_recordset(objrs, sql)
	objrs.delete()
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