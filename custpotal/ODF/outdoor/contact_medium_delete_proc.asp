<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim contidx : contidx = request("contidx")
	dim sidx : sidx = request("sidx")

	dim objrs, sql
	sql = "select * from dbo.wb_contact_md_dtl where contidx="&contidx&" and sidx="&sidx
	
	call set_recordset(objrs, sql)

	if not objrs.eof then 
		do until objrs.eof 
			objrs.delete()
			objrs.movenext
		loop
	end if
	objrs.close
	sql = "select * from dbo.wb_contact_md where contidx="&contidx&" and sidx="&sidx
	
	call set_recordset(objrs, sql)

	if not objrs.eof then 
		objrs.delete()
	end if

	objrs.close
	set objrs = nothing


	
%>
<script language="JavaScript">
<!--
	document.location.replace("/hq/outdoor/pop_contact_view.asp?contidx=<%=contidx%>");
//-->
</script>