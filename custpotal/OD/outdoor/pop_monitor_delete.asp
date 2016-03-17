<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim fso : set fso = server.createobject("scripting.filesystemobject")
	Dim pidx : pidx = request("pidx")
	Dim didx : didx = request("didx")
	dim attachFile : attachFile = server.mappath("..") & "pds\monitor"
	dim objrs
	Dim sql : sql = "select didx, pidx, typical, filename from dbo.wb_contact_monitor_dtl where didx = " & didx
	call set_recordset(objrs, sql)

	if not objrs.eof then 
		attachFile = attachFile & "\" & objrs("filename")
		if fso.fileexists(attachFile) then 	fso.deletefile(attachFile)
		objrs.delete
	end if 
	objrs.close
	set objrs = nothing
%>
<script language="JavaScript">
<!--
	location.href = "pop_monitor_edit.asp?pidx=<%=pidx%>";
//-->
</script>