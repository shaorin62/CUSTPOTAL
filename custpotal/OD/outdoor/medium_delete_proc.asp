<!--#include virtual="/inc/getdbcon.asp"-->
<!--#include virtual="/inc/func.asp"-->
<%

	dim fso : set fso = server.createobject("scripting.filesystemobject")
	dim mdidx : mdidx = request("mdidx")
	dim map : map = request("txtmap")
	dim objrs, sql
	sql = "select mdidx from dbo.wb_contact_md where mdidx="&mdidx
	call get_recordset(objrs, sql)

	if not objrs.eof then
		response.write "<script type='text/javascript'> alert('계약에 등록된 매체입니다.\n\n삭제하실 수 없습니다.'); history.back(); </script>"
		response.flush
		respons.end
	end if

	objrs.close

	sql = "select * from dbo.WB_MEDIUM_MST where mdidx = " & mdidx
	call set_recordset(objrs, sql)



	dim attachFile : attachFile = server.mappath("..")&"\pds\media" & "\"& objrs("map")
	if fso.fileexists(attachFile) then fso.deletefile(attachFile)
	set fso = nothing

	objrs.delete()

	objrs.update
	objrs.close
	set objrs = nothing

%>
<script language="JavaScript">
<!--
	document.location.replace("/od/outdoor/medium_list.asp");
//-->
</script>