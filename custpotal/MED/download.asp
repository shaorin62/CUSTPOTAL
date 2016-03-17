<%@ Language=VBScript %>
<!--#include virtual="/hq/outdoor/inc/function.asp" -->
<%
	filename = Request.QueryString("filename")
'	response.write filename
'	response.end

	Dim fso : Set fso = server.CreateObject("scripting.filesystemobject")
	if not fso.FileExists("\\11.0.12.201\adportal\report\" & filename) then
		response.write "<script> alert('존재하지 않은 파일입니다.\n\n삭제 또는 이동된 파일입니다.'); history.back(); </script>"
		response.end
	else
		response.AddHeader "Content-Disposition","attachment;filename=" & filename
		Response.ContentType = "application/unknown"
		Response.CacheControl = "public"

		Set objDownload = Server.CreateObject("DEXT.FileDownload")
		objDownload.Download "\\11.0.12.201\adportal\report\"&filename
		Set objDownload = Nothing
	end if
	set fso = nothing
%>
