<%@ Language=VBScript %>
<!--#include virtual="/hq/outdoor/inc/function.asp" -->
<%
	filename = Request.QueryString("filename")
'	response.write filename
'	response.end

	Dim fso : Set fso = server.CreateObject("scripting.filesystemobject")
	if not fso.FileExists("\\11.0.12.201\adportal\report\" & filename) then
		response.write "<script> alert('�������� ���� �����Դϴ�.\n\n���� �Ǵ� �̵��� �����Դϴ�.'); history.back(); </script>"
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
