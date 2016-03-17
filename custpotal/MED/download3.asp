<%@CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include virtual="/hq/outdoor/inc/function.asp" -->
<%
	Dim filename : filename = request("filename")
'	response.write "C:\pds\report\" & filename
'	response.End
	Dim fso : Set fso = server.CreateObject("scripting.filesystemobject")
	if not fso.FileExists("\\11.0.12.201\adportal\report\" & filename) then
		response.write "<script> alert('존재하지 않은 파일입니다.\n\n삭제 또는 이동된 파일입니다.'); history.back(); </script>"
		response.end
	Else
		response.contenttype = "application/octet-stream"
		response.addheader	 "content-disposition", "attachment; filename=" & filename
		response.addheader	 "content-transfer-Encoding", "binary"
		response.addheader "expires", "0"

		Set objStream = Server.CreateObject("ADODB.Stream")
		 objStream.Open
		 objStream.Type = 1
		 objStream.LoadFromFile "\\11.0.12.201\adportal\report\"&filename

		 download = objStream.Read
		 Response.BinaryWrite download
		Set objstream = nothing

	end if
%>