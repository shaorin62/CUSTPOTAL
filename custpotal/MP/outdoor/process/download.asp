<!--#include virtual="/mp/outdoor/inc/function.asp" -->
<%
	Dim filename : filename = request("filename")
	Dim fso : Set fso = server.CreateObject("scripting.filesystemobject")
	if not fso.FileExists("\\11.0.12.201\adportal\print\" & filename) then
		response.write "<script> alert('�������� ���� �����Դϴ�.\n\n���� �Ǵ� �̵��� �����Դϴ�.'); history.back(); </script>"
		response.end
	else
		response.contenttype = "application/octet-stream"
		response.addheader	 "content-disposition", "attachment; filename=" & filename
		response.addheader	 "content-transfer-Encoding", "binary"
		response.addheader "expires", "0"

		Set objStream = Server.CreateObject("ADODB.Stream")
		 objStream.Open
		 objStream.Type = 1
		 objStream.LoadFromFile "\\11.0.12.201\adportal\print\"&filename

		 download = objStream.Read
		 Response.BinaryWrite download
		Set objstream = nothing

	end if
%>