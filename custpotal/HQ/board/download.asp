<%

	set fso = server.createobject("scripting.filesystemobject")
	Dim downloadfile : downloadfile = Server.HTMLEncode(request("filename"))
	Dim download


	if fso.FileExists("C:\pds\file\"&downloadfile) Then
		response.contenttype = "application/octet-stream"
		response.addheader	 "content-disposition", "attachment; filename=" &downloadfile
		response.addheader	 "content-transfer-Encoding", "binary"
		response.addheader "expires", "0"
	End If

	if not fso.FileExists("C:\pds\file\" & downloadfile) then
		response.write "<script> alert('�������� ���� �����Դϴ�.\n\n���� �Ǵ� �̵��� �����Դϴ�.'); history.back(); </script>"
		response.end
	else
		response.contenttype = "application/octet-stream"
		response.addheader	 "content-disposition", "attachment; filename=" & Server.HTMLEncode(downloadfile)
		response.addheader	 "content-transfer-Encoding", "binary"
		response.addheader "expires", "0"

		Set objStream = Server.CreateObject("ADODB.Stream")
		 objStream.Open
		 objStream.Type = 1
		 objStream.LoadFromFile "C:\pds\file\"&request("filename")

		 download = objStream.Read
		 Response.BinaryWrite download

		Set download = Nothing
		Set objStream = Nothing

		response.end
		


'		response.addheader "content-disposition", "attachment; filename=" & Server.HTMLEncode(request("filename"))
'		response.contenttype = "application/unknown"
'		response.cachecontrol = "publish"
'		response.expires = 0
'
'		Set objStream = Server.CreateObject("ADODB.Stream")
'		 objStream.Open
'		 objStream.Type = 1
'		 objStream.LoadFromFile "C:\pds\file\"&request("filename")
'
'		 download = objStream.Read
'		 Response.BinaryWrite download
'		Set objstream = nothing

	end if
%>