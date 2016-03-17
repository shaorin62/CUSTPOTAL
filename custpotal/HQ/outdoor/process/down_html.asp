<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<%
'	For Each item In request.form
'		response.write item & " : "& request.form(item) & "<br>"
'	Next
'	response.End
	Dim mpath : mpath = "C:\pds\print\"
	Dim dpath : dpath = "C:\pds\report\"
	Dim murl : murl = "http://10.110.10.86:6666/pds/print\"
	Dim durl : durl = "http://10.110.10.86:6666/pds/report\"
	Dim mFullPath, dFullPath , mfile, dfile
	Dim cyear : cyear = request("cyear")
	Dim cmonth : cmonth = request("cmonth")
	Dim fso : Set fso = CreateObject("scripting.filesystemobject")

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandType = adCmdText
	cmd.parameters.append cmd.createparameter("contidx", adInteger, adParamInput)
	cmd.parameters.append cmd.createparameter("cyear", adChar, adParamInput, 4)
	cmd.parameters.append cmd.createparameter("cmonth", adChar, adParamInput, 2)
'	For intLoop = 1 To request("contidx").count
'		response.write  request("contidx")(intLoop) & "<br>"
'
'Next
'response.write request("contidx").count
' response.write request("contidx")(1)
'response.end
%>
<script language='vbs'>
	Sub OnLoading()
<%
		For intLoop = 1 To request("contidx").count
			sql = "select report from wb_report_mst where contidx=? and cyear=? and cmonth=?"
			cmd.commandText = sql
			cmd.parameters("contidx").value = request("contidx")(intLoop)
			cmd.parameters("cyear").value = cyear
			cmd.parameters("cmonth").value = cmonth

			Set rs = cmd.execute

			If Not rs.eof Then
				mFullPath = mpath & rs(0)
				If fso.FileExists(mFullPath) Then
					Set mfile = fso.GetFile(mFullPath)
					response.write "document.all(""FileDownloadManager"").AddFile """& murl & "/" & rs(0) & """, " &mfile.size &vbcrlf
				End If
			End If
			rs.close

			sql = "select filename from wb_contact_md a inner join wb_report_dtl b on a.mdidx=b.mdidx where contidx=? and cyear=? and cmonth=?"
			cmd.commandText = sql
			cmd.parameters("contidx").value = request("contidx")(intLoop)
			cmd.parameters("cyear").value = cyear
			cmd.parameters("cmonth").value = cmonth
			Set rs = cmd.execute

			do until rs.eof
				dFullPath = dpath & rs(0)
				If fso.FileExists(dFullPath) Then
					Set dfile = fso.GetFile(dFullPath)
					response.write "document.all(""FileDownloadManager"").AddFile """& durl & "/" & rs(0) & """, " &dfile.size &vbcrlf
				End If
				rs.movenext
			Loop
			rs.close
		Next
	Set rs = Nothing
	Set cmd = Nothing
	Set fso = Nothing
%>
	end sub
	Sub OpenDownloadMonitor()
		winstyle= "height=445,width=445, status=no,toolbar=no,menubar=no,location=no"
		window.open "/hq/outdoor/popup/FileDownloadMonitor.htm",null,winstyle
	End Sub
	OpenDownloadMonitor()
</script>

 <BODY onload="OnLoading()">
		 <OBJECT ID="FileDownloadManager" height="200" width="450"
		 CodeBase = "http://10.110.10.86:6666/DEXTUploadX/DEXTUploadX.cab#version=2,8,2,0"
		 CLASSID="CLSID:535AE497-8E85-45F8-AF36-2DFCBCA8B68A"></OBJECT>
<!--
<OBJECT ID="FileDownloadManager" height="200" width="450"
		 CodeBase = "http://mms.raed.co.kr/DEXTUploadX/DEXTUploadX.cab#version=2,8,2,0"
		 CLASSID="CLSID:535AE497-8E85-45F8-AF36-2DFCBCA8B68A"></OBJECT>
		 -->
 </BODY>


