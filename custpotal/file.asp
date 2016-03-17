<%

set objConn = server.createobject("ADODB.Connection")

strCon  ="provider=sqloledb; data source=10.110.10.86; initial catalog=mcdev_new; user id=devadmin; password = password"
objConn.open strCon

Set objfso = Server.CreateObject("Scripting.FileSystemObject")

Dim strFileSearch
Dim strGbn
strFileSearch = "SELECT FILENAME FROM FILES  WHERE ISNULL(FILENAME,'') <> ''"

Set rs=Createobject("adodb.recordset")
rs.Open strFileSearch, strCon,1

Do until rs.Eof
temp_filename = rs("FILENAME")
temp_no = temp_filename
temp_filename = "C:\imsi\"&rs("FILENAME")

'response.write temp_filename & "<br>"

if not objfso.FileExists(temp_filename) Then

strSQL = "update files set flag = 'N' where filename = '" & temp_filename & "'"
Else
strSQL = "update files set flag = 'Y' where filename = '" & temp_filename & "'"
End If
objConn.Execute strSQL
rs.movenext
Loop

objConn.close
Set objConn = Nothing
Set rs = nothing
%>

