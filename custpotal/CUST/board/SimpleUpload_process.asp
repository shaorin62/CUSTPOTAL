<%@ Language=VBScript %>
<HTML>
<BODY>
<%
'DEXT.FileUpload ��ü ����
set uploadform=server.CreateObject("DEXT.FileUpload")

'AutoMakeFolder �� TRUE�� �����ϸ� DefaultPath, SaveAs ���� ������ ������ �������� ���� ��� ������ �ڵ����� �����Ѵ�.
'uploadform.AutoMakeFolder = True

uploadform.DefaultPath="C:\PDs"

'TempFilePath�� ������ �����ϱ� ���� ���ؾ� �Ѵ�. ������ �����ϰ� ���� Temp File�� �����ȴ�.
Response.Write "TempFilePath: " & uploadform("file").TempFilePath & "<br>"

'Save �޼ҵ��� ù ��° ���ڴ� ����� ��δ�. �⺻���� DefaultPath�� ������ �����̴�.
'Save �޼ҵ��� �� ��° ���ڴ� ���� ������ ������ ��� ��� �������� �����̴�. �⺻���� True(������ ���)�̴�.
FilePath = uploadform("file").Save
Response.Write "Original Path : " & uploadform("file").FilePath  & "<br>"
Response.Write "Upload Path : " & FilePath & "<br>"
Response.Write "File Size : " & uploadform("file").FileLen & "<br>"
Response.Write "MimeType : " & uploadform("file").MimeType & "<br>"
Response.Write "LastSavedFileName : " & uploadform("file").LastSavedFileName & "<br>"
Response.Write "LastSavedFilePath : " & uploadform("file").LastSavedFilePath & "<br>"
Response.Write "FileNameWithoutExt : " & uploadform("file").FileNameWithoutExt & "<br>"
Response.Write "FileExtension : " & uploadform("file").FileExtension & "<br>"
%>

</BODY>
</HTML>
<%
Set uploadform =nothing
%>