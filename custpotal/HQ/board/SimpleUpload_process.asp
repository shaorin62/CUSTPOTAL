<%@ Language=VBScript %>
<HTML>
<BODY>
<%
'DEXT.FileUpload 개체 생성
set uploadform=server.CreateObject("DEXT.FileUpload")

'AutoMakeFolder 를 TRUE로 설정하면 DefaultPath, SaveAs 등등에서 지정한 폴더가 존재하지 않을 경우 폴더를 자동으로 생성한다.
'uploadform.AutoMakeFolder = True

uploadform.DefaultPath="C:\PDs"

'TempFilePath는 파일을 저장하기 전에 구해야 한다. 파일을 저장하고 나면 Temp File은 삭제된다.
Response.Write "TempFilePath: " & uploadform("file").TempFilePath & "<br>"

'Save 메소드의 첫 번째 인자는 저장될 경로다. 기본값은 DefaultPath로 지정된 폴더이다.
'Save 메소드의 두 번째 인자는 같은 파일이 존재할 경우 덮어쓸 것인지의 여부이다. 기본값은 True(파일을 덮어씀)이다.
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