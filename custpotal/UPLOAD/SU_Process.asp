<%@ LANGUAGE="VBSCRIPT"%>
<HTML>
<BODY>
<%
Set uploadform = Server.CreateObject("DEXT.FileUpload")
uploadform.DefaultPath = "c:\temp"
uploadform.Save ' uploadform("file1").Save ��� �ص� ��.
Set uploadform = Nothing
%>
</BODY>
</HTML>
