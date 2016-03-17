<%@ LANGUAGE="VBSCRIPT"%>
<HTML>
<BODY>
<%
Set uploadform = Server.CreateObject("DEXT.FileUpload")
uploadform.DefaultPath = "c:\temp"
uploadform.Save ' uploadform("file1").Save ¶ó°í ÇØµµ µÊ.
Set uploadform = Nothing
%>
</BODY>
</HTML>
