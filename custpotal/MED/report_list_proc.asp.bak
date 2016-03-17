<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '보안관련라이브러리 %>

<%
dim uploadform : set uploadform = server.createobject("DEXT.FileUpload")
uploadform.defaultpath = "C:\CUSTPOTAL\report"

dim filename : filename = uploadform("file")

' 첨부파일에 등록가능 여부 판단
	Dim strFileChk1
	
	If filename  = ""  Then
		filename =  Null
	Else
		strFileChk1 = Check_Ext(filename,"PPT,PPTX")

		If strFileChk1  = "error" Then
			Response.write "<script>"
			Response.write "alert('등록할 수 없는 파일입니다.\n\n(PPT,PPTX)만 등록하십시오.');"
			Response.write " this.close();"
			Response.write "</script>"
			Response.End
		End if
	End If

   filename = uploadform("file").save()

	If Err Then
		Response.Write Err.number & "<br>" & Err.source & "<br>" &  Err.description
		Set uploadform = Nothing
		Response.End
	End if
%>

<script language="JavaScript">
<!--
	this.close();
//-->
</script>
