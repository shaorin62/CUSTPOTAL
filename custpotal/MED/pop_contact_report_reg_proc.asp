<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '보안관련라이브러리 %>

<%
dim uploadform : set uploadform = server.createobject("DEXT.FileUpload")
uploadform.defaultpath = "C:\CUSTPOTAL\report"

dim filename : filename = uploadform("file1")
Dim title : title = uploadform("title")
Dim contidx : contidx = uploadform("contidx")
Dim cyear : cyear = uploadform("cyear")
Dim cmonth : cmonth = uploadform("cmonth")
Dim custname : custname = uploadform("custname")
Dim deptname : deptname = uploadform("deptname")
Dim medname : medname = request.cookies("custname")
Dim objrs

' 첨부파일에 등록가능 여부 판단
	Dim strFileChk1
	
	If filename  = ""  Then
			Response.write "<script>"
			Response.write "alert('등록할 파일이 첨부되지 않았습니다.\n\n등록할 보고서를 선택하세요.');"
			Response.write " window.close();"
			Response.write "</script>"
	Else
		strFileChk1 = Check_Ext(filename,"PPT,PPTX")

		If strFileChk1  = "error" Then
			Response.write "<script>"
			Response.write "alert('등록할 수 없는 파일입니다.\n\n(PPT,PPTX)만 등록하십시오.');"
			Response.write " window.close();"
			Response.write "</script>"
			Response.End
		End if
	End If

	Dim attachFile : attachFile = cyear&cmonth&"_"&title&"_"&deptname&"_"&custname&"_"&medname&"_보고서"
	Dim exe : exe = right(filename,len(filename)-instr(filename,"."))
	Dim fullpath : fullpath = uploadform.defaultpath &"\"&attachFile&"."&exe

	Dim sql : sql = "select  contidx, cyear, cmonth , custcode, reportname, cuser, cdate from dbo.wb_contact_report where contidx = " & contidx
	Call  set_recordset(objrs, sql)
	
	dim fso : set fso = server.createobject("scripting.filesystemobject")
	if fso.fileexists(fullpath) then 
		fso.deletefile(fullpath)
		objrs.delete()
	end if

  'filename = uploadform("file1").save()
	uploadform.SaveAs(fullpath) 

	objrs.addnew 
	objrs("contidx") = contidx
	objrs("cyear") = cyear
	objrs("cmonth") = cmonth
	objrs("custcode") = medname
	objrs("reportname") =  attachFile&"."&exe
	objrs("cuser") = request.cookies("userid")
	objrs("cdate") = Date
	objrs.update
	objrs.close
	Set objrs = nothing

	If Err Then
		Response.Write Err.number & "<br>" & Err.source & "<br>" &  Err.description
		Set uploadform = Nothing
		Response.End
	End if
%>

<script language="JavaScript">
<!--
	opener.location.reload();
	this.close();
//-->
</script>
