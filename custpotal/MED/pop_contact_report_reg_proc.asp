<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '���Ȱ��ö��̺귯�� %>

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

' ÷�����Ͽ� ��ϰ��� ���� �Ǵ�
	Dim strFileChk1
	
	If filename  = ""  Then
			Response.write "<script>"
			Response.write "alert('����� ������ ÷�ε��� �ʾҽ��ϴ�.\n\n����� ������ �����ϼ���.');"
			Response.write " window.close();"
			Response.write "</script>"
	Else
		strFileChk1 = Check_Ext(filename,"PPT,PPTX")

		If strFileChk1  = "error" Then
			Response.write "<script>"
			Response.write "alert('����� �� ���� �����Դϴ�.\n\n(PPT,PPTX)�� ����Ͻʽÿ�.');"
			Response.write " window.close();"
			Response.write "</script>"
			Response.End
		End if
	End If

	Dim attachFile : attachFile = cyear&cmonth&"_"&title&"_"&deptname&"_"&custname&"_"&medname&"_����"
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
