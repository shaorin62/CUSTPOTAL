<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->
<%
	Dim uploadform : Set uploadform = Server.CreateObject ("DEXT.FileUpload")
	uploadform.defaultpath = "\\11.0.12.201\adportal\media"

	Dim atag : atag = ""
	Dim crud : crud = uploadform("crud")
	Dim cyear : cyear = uploadform("cyear")
	Dim cmonth : cmonth = uploadform("cmonth")
	Dim contidx : contidx = uploadform("contidx")
	Dim no : no = Trim(uploadform("no"))
	Dim file1 : file1 = uploadform("file1")
	Dim file2 : file2 = uploadform("file2")
	Dim file3 : file3 = uploadform("file3")
	Dim file4 : file4 = uploadform("file4")

	Dim sql
	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandType = adCmdText

	sql = "select * from wb_report_photo where contidx=? and cyear=? and cmonth=?"
	cmd.commandText = sql
	cmd.parameters.append cmd.createparameter("contidx", adInteger, adparaminput)
	cmd.parameters.append cmd.createparameter("cyear", adchar, adparamInput, 4)
	cmd.parameters.append cmd.createparameter("cmonth", adChar, adParaminput, 2)
	cmd.parameters("contidx").value = contidx
	cmd.parameters("cyear").value = cyear
	cmd.parameters("cmonth").value = cmonth
	Dim rs : Set rs = cmd.execute
	clearparameter(cmd)

	If Not rs.eof Then
		If no = "" Then crud = "U"
		Dim orgphoto1 : orgphoto1 = rs("photo1")
		Dim orgphoto2 : orgphoto2 = rs("photo2")
		Dim orgphoto3 : orgphoto3 = rs("photo3")
		Dim orgphoto4 : orgphoto4 = rs("photo4")
	Else
		crud = "C"
	End If

	Select Case crud
	Case "C"
		sql = "insert into wb_report_photo(contidx, cyear, cmonth, photo1, photo2, photo3, photo4) values (?, ?, ?, ? ,? ,?, ?)"
		cmd.commandText = sql
		cmd.parameters.append cmd.createparameter("contidx", adInteger, adParaminput)
		cmd.parameters.append cmd.createparameter("cyear", adChar, adParaminput, 4)
		cmd.parameters.append cmd.createparameter("cmonth", adChar, adParaminput, 2)
		cmd.parameters.append cmd.createparameter("photo1", adVarChar, adParaminput, 200)
		cmd.parameters.append cmd.createparameter("photo2", adVarChar, adParaminput, 200)
		cmd.parameters.append cmd.createparameter("photo3", adVarChar, adParaminput, 200)
		cmd.parameters.append cmd.createparameter("photo4", adVarChar, adParaminput, 200)

		Dim filename : filename = cyear&cmonth&"_"&contidx
		If file1 = "" Then
			Dim filename1 : filename1 = Null
		Else
			Dim fileExt1 : fileExt1 = uploadform("file1").FileExtension
			filename1 = filename&"_01."&fileExt1
			uploadform("file1").SaveAs(uploadform.defaultpath&"\"&filename1)
		End If
		If file2 = "" Then
			Dim filename2 : filename2 = Null
		Else
			Dim fileExt2 : fileExt2 = uploadform("file2").FileExtension
			filename2 = filename&"_02."&fileExt2
			uploadform("file2").SaveAs(uploadform.defaultpath&"\"&filename2)
		End If
		If file3 = "" Then
			Dim filename3 : filename3 = Null
		Else
			Dim fileExt3 : fileExt3 = uploadform("file3").FileExtension
			filename3 = filename&"_03."&fileExt3
			uploadform("file3").SaveAs(uploadform.defaultpath&"\"&filename3)
		End If
		If file4 = "" Then
			Dim filename4 : filename4 = Null
		Else
			Dim fileExt4 : fileExt4 = uploadform("file4").FileExtension
			filename4 = filename&"_04."&fileExt4
			uploadform("file4").SaveAs(uploadform.defaultpath&"\"&filename4)
		End If

		cmd.parameters("contidx").value = contidx
		cmd.parameters("cyear").value = cyear
		cmd.parameters("cmonth").value = cmonth
		cmd.parameters("photo1").value = filename1
		cmd.parameters("photo2").value = filename2
		cmd.parameters("photo3").value = filename3
		cmd.parameters("photo4").value = filename4

		cmd.execute ,, adExecuteNoRecords

	Case "U"
		sql = "update wb_report_photo set photo1=?, photo2=?, photo3=?, photo4=? where contidx=? and cyear=? and cmonth=?"
		cmd.commandText = sql
		cmd.parameters.append cmd.createparameter("photo1", adVarChar, adParaminput, 200)
		cmd.parameters.append cmd.createparameter("photo2", adVarChar, adParaminput, 200)
		cmd.parameters.append cmd.createparameter("photo3", adVarChar, adParaminput, 200)
		cmd.parameters.append cmd.createparameter("photo4", adVarChar, adParaminput, 200)
		cmd.parameters.append cmd.createparameter("contidx", adInteger, adParaminput)
		cmd.parameters.append cmd.createparameter("cyear", adChar, adParaminput, 4)
		cmd.parameters.append cmd.createparameter("cmonth", adChar, adParaminput, 2)

		filename = cyear&cmonth&"_"&contidx
		If file1 = "" Then
			filename1 = orgphoto1
		Else
			fileExt1 = uploadform("file1").FileExtension
			filename1 = filename&"_01."&fileExt1
			If uploadform.FileExists(uploadform.defaultpath&"\"&filename1) Then uploadform.DeleteFile(uploadform.defaultpath&"\"&filename1)
			uploadform("file1").SaveAs(uploadform.defaultpath&"\"&filename1)
		End If
		If file2 = "" Then
			filename2 = orgphoto2
		Else
			fileExt2 = uploadform("file2").FileExtension
			filename2 = filename&"_02."&fileExt2
			If uploadform.FileExists(uploadform.defaultpath&"\"&filename2) Then uploadform.DeleteFile(uploadform.defaultpath&"\"&filename2)
			uploadform("file2").SaveAs(uploadform.defaultpath&"\"&filename2)
		End If
		If file3 = "" Then
			filename3 = orgphoto3
		Else
			fileExt3 = uploadform("file3").FileExtension
			filename3 = filename&"_03."&fileExt3
			If uploadform.FileExists(uploadform.defaultpath&"\"&filename3) Then uploadform.DeleteFile(uploadform.defaultpath&"\"&filename3)
			uploadform("file3").SaveAs(uploadform.defaultpath&"\"&filename3)
		End If
		If file4 = "" Then
			filename4 = orgphoto4
		Else
			fileExt4 = uploadform("file4").FileExtension
			filename4 = filename&"_04."&fileExt4
			If uploadform.FileExists(uploadform.defaultpath&"\"&filename4) Then uploadform.DeleteFile(uploadform.defaultpath&"\"&filename4)
			uploadform("file4").SaveAs(uploadform.defaultpath&"\"&filename4)
		End If

		cmd.parameters("photo1").value = filename1
		cmd.parameters("photo2").value = filename2
		cmd.parameters("photo3").value = filename3
		cmd.parameters("photo4").value = filename4
		cmd.parameters("contidx").value = contidx
		cmd.parameters("cyear").value = cyear
		cmd.parameters("cmonth").value = cmonth

		cmd.execute ,, adExecuteNoRecords

	Case "D"
		sql = "update wb_report_photo set photo1=?, photo2=?, photo3=?, photo4=? where contidx=? and cyear=? and cmonth=?"
		cmd.commandText = sql
		cmd.parameters.append cmd.createparameter("photo1", adVarChar, adParaminput, 200)
		cmd.parameters.append cmd.createparameter("photo2", adVarChar, adParaminput, 200)
		cmd.parameters.append cmd.createparameter("photo3", adVarChar, adParaminput, 200)
		cmd.parameters.append cmd.createparameter("photo4", adVarChar, adParaminput, 200)
		cmd.parameters.append cmd.createparameter("contidx", adInteger, adParaminput)
		cmd.parameters.append cmd.createparameter("cyear", adChar, adParaminput, 4)
		cmd.parameters.append cmd.createparameter("cmonth", adChar, adParaminput, 2)
		If no = "photo1" Then
			If uploadform.FileExists(uploadform.defaultpath&"\"&orgphoto1) Then uploadform.DeleteFile(uploadform.defaultpath&"\"&orgphoto1)
			filename1=Null
		Else
			filename1=orgphoto1
		End If
		If no = "photo2" Then
			If uploadform.FileExists(uploadform.defaultpath&"\"&orgphoto2) Then uploadform.DeleteFile(uploadform.defaultpath&"\"&orgphoto2)
			filename2=Null
		Else
			filename2=orgphoto2
		End If
		If no = "photo3" Then
			If uploadform.FileExists(uploadform.defaultpath&"\"&orgphoto3) Then uploadform.DeleteFile(uploadform.defaultpath&"\"&orgphoto3)
			filename3=Null
		Else
			filename3=orgphoto3
		End If
		If no = "photo4" Then
			If uploadform.FileExists(uploadform.defaultpath&"\"&orgphoto4) Then uploadform.DeleteFile(uploadform.defaultpath&"\"&orgphoto4)
			filename4=Null
		Else
			filename4=orgphoto4
		End If
		cmd.parameters("photo1").value = filename1
		cmd.parameters("photo2").value = filename2
		cmd.parameters("photo3").value = filename3
		cmd.parameters("photo4").value = filename4
		cmd.parameters("contidx").value = contidx
		cmd.parameters("cyear").value = cyear
		cmd.parameters("cmonth").value = cmonth
		cmd.execute ,, adExecuteNoRecords
		clearparameter(cmd)

		sql = "select photo1, photo2, photo3, photo4 from wb_report_photo where contidx=? and cyear=? and cmonth=?"
		cmd.commandText = sql
		cmd.parameters.append cmd.createparameter("contidx", adInteger, adParaminput)
		cmd.parameters.append cmd.createparameter("cyear", adChar, adParaminput, 4)
		cmd.parameters.append cmd.createparameter("cmonth", adChar, adParaminput, 2)
		cmd.parameters("contidx").value = contidx
		cmd.parameters("cyear").value = cyear
		cmd.parameters("cmonth").value = cmonth

		Set rs = cmd.execute
		If Not rs.eof Then
			If IsNull(rs(0)) And IsNull(rs(1)) And IsNull(rs(2)) And IsNull(rs(3)) Then
				sql = "delete from wb_report_photo where contidx=? and cyear=? and cmonth=?"
				cmd.commandText = sql
				cmd.execute ,, adExecuteNoRecords
			End If
		End If

	End Select

	response.write "no  : " & no & "<br>"
	response.write "crud  : " & crud & "<br>"
	response.write "filename1  : " & filename1 & "<br>"
	response.write "filename2  : " & filename2 & "<br>"
	response.write "filename3  : " & filename3 & "<br>"
	response.write "filename4  : " & filename4 & "<br>"

'
'	If Err.number <> 0 Then
'		Call Debug
'	End If
%>
<script type="text/javascript">
<!--
	parent.document.location.reload();
//-->
</script>