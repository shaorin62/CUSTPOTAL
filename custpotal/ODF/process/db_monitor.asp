<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->
<div style='width:530px; height:400px;z-index:10;top:0px;left:0px;background-color:#ffffff;margin:0 0 0 0 ;'>
<table width='100%' height='100%' border=0>
	<tr>
		<td align='center'><img src='/images/uploading.gif'><p> Please Wait <br> Images Uploading ...</td>
	</tr>
</table>
</div>
<%
'	On Error Resume Next
	If request.cookies("userid") = "" Then
		response.write "<script> try {this.close();} catch(e) {window.close();} </script>"
		response.end
	end if

	Dim uploadform : Set uploadform = Server.CreateObject ("DEXT.FileUpload")
	uploadform.defaultpath = "\\11.0.12.201\adportal\monitor"
	Dim uploadPath : uploadPath = uploadform.defaultpath

	Dim atag : atag = ""
	Dim pcrud : pcrud = clearXSS(uploadform("crud"), atag)
	Dim pcyear : pcyear = clearXSS(uploadform("cyear"), atag)
	Dim pcmonth : pcmonth = clearXSS(uploadform("cmonth"), atag)
	Dim pmdidx : pmdidx = clearXSS(uploadform("mdidx"), atag)
	Dim pside : pside = clearXSS(uploadform("side"), atag)
	Dim pcdate : pcdate = clearXSS(uploadform("txtcdate"), atag)
	Dim pnum : pnum = clearXSS(uploadform("selnum"), atag)
	Dim pstatus : pstatus = clearXSS(uploadform("rdostatus"), atag)
	Dim pcname : pcname = clearXSS(uploadform("txtcname"), atag)
	Dim pcomment : pcomment = clearXSS(uploadform("txtcomment"), atag)
	Dim orgnum : orgnum = clearXSS(uploadform("orgnum"), atag)
	Dim pcustcode : pcustcode = clearXSS(uploadform("custcode"), atag)
	Dim pteamcode : pteamcode = clearXSS(uploadform("teamcode"), atag)


	Dim filename
	Dim img(4)

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandType = adCmdText

	Select Case UCase(pcrud)
		Case "C"
			pcdate = Left(pcdate,4)&"-"&Mid(pcdate,5,2)&"-"&Right(pcdate,2)
			Dim sql : sql = "insert into wb_contact_monitor(mdidx, cyear, cmonth, num, side, status, comment, cdate, cuser, cname, img01, img02, img03, img04) values (?,?,?,?,?,?,?,?,?,?,?,?,?,?) "
			cmd.parameters.append	cmd.createparameter("mdidx", adinteger, adParaminput)
			cmd.parameters.append	cmd.createparameter("cyear", adchar, adParaminput, 4)
			cmd.parameters.append	cmd.createparameter("cmonth", adchar, adParaminput, 2)
			cmd.parameters.append	cmd.createparameter("num", adUnsignedTinyInt, adParaminput)
			cmd.parameters.append	cmd.createparameter("side", adchar, adParaminput, 1)
			cmd.parameters.append	cmd.createparameter("status", adchar, adParaminput, 1) 'Y :양호, N:불량
			cmd.parameters.append	cmd.createparameter("comment", advarchar, adParaminput, 1000)
			cmd.parameters.append	cmd.createparameter("cdate", adDBTimeStamp, adParaminput)
			cmd.parameters.append	cmd.createparameter("cuser", advarchar, adParaminput, 12)
			cmd.parameters.append	cmd.createparameter("cname", advarchar, adParaminput, 20)
			cmd.parameters.append	cmd.createparameter("img01", advarchar, adParaminput, 200)
			cmd.parameters.append	cmd.createparameter("img02", advarchar, adParaminput, 200)
			cmd.parameters.append	cmd.createparameter("img03", advarchar, adParaminput, 200)
			cmd.parameters.append	cmd.createparameter("img04", advarchar, adParaminput, 200)

			cmd.parameters("mdidx").value = pmdidx
			cmd.parameters("cyear").value = pcyear
			cmd.parameters("cmonth").value = pcmonth
			cmd.parameters("num").value = pnum
			cmd.parameters("side").value = pside
			cmd.parameters("status").value = pstatus
			cmd.parameters("comment").value = pcomment
			cmd.parameters("cdate").value = pcdate
			cmd.parameters("cuser").value = request.cookies("userid")
			cmd.parameters("cname").value = pcname
			For intLoop = 1 To uploadform("file").count
				filename = uploadform("file")(intLoop).saveAs (,false)
				filename = Right(filename, Len(filename)-InstrRev(filename,"\"))
				cmd.parameters("img0"&intLoop).value = filename
			Next
			cmd.commandText = sql
			cmd.execute ,, adExecutenorecords
		Case "U"
			If Len(pcdate) = 8 Then pcdate = Left(pcdate,4)&"-"&Mid(pcdate,5,2)&"-"&Right(pcdate,2)
			cmd.parameters.append	cmd.createparameter("num", adUnsignedTinyInt, adParaminput)
			cmd.parameters.append	cmd.createparameter("status", adchar, adParaminput, 1) 'Y :양호, N:불량
			cmd.parameters.append	cmd.createparameter("comment", advarchar, adParaminput, 1000)
			cmd.parameters.append	cmd.createparameter("cdate", adDBTimeStamp, adParaminput)
			cmd.parameters.append	cmd.createparameter("cuser", advarchar, adParaminput, 12)
			cmd.parameters.append	cmd.createparameter("cname", advarchar, adParaminput, 20)
			cmd.parameters.append	cmd.createparameter("img01", advarchar, adParaminput, 200)
			cmd.parameters.append	cmd.createparameter("img02", advarchar, adParaminput, 200)
			cmd.parameters.append	cmd.createparameter("img03", advarchar, adParaminput, 200)
			cmd.parameters.append	cmd.createparameter("img04", advarchar, adParaminput, 200)
			cmd.parameters.append	cmd.createparameter("mdidx", adinteger, adParaminput)
			cmd.parameters.append	cmd.createparameter("side", adchar, adParaminput, 1)
			cmd.parameters.append	cmd.createparameter("orgnum", adUnsignedTinyInt, adParaminput)
			cmd.parameters.append	cmd.createparameter("cyear", adchar, adParaminput, 4)
			cmd.parameters.append	cmd.createparameter("cmonth", adchar, adParaminput, 2)
			sql = "update wb_contact_monitor set num=?, status=?, comment=?, cdate=?, cuser=?, cname=?,img01=?, img02=?, img03=?, img04=? where mdidx=? and side=? and num=? and cyear=? and cmonth=?"
			cmd.commandText = sql
			cmd.parameters("num").value = pnum
			cmd.parameters("status").value = pstatus
			cmd.parameters("comment").value = pcomment
			cmd.parameters("cdate").value = pcdate
			cmd.parameters("cuser").value = request.cookies("userid")
			cmd.parameters("cname").value = pcname

			For intLoop =1 To 4
				If Len(uploadform("orgfile")(intLoop)) Then
					If uploadform("txtfile")(intLoop) = "" Then
						If uploadform.fileExists(uploadPath&"\"&uploadform("orgfile")(intLoop)) Then	uploadform.deletefile(uploadPath&"\"&uploadform("orgfile")(intLoop))
						filename = Null
					ElseIf Len(uploadform("orgfile")(intLoop)) <> Len(uploadform("txtfile")(intLoop)) Then
						If uploadform.fileExists(uploadPath&"\"&uploadform("orgfile")(intLoop)) Then	uploadform.deletefile(uploadPath&"\"&uploadform("orgfile")(intLoop))
						filename = uploadform("file")(intLoop).saveAs (,false)
						filename = Right(filename, Len(filename)-InstrRev(filename,"\"))
					Else
						filename = uploadform("orgfile")(intLoop)
					End If
				Else
					If Len(uploadform("file")(intLoop)) Then
						filename = uploadform("file")(intLoop).saveAs (,false)
						filename = Right(filename, Len(filename)-InstrRev(filename,"\"))
					Else
						filename = Null
					End If
				End If
			img(intLoop) = filename
			Next
			cmd.parameters("img01").value = img(1)
			cmd.parameters("img02").value = img(2)
			cmd.parameters("img03").value = img(3)
			cmd.parameters("img04").value = img(4)

			cmd.parameters("mdidx").value = pmdidx
			cmd.parameters("side").value = pside
			cmd.parameters("orgnum").value = orgnum
			cmd.parameters("cyear").value = pcyear
			cmd.parameters("cmonth").value = pcmonth
			cmd.commandText = sql
			cmd.execute ,, adExecutenorecords


		Case "D"
			For intLoop = 1 To uploadform("orgfile").count
				If uploadform.fileExists(uploadPath&"\"&uploadform("orgfile")(intLoop)) Then	uploadform.deletefile(uploadPath&"\"&uploadform("orgfile")(intLoop))
			Next

			cmd.parameters.append	cmd.createparameter("mdidx", adinteger, adParaminput)
			cmd.parameters.append	cmd.createparameter("side", adchar, adParaminput, 1)
			cmd.parameters.append	cmd.createparameter("num", adUnsignedTinyInt, adParaminput)
			cmd.parameters.append	cmd.createparameter("cyear", adchar, adParaminput, 4)
			cmd.parameters.append	cmd.createparameter("cmonth", adchar, adParaminput, 2)
			sql = "delete from wb_contact_monitor where mdidx=? and side=? and num=? and cyear=? and cmonth=?"
			cmd.parameters("mdidx").value = pmdidx
			cmd.parameters("side").value = pside
			cmd.parameters("num").value = pnum
			cmd.parameters("cyear").value = pcyear
			cmd.parameters("cmonth").value = pcmonth
			cmd.commandText = sql
			cmd.execute ,, adExecutenorecords
			pnum = pnum-1

	End Select

	If Err.number <> 0 Then
		Call Debug
	End If
%>
<script type="text/javascript">
<!--
	window.opener.location.href ='/odf/detail_monitor.asp?cyear=<%=pcyear%>&cmonth=<%=pcmonth%>&mdidx=<%=pmdidx%>&side=<%=pside%>&num=<%=pnum%>&custcode=<%=pcustcode%>&teamcode=<%=pteamcode%>';
	window.close();
//-->
</script>