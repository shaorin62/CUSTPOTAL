<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->
<%
	On Error Resume Next

	Dim uploadform : Set uploadform = Server.CreateObject ("DEXT.FileUpload")
	uploadform.defaultpath = "C:\pds\media"


	Dim atag : atag = ""
	Dim pcyear : pcyear = clearXSS(uploadform("cyear"), atag)
	Dim pcmonth : pcmonth = clearXSS(uploadform("cmonth"), atag)
	Dim pmdidx : pmdidx = clearXSS(uploadform("mdidx"), atag)
	Dim pside : pside = clearXSS(uploadform("side"), atag)
	Dim pseq : pseq = clearXSS(uploadform("seq"), atag)
	Dim pcol : pcol = clearXSS(uploadform("col"), atag)
	Dim plastdate : plastdate = clearXSS(uploadform("lastdate"), atag)
	Dim pcrud : pcrud = clearXSS(uploadform("crud"), atag)


	Dim sql_
	Dim attachFile_

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandType = adCmdText

	Select Case UCase(pcrud)
		Case "C"
			cmd.parameters.append cmd.createparameter("cyear", adChar, adParamInput, 4)
			cmd.parameters.append cmd.createparameter("cmonth", adChar, adParamInput, 2)
			cmd.parameters.append cmd.createparameter("mdidx", adInteger, adParamInput)
			cmd.parameters.append cmd.createparameter("side", adChar, adParamInput, 1)
			cmd.parameters.append cmd.createparameter("pht_01", adVarChar, adParamInput, 11)
			cmd.parameters.append cmd.createparameter("pht_02", adVarChar, adParamInput, 11)
			cmd.parameters.append cmd.createparameter("pht_03", adVarChar, adParamInput, 11)
			cmd.parameters.append cmd.createparameter("pht_04", adVarChar, adParamInput, 11)
			cmd.parameters.append cmd.createparameter("desc_01", adVarChar, adParamInput, 100)
			cmd.parameters.append cmd.createparameter("desc_02", adVarChar, adParamInput, 100)
			cmd.parameters.append cmd.createparameter("desc_03", adVarChar, adParamInput, 100)
			cmd.parameters.append cmd.createparameter("desc_04", adVarChar, adParamInput, 100)
			cmd.parameters.append cmd.createparameter("startdate", adDBTimeStamp, adParamInput)
			cmd.parameters.append cmd.createparameter("enddate", adDBTimeStamp, adParamInput)
			'cmd.parameters.append cmd.createparameter("seq", adInteger, adParamInput)

			If Len(uploadform("file01")) Then
				attachFile_ = uploadform("file01").save (, false)
				pfile01 = Right(attachFile_, Len(attachFile_)-InstrRev(attachFile_,"\"))
			Else
				pfile01 = Null
			End If
			If Len(uploadform("file02")) Then
				attachFile_ = uploadform("file02").save (, false)
				pfile02 = Right(attachFile_, Len(attachFile_)-InstrRev(attachFile_,"\"))
			Else
				pfile02 = Null
			End If
			If Len(uploadform("file03")) Then
				attachFile_ = uploadform("file03").save (, false)
				pfile03 = Right(attachFile_, Len(attachFile_)-InstrRev(attachFile_,"\"))
			Else
				pfile03 = Null
			End If
			If Len(uploadform("file04")) Then
				attachFile_ = uploadform("file04").save (, false)
				pfile04 = Right(attachFile_, Len(attachFile_)-InstrRev(attachFile_,"\"))
			Else
				pfile04 = Null
			End If

			sql_="insert into wb_contact_photo (cyear, cmonth, mdidx, side, pht_01, pht_02, pht_03, pht_04, desc_01, desc_02, desc_03, desc_04, startdate, enddate) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

			cmd.commandText = sql_
			cmd.parameters("cyear").value = pcyear
			cmd.parameters("cmonth").value = pcmonth
			cmd.parameters("mdidx").value = pmdidx
			cmd.parameters("side").value = pside
			cmd.parameters("pht_01").value = null
			cmd.parameters("pht_02").value = null
			cmd.parameters("pht_03").value = null
			cmd.parameters("pht_04").value = null
			cmd.parameters("desc_01").value = pfile01
			cmd.parameters("desc_02").value = pfile02
			cmd.parameters("desc_03").value = pfile03
			cmd.parameters("desc_04").value = pfile04
			cmd.parameters("startdate").value = null
			cmd.parameters("enddate").value = Null


			cmd.execute ,, adExecuteNoRecords

		Case "U"
			sql_="select " & pcol &"  from wb_contact_photo where seq = ?"
			cmd.commandText = sql_
			cmd.parameters.append cmd.createparameter("seq", adInteger, adParamInput)
			cmd.parameters("seq").value = pseq
			Dim rs_ : Set rs_ = cmd.execute

			If Len(uploadform("file05")) Then
				uploadform.deleteFile(uploadform.defaultPath&"\"&rs_(0))
				attachFile_ = uploadform("file05").save (, false)
				pfile01 = Right(attachFile_, Len(attachFile_)-InstrRev(attachFile_,"\"))
			Else
				pfile01 = rs_(0)
			End If

			cmd.parameters.delete("seq")
			sql_ = "update wb_contact_photo set "& pcol &"=? where seq = ?"
			cmd.commandText = sql_
			cmd.parameters.append cmd.createparameter("desc", adVarChar, adParamInput, 100)
			cmd.parameters.append cmd.createparameter("seq", adInteger, adParamInput)
			cmd.parameters("desc").value = pfile01
			cmd.parameters("seq").value = pseq
			cmd.execute ,, adExecuteNoRecords

		Case "D"
			sql_ = "select desc_01, desc_02, desc_03, desc_04 from wb_contact_photo where seq = ?"
			cmd.commandText = sql_
			cmd.parameters.append cmd.createparameter("seq", adInteger, adParamInput)
			cmd.parameters("seq").value = pseq
			Set rs_ = cmd.execute

			Select Case pcol
				Case "desc_01"
					If IsNull(rs_(1)) And IsNull(rs_(2)) And IsNull(rs_(3)) Then
						sql_ ="delete from wb_contact_photo where seq=?"
					Else
						sql_ = "update wb_contact_photo set "& pcol &"=null where seq = ?"
					End If
					uploadform.deleteFile(uploadform.defaultpath&"\"&rs_(0))
				Case "desc_02"
					If IsNull(rs_(0)) And IsNull(rs_(2)) And IsNull(rs_(3)) Then
						sql_ ="delete from wb_contact_photo where seq=?"
					Else
						sql_ = "update wb_contact_photo set "& pcol &"=null where seq = ?"
					End If
					uploadform.deleteFile(uploadform.defaultpath&"\"&rs_(1))
				Case "desc_03"
					If IsNull(rs_(0)) And IsNull(rs_(1)) And IsNull(rs_(3)) Then
						sql_ ="delete from wb_contact_photo where seq=?"
					Else
						sql_ = "update wb_contact_photo set "& pcol &"=null where seq = ?"
					End If
					uploadform.deleteFile(uploadform.defaultpath&"\"&rs_(2))
				Case "desc_04"
					If IsNull(rs_(0)) And IsNull(rs_(1)) And IsNull(rs_(2)) Then
						sql_ ="delete from wb_contact_photo where seq=?"
					Else
						sql_ = "update wb_contact_photo set "& pcol &"=null where seq = ?"
					End If
					uploadform.deleteFile(uploadform.defaultpath&"\"&rs_(3))
			End Select
			cmd.commandText = sql_
			cmd.execute ,, adExecuteNoRecords
	End Select
	Set cmd = Nothing

	If Err.number <> 0 Then
		Call Debug
	End If
%>
<script type="text/javascript">
<!--
//	location.href="/hq/outdoor/popup/view_photo.asp?mdidx=<%=pmdidx%>&side=<%=pside%>&lastdate=<%=plastdate%>";

	window.opener.getcontactphoto();
	window.close();
//-->
</script>