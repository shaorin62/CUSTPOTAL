<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim midx : midx = request("midx")
	dim fso : set fso = server.createobject("scripting.filesystemobject")

	dim objrs, objrs2, objrs3, objrs4, objrs5, objrs6, sql ,attachFile

	sql = "select midx, title, custcode, lvl, isfile, iscomment, isemail, isuse, cuser, cdate, uuser, udate, ref, highmenu from dbo.wb_menu_mst where midx="& midx &" or  ref ="& midx
	call set_recordset(objrs, sql)

	if not objrs.eof then
		do until objrs.eof
			
			sql = "select midx, title, custcode, lvl, isfile, iscomment, isemail, isuse, cuser, cdate, uuser, udate, ref, highmenu from dbo.wb_menu_mst where midx ="& midx  &" or  ref ="& objrs("midx") 
			response.write sql
			call set_recordset(objrs2, sql)

			if not objrs2.eof then
				do until objrs2.eof
					sql = "select ridx from dbo.wb_report where midx = "&objrs2("midx")
					call set_recordset(objrs3, sql)

					if not objrs3.eof then
						do until objrs3.eof
							sql = "select idx, ridx, attachfile from dbo.wb_report_pds where ridx ="&objrs3("ridx")
							call set_recordset(objrs4, sql)

							if not objrs4.eof then
								do until objrs4.eof
									attachFile = "C:\pds\file" & "\"& objrs4("attachfile")
									if fso.fileexists(attachFile) then  fso.deletefile(attachFile)
									objrs4.delete()
									objrs4.movenext
								Loop
							End if
							objrs4.close

							sql = "select cidx, ridx, comment, attachfile, cuser, cdate from dbo.wb_report_comment where ridx ="&objrs3("ridx")
							call set_recordset(objrs4, sql)

							if not objrs4.eof then
								do until objrs4.eof
									attachFile = "C:\pds\file" & "\"& objrs4("attachfile")
									if fso.fileexists(attachFile) then  fso.deletefile(attachFile)
									objrs4.delete()
									objrs4.movenext
								Loop
							End if
							objrs4.close

							objrs3.delete()
							objrs3.movenext
						Loop
					End if
					objrs2.delete()
					objrs2.movenext
				Loop
			End If
			'objrs.delete()
			objrs.movenext
		Loop
	End if

	sql  = "select midx, title, custcode, lvl, isfile, iscomment, isemail, isuse, cuser, cdate, uuser, udate, ref, highmenu from dbo.wb_menu_mst where midx = " & midx
	call set_recordset(objrs5, sql)

	if not objrs5.eof  then
		objrs5.delete()
	end if

	Set objrs5 = Nothing
	Set objrs4 = Nothing
	Set objrs3 = Nothing
	Set objrs2 = Nothing
	Set objrs = Nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	//window.opener.location.reload();
	window.opener.document.location.href = window.opener.document.URL;
	this.close();
//-->
</SCRIPT>