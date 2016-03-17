<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim midx : midx = request("midx")
	dim fso : set fso = server.createobject("scripting.filesystemobject")

	dim objrs, objrs2, objrs3, objrs4, objrs5, objrs6, sql ,attachFile


	sql  = "select ref, highmenu from dbo.wb_menu_mst where midx = " & midx
	call set_recordset(objrs3, sql)

	if not objrs3.eof  then
		dim strref : strref = objrs3("ref")
		dim highmenu : highmenu = objrs3("highmenu")
	end if

	'highmenu = true  상위메뉴일때 하위메뉴를 가지고 있을때
	if highmenu  then

		sql = "select midx, title, custcode, lvl, isfile, iscomment, isemail, isuse, cuser, cdate, uuser, udate, ref, highmenu from dbo.wb_menu_mst where ref ="& strref

		call set_recordset(objrs4, sql)

		if not objrs4.eof then
			do until objrs4.eof

				sql = "select ridx from dbo.wb_report where midx = "&objrs4("midx")
				call set_recordset(objrs, sql)

				if not objrs.eof then
					do until objrs.eof
						sql = "select idx, ridx, attachfile from dbo.wb_report_pds where ridx ="&objrs("ridx")
						call set_recordset(objrs2, sql)

						if not objrs2.eof then
							do until objrs2.eof
								 attachFile = "C:\pds\file" & "\"& objrs2("attachfile")
								if fso.fileexists(attachFile) then  fso.deletefile(attachFile)
								objrs2.delete()
							objrs2.movenext
							Loop
						end if
						objrs2.close

						sql = "select cidx, ridx, comment, attachfile, cuser, cdate from dbo.wb_report_comment where ridx ="&objrs("ridx")
						call set_recordset(objrs2, sql)

						if not objrs2.eof then
							do until objrs2.eof
								 attachFile = "C:\pds\file" & "\"& objrs2("attachfile")
								if fso.fileexists(attachFile) then  fso.deletefile(attachFile)
								objrs2.delete()
							objrs2.movenext
							Loop
						end if
						objrs2.close


						objrs.delete()
					objrs.movenext
					loop
				end if
			objrs4.delete()
			objrs4.movenext
			Loop
		end if



	else  'highmenu = true else

		sql = "select ridx, attachfile from dbo.wb_report where midx = " & midx
		call set_recordset(objrs, sql)

		if not objrs.eof then
			do until objrs.eof

				sql = "select idx, ridx, attachfile from dbo.wb_report_pds where ridx ="&objrs("ridx")
				call set_recordset(objrs2, sql)

				if not objrs2.eof then
					do until objrs2.eof
						 attachFile = "C:\pds\file" & "\"& objrs2("attachfile")
						if fso.fileexists(attachFile) then  fso.deletefile(attachFile)
						objrs2.delete()
					objrs2.movenext
					Loop
				end if
				objrs2.close


				sql = "select cidx, ridx, comment, attachfile, cuser, cdate from dbo.wb_report_comment where ridx ="&objrs("ridx")
				call set_recordset(objrs2, sql)
				if not objrs2.eof then
					do until objrs2.eof
						 attachFile = "C:\pds\file" & "\"& objrs2("attachfile")
						if fso.fileexists(attachFile) then  fso.deletefile(attachFile)
						objrs2.delete()
					objrs2.movenext
					Loop
				end if
				objrs2.close

				objrs.delete()
			objrs.movenext
			loop
		end if


		'다 지우고나서 하위메뉴인놈은 삭제가될때  자기가 마지막하위메뉴였는지를 파악해서 부모메뉴의 highmenu = 0 을 업데이트쳐줘야 한다.
'		sql = "select count(*) num from dbo.wb_menu_mst where ref = " & strref & " group by highmenu"
'		call set_recordset(objrs5, sql)
'		if not objrs5.eof then
'			if objrs5("num") = 1 then
'				sql = "select highmenu from dbo.wb_menu_mst where midx = " & strref
''				response.write sql
'				response.end
'				call set_recordset(objrs6, sql)
'				objrs6.fields("highmenu").value = 0
'				objrs6.update
'				objrs6.close
'			end if
'		end if
	end if


	sql  = "select midx, title, custcode, lvl, isfile, iscomment, isemail, isuse, cuser, cdate, uuser, udate, ref, highmenu from dbo.wb_menu_mst where midx = " & midx
	call set_recordset(objrs3, sql)

	if not objrs3.eof  then
		objrs3.delete()
	end if



	Set objrs6 = Nothing
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