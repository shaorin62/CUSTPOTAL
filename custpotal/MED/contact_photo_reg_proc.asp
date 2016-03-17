<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim objrs,  sql, intLoop
	dim filename(3)


	sql = "select filename from dbo.wb_contact_photo_dtl where idx in (" & request("photoIdx") & ")"
	call get_recordset(objrs, sql)

	intLoop = 0

	do until objrs.eof
		filename(intLoop) = objrs("filename")
		intLoop = intLoop + 1
		objrs.movenext
	Loop
	objrs.close


	if filename(0) = "" then filename(0) = null
	if filename(1) = "" then filename(1) = null
	if filename(2) = "" then filename(2) = null
	if filename(3) = "" then filename(3) = null


	sql = "select photo_1, photo_2, photo_3, photo_4 from dbo.wb_contact_md_dtl_account where idx = "& request("idx") &" and cyear = '" & request("cyear") & "'  and cmonth = '" & request("cmonth") & "' "
	call set_recordset(objrs, sql)

	if not isnull(filename(0)) or filename(0) <> ""  then	objrs("photo_1") = filename(0) 	else  	objrs("photo_1") = null 	end if
	if not isnull(filename(1))  or filename(1) <> "" then	objrs("photo_2") = filename(1) 	else  	objrs("photo_2") = null 	end if
	if not isnull(filename(2))  or filename(2) <> "" then	objrs("photo_3") = filename(2) 	else  	objrs("photo_3") = null 	end if
	if not isnull(filename(3))  or filename(3) <> "" then	objrs("photo_4") = filename(3) 	else  	objrs("photo_4") = null 	end if

	objrs.update
	objrs.close
	set objrs = nothing
	response.redirect "pop_contact_photo_reg.asp?idx="& request("idx") & "&cyear=" & request("cyear") & "&cmonth=" & request("cmonth")
%>