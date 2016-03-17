<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<%
'	Dim item
'	For Each item In request.Form
'		response.write item &  " :" & request.Form(item) & "<br>"
'	Next
'	response.end

	dim midx : midx = request("midx")

	dim objrs, sql 

	sql = "select * from dbo.wb_report_comment c inner join dbo.wb_report r on r.ridx = c.ridx where r.midx = "&midx
	call set_recordset(objrs, sql)

	if not objrs.eof then 
		do until objrs.eof 
			objrs.delete()
		objrs.movenext
		loop
	end if
	objrs.close

	sql = "select * from  dbo.wb_report  where midx = "&midx
	call set_recordset(objrs, sql)

	if not objrs.eof then 
		do until objrs.eof 
			objrs.delete()
		objrs.movenext
		loop
	end if
	objrs.close


	sql  = "select midx, title, custcode, lvl, isfile, iscomment, isemail, isuse, cuser, cdate, uuser, udate from dbo.wb_menu_mst where midx="&midx
	call set_recordset(objrs, sql)
	
	objrs.delete()

	objrs.update
	objrs.close
	Set objrs = Nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	location.href="menu_list.asp";
//-->
</SCRIPT>