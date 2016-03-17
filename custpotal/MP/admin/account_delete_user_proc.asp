<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<%

	dim userid : userid = request("userid")

	dim objrs, sql 

	sql = "select * from dbo.wb_account where userid = '"&userid &"'"
	call set_recordset(objrs, sql)

	if not objrs.eof then 
		do until objrs.eof 
			objrs.delete()
		objrs.movenext
		loop
	end if


	sql = "select * from  dbo.wb_account_cust  where userid = '"&userid &"'"
	call set_recordset(objrs, sql)

	if not objrs.eof then 
		do until objrs.eof 
			objrs.delete()
		objrs.movenext
		loop
	end if



	sql  = "select * from  dbo.wb_account_tim  where userid = '"&userid &"'"
	call set_recordset(objrs, sql)
	
	if not objrs.eof then 
		do until objrs.eof 
			objrs.delete()
		objrs.movenext
		loop
	end if
	objrs.close

	Set objrs = Nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	
//-->
</SCRIPT>