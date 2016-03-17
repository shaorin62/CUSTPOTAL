<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<%

	dim userid : userid = request("userid")
	dim custcode : custcode = request("custcode")
	dim timcode : timcode = request("timcode")

	

	dim objrs, sql 

	sql  = "select * from  dbo.wb_account_tim  where userid = '"&userid &"' and clientcode='" & custcode &"' and timcode = '" & timcode & "'"
	
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