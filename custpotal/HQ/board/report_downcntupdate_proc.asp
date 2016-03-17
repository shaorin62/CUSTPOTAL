<!--#Include virtual="/inc/getdbcon.asp"-->
<!--#Include virtual="/inc/func.asp"-->
<%

	dim ridx : ridx = request("ridx")
	dim midx : midx = request("midx")

	dim objrs, objrs1, sql

	sql = "select downcnt  from dbo.wb_report where ridx = " & ridx

	call get_recordset(objrs, sql)

	dim cnt
	if not objrs.eof then
		set cnt = objrs("downcnt")
	end If
	

	sql = "select ridx,  downcnt  from dbo.wb_report where ridx = " & ridx

	call set_recordset(objrs1, sql)

	Dim strcnt

	strcnt = CLng(cnt) + 1
	objrs1.fields("ridx").value =ridx
	objrs1.fields("downcnt").value =strcnt

	
	objrs1.update
	objrs1.close

	Set objrs = Nothing
	Set objrs1 = Nothing

%>

<SCRIPT LANGUAGE="JavaScript">
<!--
	this.close();
//-->
</SCRIPT>