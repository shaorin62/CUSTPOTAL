<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%	 
	Dim progname : progname = Trim(Request("progname"))

	If progname = "" Then progname = null
	Dim sql 
	Dim objrs
	sql = "select userid, program_name , opendate from wb_log_mst where userid = '?'"

	call set_recordset(objrs, sql)

	objrs.addnew
	objrs.fields("userid").value = request.Cookies("userid")
	objrs.fields("program_name").value = progname
	objrs.fields("opendate").value = now
	objrs.update
	objrs.close

'%>