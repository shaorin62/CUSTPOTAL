<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<%
	dim userid : userid = request("userid")

	dim objrs, sql
	sql = "select password, initpwd, clipinglevel from dbo.wb_account where userid = '" & userid & "' "
	call set_recordset(objrs, sql)

	if not objrs.eof then
		objrs.fields("password").value = objrs.fields("initpwd").value
		objrs.fields("clipinglevel").value = 0
		objrs.update
		response.write "<script> alert('비밀번호가 초기화 되었습니다.'); </script>"
	end if
	objrs.close
	set objrs = nothing
%>