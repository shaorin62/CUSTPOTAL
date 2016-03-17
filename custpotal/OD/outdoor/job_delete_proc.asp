<!--#include virtual="/inc/getdbcon.asp"-->
<!--#include virtual="/inc/func.asp"-->
<%
'	dim item
'	for each item in request.form
'		response.write item & " :" & request.form(item) & "<BR>"
'	next
'	response.end

	dim jobidx : jobidx = request("jobidx")
	
	dim objrs, sql
	sql = "select * from dbo.wb_contact_md_dtl where jobidx = " & jobidx
	call get_recordset(objrs, sql)

	if not objrs.eof then 
		response.write "<script type='text/javascript'> alert('계약 정보에 등록된 소재입니다.\n\n삭제하실 수 없습니다.'); history.back(); </script>"
		response.flush
		respons.end
	end if

	objrs.close

	sql = "select seqno, thema, uuser, udate from dbo.wb_jobcust where jobidx = " & jobidx
	call set_recordset(objrs, sql)

	objrs.fields("thema").value = thema
	objrs.fields("uuser").value = request.cookies("userid")
	objrs.fields("udate").value = date
	objrs.update

	objrs.close
	set objrs = nothing

%>
<script language="JavaScript">
<!--
	window.opener.location.reload();
	this.close();
//-->
</script>