<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim custcode : custcode = request("custcode")
	dim custcode2 :custcode2 = request("custcode2")
	dim custcode3 : custcode3 = request("custcode3")
	dim seqno : seqno = request("seqno")
	if seqno = "" then seqno = null
	if custcode2 =  "" then custcode2 = null
	if custcode3 = "" then  custcode3 = null

	'response.write typename(custcode) & " : " & typename(custcode2) & " : " & typename(custcode3) & " : " & typename(seqno)

	dim rs , sql, str, str2, str3, str4
	sql = "SELECT custcode, custname  FROM dbo.SC_CUST_TEMP where medflag='A' and custcode <>  highcustcode and highcustcode = '" & custcode &"' order by custname"
	call get_recordset(rs, sql)

	str = "<select name='selcustcode2' style='width:207;'>"
	str = str & "<option value=''>전체 사업부</option>"
	Do Until rs.eof
		str = str & "<option value='"&rs("custcode")&"' "
			if custcode2 = rs("custcode") then str = str &  " selected "
		str = str & "> " & rs("custname") &"</option>"
		rs.movenext
	loop
	str = str & "</select>"
	rs.close

	sql = "SELECT seqno, seqname  FROM dbo.SC_JOBCUST where custcode = '" & custcode &"' order by seqname"
	call get_recordset(rs, sql)

	str2 = "<select name='seljobcust' style='width:207;'>"
	str2 = str2 & "<option value=''>전체 브랜드</option>"
	Do Until rs.eof
		str2 = str2 & "<option value='"&rs("seqno")&"' "
			if seqno = rs("seqno") then str2 = str2 &  " selected "
		str2 = str2 & "> " & rs("seqname") &"</option>"
		rs.movenext
	loop
	str2 = str2 & "</select>"
	rs.close

	sql = "select custcode, custname from dbo.sc_cust_temp where medflag='B' and meddiv = '5' order by custname"
	call get_recordset(rs, sql)

	str3 = "<select name='selcustcode3' style='width:207;'>"
	str3 = str3 & "<option value=''>전체 매체사</option>"
	Do Until rs.eof
		str3 = str3 & "<option value='"&rs("custcode")&"' "
			if custcode3 = rs("custcode") then str3 = str3 &  " selected "
		str3 = str3 & "> " & rs("custname") &"</option>"
		rs.movenext
	loop
	str3 = str3 & "</select>"
	rs.close

	sql = "select jobidx, thema from dbo.wb_jobcust j inner join dbo.sc_jobcust j2 on j.seqno = j2.seqno where custcode = '" & custcode &"' order by thema"
	call get_recordset(rs, sql)

	str4 = "<select name='selthema' style='width:207;'>"
	str4 = str4 & "<option value=''>전체 소재</option>"
	Do Until rs.eof
		str4 = str4 & "<option value='"&rs("thema")&"' > " & rs("thema") &"</option>"
		rs.movenext
	loop
	str4 = str4 & "</select>"
	rs.close

'
'	response.write str
'	response.write str2
'	response.write str3
'	response.write str4

	set rs = nothing
	response.write request.cookies("class")

	if request.cookies("class") <> "D" then
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
		var str = "<%=str%>" ;
		var custcode2 = parent.document.getElementById("custcode2");
		custcode2.innerHTML = str ;

	<%if not isnull(seqno) then %>
		var str2 = "<%=str2%>" ;
		var str4 = "<%=str4%>" ;
		var jobcust = parent.document.getElementById("jobcust");
		jobcust.innerHTML = str2 ;

		var thema = parent.document.getElementById("thema");
		thema.innerHTML = str4 ;
	<%end if%>

	<%if not isnull(custcode3) then %>
		var str3 = "<%=str3%>" ;
		var custcode3 = parent.document.getElementById("custcode3");
		custcode3.innerHTML = str3 ;
	<%end if%>

//-->
</SCRIPT>
<% else %>

<SCRIPT LANGUAGE="JavaScript">
<!--
	<%if not isnull(seqno) then %>
		var str2 = "<%=str2%>" ;
		var str4 = "<%=str4%>" ;
		var jobcust = parent.document.getElementById("jobcust");
		jobcust.innerHTML = str2 ;

		var thema = parent.document.getElementById("thema");
		thema.innerHTML = str4 ;
	<%end if%>

	<%if not isnull(custcode3) then %>
		var str3 = "<%=str3%>" ;
		var custcode3 = parent.document.getElementById("custcode3");
		custcode3.innerHTML = str3 ;
	<%end if%>

//-->
</SCRIPT>
<% end if%>