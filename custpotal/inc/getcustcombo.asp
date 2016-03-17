<%@CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
	' parameter
	' global : null 경우 전체 광고주, gbl 인경우 광고계약이 된 광고주만
	' custcode : 광고주 코드 -  선택할  광고주이 없는 경우  null, 코드가 있으면 해당 광고주을 선택
	Dim scope : scope = UCase(Trim(Request("scope")))
	Dim custcode : custcode = UCase(Trim(Request("custcode")))
	If scope = "" Then scope = null 
	If custcode = "" Then custcode = null
	Dim sql 
	If  Not IsNull(scope) Then 
		sql = "select highcustcode, custname from sc_cust_hdr where medflag = 'A' and use_flag=1 order by custname"
	Else 
		sql = "select distinct h.highcustcode, h.custname from wb_contact_mst m inner join sc_cust_dtl d on m.custcode = d.custcode inner join sc_cust_hdr h on d.highcustcode = h.highcustcode order by custname"
	End if
	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	Dim rs : Set rs = cmd.Execute 
	Set cmd = Nothing
	
	Response.write "<select id='cmbcustcode' name='cmbcustcode' style='width:266px'>"&vbCrLf
	response.write "<option value=''> -- 광고주를 선택하세요 -- </option>"&vbCrLf
	Do Until rs.eof 
		response.write "<option value='" & rs(0) & "' "
		If custcode = rs(0) Then Response.write "selected"
		response.write ">" & rs(1) & "</option>" & vbCrLf
		rs.movenext
	Loop
	Response.write "</select>"
%>
