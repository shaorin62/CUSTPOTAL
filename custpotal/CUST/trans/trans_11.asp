
<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<body oncontextmenu='return false' >
<%
	dim cyear, cyear2, cmonth, cmonth2
	cyear = request("cyear")
	if cyear = "" then cyear = year(date)
	cmonth = request("cmonth")
	if cmonth = "" then cmonth = month(date)
	if len(cmonth) = 1 then cmonth = "0"&cmonth
	dim yearmon : yearmon = cyear&cmonth
	dim custcode : custcode =request("tcustcode")
	dim custcode2 : custcode2 = request("tcustcode2")
	dim intTR, intTD' for 을 위한 일련번호
	dim objrs, sql' 레코드셋, 쿼리문장
	dim hearder' 테이블 헤어로 이용할 레코드셋 배열
	dim intLoop
	' 선택된 광고주에 해당하는 사업부서 쿼리
	sql = "select highcustcode, custname from dbo.sc_cust_hdr where  MEDFLAG = 'A'  order by custname"
	call get_recordset(objrs, sql)

	dim str
	str = "<select name='tcustcode2' id='tcustcode2'>"
	do until objrs.eof
		str = str & "<option value='" & objrs("highcustcode") & "'"
			if custcode2 = objrs("highcustcode") then str = str & " selected"
		str = str & ">" & objrs("custname") & "</option>"
		objrs.movenext
	Loop
	str = str & "</select>"
	objrs.close

	'if custcode2 = "" or custcode = custcode2 then custcode2 = Null

	if request.cookies("class") = "D" or request.cookies("class") = "H"  then
		custcode2 = request.cookies("custcode2")
	end if

	if not isnull(custcode2) then

	'sql = "select m.clientsubcode, c.custname from dbo.md_report_mst m inner join dbo.sc_cust_temp c on m.clientsubcode = c.custcode where m.medflag='O' and substring(m.yearmon, 1, 4)='"&cyear&"' and m.clientcode = '"&custcode&"' and m.clientsubcode like '" & custcode2 &"%' group by m.clientsubcode, c.custname order by m.clientsubcode"

	sql = " select m.timcode, c.custname  from dbo.md_report_mst_v m inner join dbo.sc_cust_dtl c on m.timcode = c.custcode where m.med_flag='O' and substring(m.yearmon, 1, 4)='"&cyear&"' and m.clientcode = '"&custcode2&"' group by m.timcode, c.custname order by m.timcode  "



	call get_recordset(objrs, sql)

	if not objrs.eof then
	dim clientsubcode, clientsubname, tRow
	tRow = objrs.recordcount+1
	hearder = objrs.getRows()

	dim trRow : trRow = ubound(hearder, 1)    ' tr count
	dim tdRow : tdRow = ubound(hearder, 2)	'td count

%>

<TABLE  width="<%=(tRow* 13 * 100)+350%>" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC">
<% for intTR = 0 to trRow %>
	<tr class="trhd2">
	<% if intTR = 0 then %>
		<TD rowspan="2" width="150" align="center">사이트명</TD>
	<% for intLoop = 1 to 12 %>
		<TD colspan="<%=tRow%>"  align="center"><%=intLoop%>월</TD>
	<% next %>
		<TD colspan="<%=tRow%>"  align="center">TOTAL</TD>
		<TD rowspan="2" width="200"  align="center"> 청구지 </TD>
	<% else%>
	<% for intLoop = 1 to 13 %>
	<% for intTD = 0 to tdRow %>
		<TD width="100" align="center"><%=hearder(intTR,intTD)%></TD>
	<% next %>
		<TD width="100" align="center">소계</TD>
	<% next %>
	<% end if %>
	</tr>
<% next %>

<!-- 데이터 릿스트 -->
<%
	sql = "select c2.custname "

	for intLoop = 1 to 12

	if len(intLoop) = 1 then intLoop = "0" & intLoop
		for intTD = 0 to tdRow
			sql = sql & ", sum(case when substring(m.yearmon, 5, 2) = '" & intLoop & "' and m.timcode = '" & hearder(0,intTD) & "' then isnull(amt, 0) else 0 end) as '" & intLoop &"' "
		next
			sql = sql & ", sum(case when substring(m.yearmon, 5, 2) = '" & intLoop & "' then isnull(amt, 0) else 0 end) as  '" & intLoop & " total'"
	next
		for intTD = 0 to tdRow
			sql = sql & ", sum(case when m.timcode = '" & hearder(0,intTD) & "' then isnull(amt, 0) else 0 end) as 'TOTAL'"
		next
			sql = sql & ", sum(isnull(amt, 0)) as  '" & intLoop & " total', max(c.companyname) "
	sql = sql & " from dbo.md_report_mst_v m inner join dbo.sc_cust_hdr c on m.real_med_code = c.highcustcode inner join dbo.sc_cust_dtl c2 on m.medcode = c2.custcode where m.med_flag='O' and substring(m.yearmon, 1, 4)='"&cyear&"' and m.clientcode = '"&custcode2&"'  group by c2.custname  with rollup"




	call get_recordset(objrs, sql)

	dim tblData : tblData = objrs.getrows
	trRow = ubound(tblData, 2)
	tdRow = ubound(tblData, 1)
%>

<% for intTR = 0 to trRow %>
<% if intTR = trRow then %>
	<TR  class="trbd" bgcolor="#FFFFC1" >
<% else %>
	<TR  class="trbd" bgcolor="#FFFFFF" >
<% end if %>
		<% for intTD = 0 to tdRow %>

<TD   style="<%if intTD =0 then response.write "padding-left:10px;" else response.write "padding-right:10px;text-align:right"%>">
<%
	if intTD > 0 and intTD < tdRow then
		response.write formatnumber(tblData(intTD,intTR),0)
	else
		if isnull(tblData(intTD,intTR)) then
			response.write "합계"
		elseif intTR = trRow and intTD = tdRow then
			response.write ""
		else
			response.write tblData(intTD,intTR)
		end if
	end if
%> </TD>
		<% next %>
	</TR>
<% next %>
<%
	end if
	else

	sql = "select m.clientsubcode, c.custname from dbo.md_report_mst m inner join dbo.sc_cust_temp c on m.clientsubcode = c.custcode where m.medflag='O' and substring(m.yearmon, 1, 4)='"&cyear&"' and m.clientcode = '"&custcode&"' and m.clientsubcode like '" & custcode2 &"%' group by m.clientsubcode, c.custname order by m.clientsubcode"

	call get_recordset(objrs, sql)

	if not objrs.eof then

	tRow = objrs.recordcount
	hearder = objrs.getRows()

	trRow = ubound(hearder, 1)    ' tr count
	tdRow = ubound(hearder, 2)	'td count

%>

<TABLE  width="<%=(tRow* 13 * 100)+350%>" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC">
	<tr class="trhd2">
		<TD  width="150" align="center">사이트명</TD>
	<% for intLoop = 1 to 12 %>
		<TD  width ="100" align="center"><%=intLoop%>월</TD>
	<% next %>
		<TD  align="center">TOTAL</TD>
		<TD width="200"  align="center"> 청구지 </TD>
	</tr>

<!-- 데이터 릿스트 -->
<%
	sql = "select c.custname "

	for intLoop = 1 to 12

	if len(intLoop) = 1 then intLoop = "0" & intLoop
		for intTD = 0 to tdRow
			sql = sql & ", sum(case when substring(m.yearmon, 5, 2) = '" & intLoop & "' and m.clientsubcode = '" & hearder(0,intTD) & "' then isnull(amt, 0) else 0 end) as '" & intLoop &"' "
		next
	next
		for intTD = 0 to tdRow
			sql = sql & ", sum(case when m.clientsubcode = '" & hearder(0,intTD) & "' then isnull(amt, 0) else 0 end) as 'TOTAL'"
		next
		sql = sql & ", max(c.companyname) "
	sql = sql & " from dbo.md_report_mst m inner join dbo.sc_cust_temp c on m.medcode = c.custcode inner join dbo.sc_cust_temp c2 on m.medcode = c2.custcode where m.medflag='O' and substring(m.yearmon, 1, 4)='"&cyear&"' and m.clientcode = '"&custcode&"' and m.clientsubcode like '" & custcode2 &"%' group by c.custname  with rollup"

	call get_recordset(objrs, sql)

	tblData = objrs.getrows
	trRow = ubound(tblData, 2)
	tdRow = ubound(tblData, 1)
%>

<% for intTR = 0 to trRow %>
<% if intTR = trRow then %>
	<TR  class="trbd" bgcolor="#FFFFC1" >
<% else %>
	<TR  class="trbd" bgcolor="#FFFFFF" >
<% end if %>
		<% for intTD = 0 to tdRow %>

<TD   style="<%if intTD =0 then response.write "padding-left:10px;" else response.write "padding-right:10px;text-align:right"%>">
<%
	if intTD > 0 and intTD < tdRow then
		response.write formatnumber(tblData(intTD,intTR),0)
	else
		if isnull(tblData(intTD,intTR)) then
			response.write "합계"
		elseif intTR = trRow and intTD = tdRow then
			response.write ""
		else
			response.write tblData(intTD,intTR)
		end if
	end if
%> </TD>
		<% next %>
	</TR>
<% next %>
<% end if %>
<% end if %>
</body>
<!-- page end  -->
<% if request.cookies("class") = "C" then %>
<SCRIPT LANGUAGE="JavaScript">
<!--
	var custcode2 = parent.document.getElementById("custcode2") ;
	custcode2.innerHTML = "<%=str%>";
//-->
</SCRIPT>
<% end if%>