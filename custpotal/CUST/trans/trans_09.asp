
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
	' 해당 사업부를 콤보박스로 세팅
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

	sql = "select subseq, dbo.pd_jobcust_name_fun(subseq) as subname,  count(campaign_code) as subcount from (select subseq, campaign_code from dbo.md_internet_medium m where yearmon ='" & cyear&cmonth & "'  and m.clientcode = '" & custcode2 & "' group by subseq, campaign_code) as a group by subseq"


	call get_recordset(objrs, sql)

	if not objrs.eof then
	dim hdrow : hdrow = objrs.getrows
	dim hdrowcount : hdrowcount = ubound(hdrow, 1)
	dim hdcolcount : hdcolcount = ubound(hdrow, 2)
	objrs.close

	'sql = "select b.seqname, c.campaign_name, substring(m.tbrdstdate,5,2) + '/' + substring(m.tbrdstdate,7,2) + '~' + substring(m.tbrdeddate,5,2) + '/' + substring(m.tbrdeddate,7,2), m.subseq, m.campaign_code  from dbo.md_internet_medium m inner join dbo.sc_jobcust b on b.seqno = m.subseq inner join dbo.md_internet_campaign c on m.campaign_code = c.campaign_code where m.clientcode = '" & custcode2 & "'  and m.yearmon = '" & cyear&cmonth & "' group by b.seqname,  c.campaign_name, m.tbrdstdate, m.tbrdeddate,m.subseq, m.campaign_code order by m.subseq, m.campaign_code "

	sql = "select b.seqname, c.campaign_name,  substring(m.tbrdstdate,5,2) + '/' +  substring(m.tbrdstdate,7,2) + '~' +  substring(m.tbrdeddate,5,2) + '/' +  substring(m.tbrdeddate,7,2), m.subseq,  m.campaign_code  from dbo.md_internet_medium m  inner join dbo.sc_subseq_dtl b  on b.seqno = m.subseq  inner join dbo.md_internet_campaign c on m.campaign_code = c.campaign_code  where m.clientcode = '" & custcode2 & "' and m.yearmon = '" & cyear&cmonth & "'  group by b.seqname, c.campaign_name, m.tbrdstdate, m.tbrdeddate,m.subseq, m.campaign_code  order by m.subseq, m.campaign_code "



	call get_recordset(objrs, sql)
	if not objrs.eof then
		hearder = objrs.getrows()
	dim tmpRow
	dim trRow : trRow = ubound(hearder, 1)
	dim tdRow : tdRow = ubound(hearder, 2)
	dim intTD1
	dim prev_brand, colcount

	'hearder, 1 = tr
	'hreader, 2 = td
%>

<TABLE  width="<%=((tdRow+4) * 100)+150%>" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC">
<% for intTR = 0 to trRow-2 %>
<TR class="trhd2">


	<% if intTR = 0 then %>
		<TD rowspan="3" width="150" align="center">사이트명</TD>
		<% for intLoop = 0 to hdcolcount %>
			<TD    align="center" colspan="<%=hdrow(hdrowcount, intLoop)%>"><%=hdrow(1, intLoop)%> (<%=hdrow(hdrowcount, intLoop)%>건)</TD>
		<% next %>
		<TD rowspan="3" width="100"  align="center"> 소 계 </TD>
		<TD rowspan="3" width="200"  align="center"> 청구지 </TD>
	<%else%>
	<% for intTD = 0 to tdRow %>
				<TD  width="100" align="center"><%=hearder(intTR,intTD)%></TD>
	<%next	%>
	<% end if %>
</TR>
<% next%>
<!-- 데이터 릿스트 -->
<%
	'sql = "select m.subseq, m.campaign_code from dbo.md_internet_medium m  inner join dbo.sc_jobcust b on b.seqno = m.subseq inner join dbo.md_internet_campaign c on m.campaign_code = c.campaign_code where m.clientcode = '" & custcode & "' and m.clientsubcode like '" & custcode2 & "%' and m.yearmon = '" & cyear&cmonth & "' group by m.subseq, m.campaign_code order by m.subseq, m.campaign_code"

	sql = "select m.subseq, m.campaign_code  from dbo.md_internet_medium m  inner join dbo.sc_subseq_dtl b on b.seqno = m.subseq  inner join dbo.md_internet_campaign c on m.campaign_code = c.campaign_code  where m.clientcode = '" & custcode2 & "' and m.yearmon = '" & cyear&cmonth & "'  group by m.subseq, m.campaign_code order by m.subseq, m.campaign_code "

	call get_recordset(objrs, sql)

	sql = "select c.custname"
	do until objrs.eof
		sql = sql & ", sum(case when m.campaign_code = '"&objrs("campaign_code")&"' and m.subseq = '"&objrs("subseq")&"' then isnull(amt,0) else 0 end)"
		objrs.movenext
		loop
	sql = sql & ", sum(isnull(amt,0)), max(c2.companyname)"
	sql = sql & " from dbo.md_internet_medium m inner join dbo.sc_cust_dtl c on m.medcode = c.custcode inner join dbo.sc_cust_hdr c2 on m.real_med_code = c2.highcustcode where m.clientcode = '"&custcode2&"'  and  m.yearmon = '"&cyear&cmonth&"' group by c.custname with rollup "

	objrs.close

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
			response.write "소계"
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
<!-- page end  -->
<% end if %>
<% end if %>
</body>
<SCRIPT LANGUAGE="JavaScript">
<!--
	var custcode2 = parent.document.getElementById("custcode2") ;
	custcode2.innerHTML = "<%=str%>";
//-->
</SCRIPT>
