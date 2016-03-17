<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<%
	Dim userid : userid = request.cookies("userid")
	Dim pcustcode : pcustcode = request("custcode")
	Dim pteamcode : pteamcode = request("teamcode")
	Dim strstat : strstat = request("strstat")
	Dim cyear : cyear = request("cyear")
	Dim cmonth : cmonth = request("cmonth")
	Dim cyear2 : cyear2 = request("cyear2")
	Dim cmonth2 : cmonth2 = request("cmonth2")


	If cyear = "" Then cyear = Year(date)
	If cmonth = "" Then cmonth = Month(date)
	If cyear2 = "" Then cyear2 = Year(date)
	If cmonth2 = "" Then cmonth2 = Month(date)
	If Len(cmonth) = 1 Then cmonth = "0"&cmonth
	If Len(cmonth2) = 1 Then cmonth2 = "0"&cmonth2
	dim schdate : schdate = cyear&cmonth

	dim sdate : sdate = dateserial(cyear, cmonth, "01")
	dim edate : edate = dateadd("d", -1, dateadd("m", 1,  DateSerial(cyear2, cmonth2, "01")))



	Dim sql
	Dim chkcount
	chkcount= 1
	Dim Custcodesql
	Dim Custcoderecord
	Dim Timcodesql
	Dim Timcoderecord
	Dim objrs_1
	Dim objrs
	Dim rs
	Dim strwhere
dim Custcoderecord_cnt
	Dim objrs_cnt

	if pcustcode = "" or pcustcode = null then
		'=========================================================================================
		Custcodesql = "select clientcode from wb_account_cust where userid ='" & userid & "' "

		Set objrs_cnt = server.CreateObject("adodb.recordset")
		objrs_cnt.activeconnection = application("connectionstring")
		objrs_cnt.cursorLocation = aduseclient
		objrs_cnt.cursortype = adopenstatic
		objrs_cnt.locktype = adlockoptimistic
		objrs_cnt.source = Custcodesql
		objrs_cnt.open

		Custcoderecord_cnt = objrs_cnt.recordcount
		'=========================================================================================

		if Custcoderecord_cnt = 0 then
			Custcodesql = "select distinct h.highcustcode as clientcode from wb_contact_mst m inner join sc_cust_dtl d on m.custcode = d.custcode inner join sc_cust_hdr h on d.highcustcode = h.highcustcode "

			Set objrs_1 = server.CreateObject("adodb.recordset")
			objrs_1.activeconnection = application("connectionstring")
			objrs_1.cursorLocation = aduseclient
			objrs_1.cursortype = adopenstatic
			objrs_1.locktype = adlockoptimistic
			objrs_1.source = Custcodesql
			objrs_1.open

			Custcoderecord = objrs_1.recordcount
		else
			Custcodesql = "select clientcode from wb_account_cust where userid ='" & userid & "' "

			Set objrs_1 = server.CreateObject("adodb.recordset")
			objrs_1.activeconnection = application("connectionstring")
			objrs_1.cursorLocation = aduseclient
			objrs_1.cursortype = adopenstatic
			objrs_1.locktype = adlockoptimistic
			objrs_1.source = Custcodesql
			objrs_1.open

			Custcoderecord = objrs_1.recordcount
		end if
		'=========================================================================================



		if not objrs_1.eof then
			do until objrs_1.eof
				'=========================================================================================
				Timcodesql = "select timcode from wb_account_tim where userid ='" & userid & "' and clientcode = '" & objrs_1("clientcode") &"'"

				Set objrs = server.CreateObject("adodb.recordset")
				objrs.activeconnection = application("connectionstring")
				objrs.cursorLocation = aduseclient
				objrs.cursortype = adopenstatic
				objrs.locktype = adlockoptimistic
				objrs.source = Timcodesql
				objrs.open

				Timcoderecord = objrs.recordcount
				'=========================================================================================

				if chkcount > 1 then
					sql = sql  & " Union all "
				end if

				if strstat = "9" then
					strwhere = " and isnull(md.categoryidx,'') = '9' "
				elseif strstat = "10" then
					strwhere = " and isnull(md.categoryidx,'') = '10' "
				elseif strstat = "11" then
					strwhere = " and isnull(md.categoryidx,'') = '11' "
				elseif strstat = "135" then
					strwhere = " and isnull(md.categoryidx,'') = '135' "
				else
					strwhere = " and isnull(md.categoryidx,'') = '9' "
				end if

				sql = sql  & " select c.contidx, c.title, c.firstdate, c.startdate, "
				sql = sql  & " c.enddate, isnull(sum(m.monthly),0) as monthly, "
				sql = sql  & " c.flag , "
				sql = sql  & " isnull(max(s.a_val),0) a_val, isnull(max(s.b_val),0) b_val, isnull(max(s.c_val),0) c_val, isnull(max(s.d_val),0) d_val, isnull(max(s.e_val),0) e_val, "
				sql = sql  & " isnull(max(s.a_val),0) + isnull(max(s.b_val),0) + isnull(max(s.c_val),0) + isnull(max(s.d_val),0) + isnull(max(s.e_val),0) tot, "
				sql = sql  & " max(s.class) totclass"
				sql = sql  & " from wb_contact_mst c  "
				sql = sql  & " left outer join sc_cust_dtl d on c.custcode = d.custcode  "
				sql = sql  & " inner  join wb_account_cust n on d.highcustcode  = n.clientcode and n.userid='"&userid&"' and n.clientcode =  '" & objrs_1("clientcode") &"' "
				If Timcoderecord > 0 then
					sql = sql  & " inner  join wb_account_tim t on c.custcode = t.timcode and t.userid='"&userid&"' and t.clientcode = '" & objrs_1("clientcode") &"' "
				End If
				sql = sql  & " left outer join vw_contact_exe_monthly m  "
				sql = sql  & " on m.contidx = c.contidx  "
				sql = sql  & " left outer join wb_validation_class s on c.contidx = s.contidx and s.isuse = 1 "
				sql = sql  & " left outer join  "
				sql = sql  & " ( "
				sql = sql  & " 	select contidx, max(DBO.WB_CATEGORYIDX_FUN(categoryidx)) categoryidx "
				sql = sql  & " 	from  wb_contact_md  "
				sql = sql  & " 	group by contidx "
				sql = sql  & " ) md on c.contidx = md.contidx "
				sql = sql  & " where  m.cyear+m.cmonth = '"&schdate&"' " & strwhere
				sql = sql  & " and isnull(DBO.WB_CATEGORYIDX_FUN(md.categoryidx) ,'') <> '' "
				sql = sql  & " and c.flag = 'B' "
				sql = sql  & " group by c.contidx, c.title, c.firstdate,  "
				sql = sql  & " c.startdate, c.enddate, isnull(c.totalprice,0), c.custcode ,c.flag  "

				chkcount = chkcount +1
				objrs_1.movenext
			Loop
			'sql = sql  & " order by c.enddate,  c.title,  contidx desc  "
		end if


else

	'=========================================================================================
	Timcodesql = "select timcode from wb_account_tim where userid ='" & userid & "' and clientcode ='" & pcustcode & "'"

	Set objrs = server.CreateObject("adodb.recordset")
	objrs.activeconnection = application("connectionstring")
	objrs.cursorLocation = aduseclient
	objrs.cursortype = adopenstatic
	objrs.locktype = adlockoptimistic
	objrs.source = Timcodesql
	objrs.open

	Timcoderecord = objrs.recordcount
	'=========================================================================================


	if strstat = "9" then
		strwhere = " and isnull(md.categoryidx,'') = '9' "
	elseif strstat = "10" then
		strwhere = " and isnull(md.categoryidx,'') = '10' "
	elseif strstat = "11" then
		strwhere = " and isnull(md.categoryidx,'') = '11' "
	elseif strstat = "135" then
		strwhere = " and isnull(md.categoryidx,'') = '135' "
	else
		strwhere = " and isnull(md.categoryidx,'') = '9' "
	end if

	sql = sql  & " select c.contidx, c.title, c.firstdate, c.startdate, "
	sql = sql  & " c.enddate, isnull(sum(m.monthly),0) as monthly, "
	sql = sql  & " c.flag , "
	sql = sql  & " isnull(max(s.a_val),0) a_val, isnull(max(s.b_val),0) b_val, isnull(max(s.c_val),0) c_val, isnull(max(s.d_val),0) d_val, isnull(max(s.e_val),0) e_val, "
	sql = sql  & " isnull(max(s.a_val),0) + isnull(max(s.b_val),0) + isnull(max(s.c_val),0) + isnull(max(s.d_val),0) + isnull(max(s.e_val),0) tot, "
	sql = sql  & " max(s.class) totclass"
	sql = sql  & " from wb_contact_mst c  "
	sql = sql  & " left outer join sc_cust_dtl d on c.custcode = d.custcode  "
	sql = sql  & " inner  join wb_account_cust n on d.highcustcode  = n.clientcode and n.userid='"&userid&"' "
	If Timcoderecord > 0 then
		sql = sql  & " inner  join wb_account_tim t on c.custcode = t.timcode and t.userid='"&userid&"' "
	End If
	sql = sql  & " left outer join vw_contact_exe_monthly m  "
	sql = sql  & " on m.contidx = c.contidx  "
	sql = sql  & " left outer join wb_validation_class s on c.contidx = s.contidx and s.isuse = 1 "
	sql = sql  & " left outer join  "
	sql = sql  & " ( "
	sql = sql  & " 	select contidx, max(DBO.WB_CATEGORYIDX_FUN(categoryidx)) categoryidx "
	sql = sql  & " 	from  wb_contact_md  "
	sql = sql  & " 	group by contidx "
	sql = sql  & " ) md on c.contidx = md.contidx "
	sql = sql  & " where  m.cyear+m.cmonth = '"&schdate&"' "
	sql = sql  & " and d.highcustcode like '"&pcustcode&"%'  "
	sql = sql  & " and c.custcode like  '"&pteamcode&"%'   "& strwhere
	sql = sql  & " and isnull(DBO.WB_CATEGORYIDX_FUN(md.categoryidx) ,'') <> '' "
	sql = sql  & " and c.flag = 'B' "
	sql = sql  & " group by c.contidx, c.title, c.firstdate,  "
	sql = sql  & " c.startdate, c.enddate, isnull(c.totalprice,0), c.custcode ,c.flag  "
	'sql = sql  & " order by c.enddate,  c.title,  contidx desc  "

end if


	Set rs = server.CreateObject("adodb.recordset")
	rs.activeconnection = application("connectionstring")
	rs.cursorLocation = adUseClient
	rs.cursorType = adOpenStatic
	rs.lockType = adLockOptimistic
	rs.source = sql
	rs.open

	Dim totalrecord : totalrecord = rs.recordcount
	Dim real_totalrecord : real_totalrecord = rs.recordcount

	Dim contidx : Set contidx = rs(0)
	Dim title : Set title = rs(1)
	Dim firstdate : Set firstdate = rs(2)
	Dim startdate : Set startdate = rs(3)
	Dim enddate : Set enddate = rs(4)
	Dim monthly : Set monthly = rs(5)
	Dim flag : Set flag = rs(6)
	Dim a_val : Set a_val = rs(7)
	Dim b_val : Set b_val = rs(8)
	Dim c_val : Set c_val = rs(9)
	Dim d_val : Set d_val = rs(10)
	Dim e_val : Set e_val = rs(11)
	Dim tot : Set tot = rs(12)
	Dim totclass : Set totclass = rs(13)

	Dim grandmonthly : grandmonthly = 0
	Dim granda_val : granda_val = 0
	Dim grandb_val : grandb_val = 0
	Dim grandc_val : grandc_val = 0
	Dim grandd_val : grandd_val = 0
	Dim grande_val : grande_val = 0
	Dim grandtot : grandtot = 0


	Response.CacheControl  = "public"
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader  "Content-Disposition" , "attachment; filename="&cyear&"년"&cmonth&"월 효용성평가.xls"

%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  >
<h2> <u>효용성평가 ('<%=cyear%>.<%=CInt(cmonth)%>') 월</u> </h2>
	 <table border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" width='2000'>
          <tr>
            <td valign="top"  colspan='2'>
<% If strstat = "9"  Then %>
				  <table border="3px"   width="1030" cellpadding="0" cellspacing="0" bordercolor="#8d652b" id="contact">
				  <thead>
					<tr height='20' align='center'>
						<th width="30" style=' text-align:center;' rowSpan="2">no</th>
						<th style=' text-align:center;' rowSpan="2">매체명</th>
						<th width="70" style=' text-align:center;' rowSpan="2">최초계약일자</th>
						<th style=' text-align:center;' colSpan="2">계약기간</th>
						<th width="90" style=' text-align:center;' rowSpan="2">월광고료(원)</th>
						<th style=' text-align:center;' colSpan="5">평가항목</th>
						<th width="50" style=' text-align:center;' rowSpan="2">총점(100)</th>
						<th width="50" style=' text-align:center;' rowSpan="2">등급</th>
					</tr>
					<tr height='35'>
						<th width="70"style=' text-align:center;'>시작일</th>
						<th width="70"style=' text-align:center;'>종료일</th>
						<th width="60" style=' text-align:center;'>지역(40)</th>
						<th width="60" style=' text-align:center;'>매체사양(20)</th>
						<th width="60" style=' text-align:center;'>가시환경(30)</th>
						<th width="60" style=' text-align:center;'>경쟁환경(5)</th>
						<th width="60" style=' text-align:center;'>기타(5)</th>
					</tr>
					</thead>
					</table>
					<table border="1"   width="1030" cellpadding="0" cellspacing="0" id="contact">
					<tbody id='tbody'>
					<%
						Do Until rs.eof
					%>
					<tr class="trbd" height='32'>
						<td width="30"  class="hd none" style='padding-left:10px; text-align:left;'><span name="totalrecord"><%=totalrecord%></span></td>
						<td width="300"  class="hd none" style="padding-left:5px;"><%=cutTitle(title, 44)%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=firstdate%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=startdate%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=enddate%></td>
						<td width="90"  class="hd none" style='padding-right:3px; text-align:right;'><%=FormatNumber(monthly, 0)%></td>
						<td width="60"  class="hd none" style='padding-right:3px; text-align:right;'><%=a_val%></td>
						<td width="60"  class="hd none" style='padding-right:3px; text-align:right;'><%=b_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=c_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=d_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=e_val%></td>
						<td width="50"  class="hd none" style='padding-left:3px; text-align:center'><%=tot%></td>
						<td width="50"  class="hd none" style='padding-left:3px; text-align:center;'><%=totclass%></td>
					</tr>

				  <%
							totalrecord = totalrecord - 1
							grandmonthly = CDbl(grandmonthly) + CDbl(monthly)
							granda_val = CDbl(granda_val) + CDbl(a_val)
							grandb_val = CDbl(grandb_val) + CDbl(b_val)
							grandc_val = CDbl(grandc_val) + CDbl(c_val)
							grandd_val = CDbl(grandd_val) + CDbl(d_val)
							grande_val = CDbl(grande_val) + CDbl(e_val)
							grandtot = CDbl(grandtot) + CDbl(tot)
							rs.movenext
						Loop

						if cdbl(real_totalrecord ) <> 0 then
							grandtot = CDbl(grandtot) / real_totalrecord /4
						end if
				  %>
				  </tbody>
				  <tfoot>
                <tr height="30">
                    <td class="hd left"  colspan='5' style="text-align:center;"><strong>평균</strong> </td>
                    <td class="hd center" style=' text-align:right; font-weight:bold'><%=formatnumber(grandmonthly,0)%>&nbsp;</td>
                    <td  class="hd center" style='padding-right:3px; text-align:right;'><%=granda_val%></td>
					<td  class="hd center" style='padding-right:3px; text-align:right;'><%=grandb_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grandc_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grandd_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grande_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:center;'><%=grandtot%></td>
                    <td class="hd right"></td>
                  </tr>
				  </tfoot>
              </table>
<% elseIf strstat = "10"  Then %>
				  <table border="3px"   width="1030" cellpadding="0" cellspacing="0" bordercolor="#8d652b" id="contact">
				  <thead>
					<tr height='20' align='center'>
						<th width="30" style=' text-align:center;' rowSpan="2">no</th>
						<th style=' text-align:center;' rowSpan="2">매체명</th>
						<th width="70" style=' text-align:center;' rowSpan="2">최초계약일자</th>
						<th style=' text-align:center;' colSpan="2">계약기간</th>
						<th width="90" style=' text-align:center;' rowSpan="2">월광고료(원)</th>
						<th style=' text-align:center;' colSpan="5">평가항목</th>
						<th width="50" style=' text-align:center;' rowSpan="2">총점(100)</th>
						<th width="50" style=' text-align:center;' rowSpan="2">등급</th>
					</tr>
					<tr height='35'>
						<th width="70"style=' text-align:center;'>시작일</th>
						<th width="70"style=' text-align:center;'>종료일</th>
						<th width="60" style=' text-align:center;'>지역(30)</th>
						<th width="60" style=' text-align:center;'>매체사양(25)</th>
						<th width="60" style=' text-align:center;'>가시환경(25)</th>
						<th width="60" style=' text-align:center;'>경쟁환경(10)</th>
						<th width="60" style=' text-align:center;'>기타(10)</th>
					</tr>
					</thead>
					</table>
					<table border="1"   width="1030" cellpadding="0" cellspacing="0" id="contact">
					<tbody id='tbody'>
					<%
						Do Until rs.eof
					%>
					<tr class="trbd" height='32'>
						<td width="30"  class="hd none" style='padding-left:10px; text-align:left;'><span name="totalrecord"><%=totalrecord%></span></td>
						<td width="300"  class="hd none" style="padding-left:5px;"><%=cutTitle(title, 44)%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=firstdate%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=startdate%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=enddate%></td>
						<td width="90"  class="hd none" style='padding-right:3px; text-align:right;'><%=FormatNumber(monthly, 0)%></td>
						<td width="60"  class="hd none" style='padding-right:3px; text-align:right;'><%=a_val%></td>
						<td width="60"  class="hd none" style='padding-right:3px; text-align:right;'><%=b_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=c_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=d_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=e_val%></td>
						<td width="50"  class="hd none" style='padding-left:3px; text-align:center'><%=tot%></td>
						<td width="50"  class="hd none" style='padding-left:3px; text-align:center;'><%=totclass%></td>
					</tr>

				  <%
							totalrecord = totalrecord - 1
							grandmonthly = CDbl(grandmonthly) + CDbl(monthly)
							granda_val = CDbl(granda_val) + CDbl(a_val)
							grandb_val = CDbl(grandb_val) + CDbl(b_val)
							grandc_val = CDbl(grandc_val) + CDbl(c_val)
							grandd_val = CDbl(grandd_val) + CDbl(d_val)
							grande_val = CDbl(grande_val) + CDbl(e_val)
							grandtot = CDbl(grandtot) + CDbl(tot)
							rs.movenext
						Loop

						if cdbl(real_totalrecord ) <> 0 then
							grandtot = CDbl(grandtot) / real_totalrecord /4
						end if
				  %>
				  </tbody>
				  <tfoot>
                <tr height="30">
                    <td class="hd left"  colspan='5' style="text-align:center;"><strong>평균</strong> </td>
                    <td class="hd center" style=' text-align:right; font-weight:bold'><%=formatnumber(grandmonthly,0)%>&nbsp;</td>
                    <td  class="hd center" style='padding-right:3px; text-align:right;'><%=granda_val%></td>
					<td  class="hd center" style='padding-right:3px; text-align:right;'><%=grandb_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grandc_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grandd_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grande_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:center;'><%=grandtot%></td>
                    <td class="hd right"></td>
                  </tr>
				  </tfoot>
              </table>
<% elseIf strstat = "11"  Then %>
				  <table border="3px"   width="1030" cellpadding="0" cellspacing="0" bordercolor="#8d652b" id="contact">
				  <thead>
					<tr height='20' align='center'>
						<th width="30" style=' text-align:center;' rowSpan="2">no</th>
						<th style=' text-align:center;' rowSpan="2">매체명</th>
						<th width="70" style=' text-align:center;' rowSpan="2">최초계약일자</th>
						<th style=' text-align:center;' colSpan="2">계약기간</th>
						<th width="90" style=' text-align:center;' rowSpan="2">월광고료(원)</th>
						<th style=' text-align:center;' colSpan="5">평가항목</th>
						<th width="50" style=' text-align:center;' rowSpan="2">총점(100)</th>
						<th width="50" style=' text-align:center;' rowSpan="2">등급</th>
					</tr>
					<tr height='35'>
						<th width="70"style=' text-align:center;'>시작일</th>
						<th width="70"style=' text-align:center;'>종료일</th>
						<th width="60" style=' text-align:center;'>지역(30)</th>
						<th width="60" style=' text-align:center;'>매체사양(25)</th>
						<th width="60" style=' text-align:center;'>가시환경(25)</th>
						<th width="60" style=' text-align:center;'>기타(10)</th>
						<th width="60" style=' text-align:center;'></th>
					</tr>
					</thead>
					</table>
					<table border="1"   width="1030" cellpadding="0" cellspacing="0" id="contact">
					<tbody id='tbody'>
					<%
						Do Until rs.eof
					%>
					<tr class="trbd" height='32'>
						<td width="30"  class="hd none" style='padding-left:10px; text-align:left;'><span name="totalrecord"><%=totalrecord%></span></td>
						<td width="300"  class="hd none" style="padding-left:5px;"><%=cutTitle(title, 44)%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=firstdate%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=startdate%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=enddate%></td>
						<td width="90"  class="hd none" style='padding-right:3px; text-align:right;'><%=FormatNumber(monthly, 0)%></td>
						<td width="60"  class="hd none" style='padding-right:3px; text-align:right;'><%=a_val%></td>
						<td width="60"  class="hd none" style='padding-right:3px; text-align:right;'><%=b_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=c_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=d_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'>0</td>
						<td width="50"  class="hd none" style='padding-left:3px; text-align:center'><%=tot%></td>
						<td width="50"  class="hd none" style='padding-left:3px; text-align:center;'><%=totclass%></td>
					</tr>

				  <%
							totalrecord = totalrecord - 1
							grandmonthly = CDbl(grandmonthly) + CDbl(monthly)
							granda_val = CDbl(granda_val) + CDbl(a_val)
							grandb_val = CDbl(grandb_val) + CDbl(b_val)
							grandc_val = CDbl(grandc_val) + CDbl(c_val)
							grandd_val = CDbl(grandd_val) + CDbl(d_val)
							grande_val = CDbl(grande_val) + CDbl(e_val)
							grandtot = CDbl(grandtot) + CDbl(tot)
							rs.movenext
						Loop
						if cdbl(real_totalrecord ) <> 0 then
							grandtot = CDbl(grandtot) / real_totalrecord /4
						end if
				  %>
				  </tbody>
				  <tfoot>
                <tr height="30">
                    <td class="hd left"  colspan='5' style="text-align:center;"><strong>평균</strong> </td>
                    <td class="hd center" style=' text-align:right; font-weight:bold'><%=formatnumber(grandmonthly,0)%>&nbsp;</td>
                    <td  class="hd center" style='padding-right:3px; text-align:right;'><%=granda_val%></td>
					<td  class="hd center" style='padding-right:3px; text-align:right;'><%=grandb_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grandc_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grandd_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grande_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:center;'><%=grandtot%></td>
                    <td class="hd right"></td>
                  </tr>
				  </tfoot>
              </table>
<% elseIf strstat = "135"  Then %>
				  <table border="3px"   width="1030" cellpadding="0" cellspacing="0" bordercolor="#8d652b" id="contact">
				  <thead>
					<tr height='20' align='center'>
						<th width="30" style=' text-align:center;' rowSpan="2">no</th>
						<th style=' text-align:center;' rowSpan="2">매체명</th>
						<th width="70" style=' text-align:center;' rowSpan="2">최초계약일자</th>
						<th style=' text-align:center;' colSpan="2">계약기간</th>
						<th width="90" style=' text-align:center;' rowSpan="2">월광고료(원)</th>
						<th style=' text-align:center;' colSpan="5">평가항목</th>
						<th width="50" style=' text-align:center;' rowSpan="2">총점(100)</th>
						<th width="50" style=' text-align:center;' rowSpan="2">등급</th>
					</tr>
					<tr height='35'>
						<th width="70"style=' text-align:center;'>시작일</th>
						<th width="70"style=' text-align:center;'>종료일</th>
						<th width="60" style=' text-align:center;'>지역(30)</th>
						<th width="60" style=' text-align:center;'>매체사양(40)</th>
						<th width="60" style=' text-align:center;'>가시환경(15)</th>
						<th width="60" style=' text-align:center;'>기타(15)</th>
						<th width="60" style=' text-align:center;'></th>
					</tr>
					</thead>
					</table>
					<table border="1"   width="1030" cellpadding="0" cellspacing="0" id="contact">
					<tbody id='tbody'>
					<%
						Do Until rs.eof
					%>
					<tr class="trbd" height='32'>
						<td width="30"  class="hd none" style='padding-left:10px; text-align:left;'><span name="totalrecord"><%=totalrecord%></span></td>
						<td width="300"  class="hd none" style="padding-left:5px;"><%=cutTitle(title, 44)%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=firstdate%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=startdate%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=enddate%></td>
						<td width="90"  class="hd none" style='padding-right:3px; text-align:right;'><%=FormatNumber(monthly, 0)%></td>
						<td width="60"  class="hd none" style='padding-right:3px; text-align:right;'><%=a_val%></td>
						<td width="60"  class="hd none" style='padding-right:3px; text-align:right;'><%=b_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=c_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=d_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'>0</td>
						<td width="50"  class="hd none" style='padding-left:3px; text-align:center'><%=tot%></td>
						<td width="50"  class="hd none" style='padding-left:3px; text-align:center;'><%=totclass%></td>
					</tr>

				  <%
							totalrecord = totalrecord - 1
							grandmonthly = CDbl(grandmonthly) + CDbl(monthly)
							granda_val = CDbl(granda_val) + CDbl(a_val)
							grandb_val = CDbl(grandb_val) + CDbl(b_val)
							grandc_val = CDbl(grandc_val) + CDbl(c_val)
							grandd_val = CDbl(grandd_val) + CDbl(d_val)
							grande_val = CDbl(grande_val) + CDbl(e_val)
							grandtot = CDbl(grandtot) + CDbl(tot)
							rs.movenext
						Loop

						if cdbl(real_totalrecord ) <> 0 then
							grandtot = CDbl(grandtot) / real_totalrecord /4
						end if
				  %>
				  </tbody>
				  <tfoot>
                <tr height="30">
                    <td class="hd left"  colspan='5' style="text-align:center;"><strong>평균</strong> </td>
                    <td class="hd center" style=' text-align:right; font-weight:bold'><%=formatnumber(grandmonthly,0)%>&nbsp;</td>
                    <td  class="hd center" style='padding-right:3px; text-align:right;'><%=granda_val%></td>
					<td  class="hd center" style='padding-right:3px; text-align:right;'><%=grandb_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grandc_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grandd_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grande_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:center;'><%=grandtot%></td>
                    <td class="hd right"></td>
                  </tr>
				  </tfoot>
              </table>
<% Else %>
				  <table border="3px"   width="1030" cellpadding="0" cellspacing="0" bordercolor="#8d652b" id="contact">
				  <thead>
					<tr height='20' align='center'>
						<th width="30" style=' text-align:center;' rowSpan="2">no</th>
						<th style=' text-align:center;' rowSpan="2">매체명</th>
						<th width="70" style=' text-align:center;' rowSpan="2">최초계약일자</th>
						<th style=' text-align:center;' colSpan="2">계약기간</th>
						<th width="90" style=' text-align:center;' rowSpan="2">월광고료(원)</th>
						<th style=' text-align:center;' colSpan="5">평가항목</th>
						<th width="50" style=' text-align:center;' rowSpan="2">총점(100)</th>
						<th width="50" style=' text-align:center;' rowSpan="2">등급</th>
					</tr>
					<tr height='35'>
						<th width="70"style=' text-align:center;'>시작일</th>
						<th width="70"style=' text-align:center;'>종료일</th>
						<th width="60" style=' text-align:center;'>지역(40)</th>
						<th width="60" style=' text-align:center;'>매체사양(20)</th>
						<th width="60" style=' text-align:center;'>가시환경(30)</th>
						<th width="60" style=' text-align:center;'>경쟁환경(5)</th>
						<th width="60" style=' text-align:center;'>기타(5)</th>
					</tr>
					</thead>
					</table>
					<table border="1"   width="1030" cellpadding="0" cellspacing="0" id="contact">
					<tbody id='tbody'>
					<%
						Do Until rs.eof
					%>
					<tr class="trbd" height='32'>
						<td width="30"  class="hd none" style='padding-left:10px; text-align:left;'><span name="totalrecord"><%=totalrecord%></span></td>
						<td width="300"  class="hd none" style="padding-left:5px;"><%=cutTitle(title, 44)%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=firstdate%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=startdate%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=enddate%></td>
						<td width="90"  class="hd none" style='padding-right:3px; text-align:right;'><%=FormatNumber(monthly, 0)%></td>
						<td width="60"  class="hd none" style='padding-right:3px; text-align:right;'><%=a_val%></td>
						<td width="60"  class="hd none" style='padding-right:3px; text-align:right;'><%=b_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=c_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=d_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=e_val%></td>
						<td width="50"  class="hd none" style='padding-left:3px; text-align:center'><%=tot%></td>
						<td width="50"  class="hd none" style='padding-left:3px; text-align:center;'><%=totclass%></td>
					</tr>

				  <%
							totalrecord = totalrecord - 1
							grandmonthly = CDbl(grandmonthly) + CDbl(monthly)
							granda_val = CDbl(granda_val) + CDbl(a_val)
							grandb_val = CDbl(grandb_val) + CDbl(b_val)
							grandc_val = CDbl(grandc_val) + CDbl(c_val)
							grandd_val = CDbl(grandd_val) + CDbl(d_val)
							grande_val = CDbl(grande_val) + CDbl(e_val)
							grandtot = CDbl(grandtot) + CDbl(tot)
							rs.movenext
						Loop

						if cdbl(real_totalrecord ) <> 0 then
							grandtot = CDbl(grandtot) / real_totalrecord /4
						end if
				  %>
				  </tbody>
				  <tfoot>
                <tr height="30">
                    <td class="hd left"  colspan='5' style="text-align:center;"><strong>평균</strong> </td>
                    <td class="hd center" style=' text-align:right; font-weight:bold'><%=formatnumber(grandmonthly,0)%>&nbsp;</td>
                    <td  class="hd center" style='padding-right:3px; text-align:right;'><%=granda_val%></td>
					<td  class="hd center" style='padding-right:3px; text-align:right;'><%=grandb_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grandc_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grandd_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grande_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:center;'><%=grandtot%></td>
                    <td class="hd right"></td>
                  </tr>
				  </tfoot>
              </table>
<% End If %>




			  <div id="debugConsole"></div>
			  </td>
          </tr>
      </table></td>
    </tr>
  </table>
 </body>
</html>