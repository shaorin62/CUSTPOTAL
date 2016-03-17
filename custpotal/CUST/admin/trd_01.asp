<%@ language="vbscript" codepage="65001" %>
<!--#include virtual="/inc/func.asp" -->
<%


	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1000

	dim cyear : cyear = request.querystring("cyear")
	dim cmonth : cmonth = request.querystring("cmonth")
	dim cyear2 : cyear2 = request.querystring("cyear2")
	dim cmonth2 : cmonth2 = request.querystring("cmonth2")
	dim custcode2 : custcode2 = request.querystring("custcode2")
	dim initpage : initpage  = request.querystring("initpage")

	if cyear =  "" then cyear = Cstr(Year(date))
	if cmonth = "" then cmonth = Cstr(Month(Date))
	if cyear2 =  "" then cyear2 = Cstr(Year(date))
	if cmonth2 = "" then cmonth2 = Cstr(Month(Date))

	if len(cmonth) = 1 then cmonth = "0"&cmonth
	if len(cmonth2) = 1 then cmonth2 = "0"&cmonth2


	dim yearmon : yearmon = cyear&cmonth
	dim yearmon2 : yearmon2 = cyear2&cmonth2
%>


<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<div style='margin-top:10px;'>
<TABLE  width="100%">
	<TR>
		<TD ><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle" id="subtitle"> 월별 매체 집행내역</span></TD>
		<TD  align="right" valign="top">  <span class="navigator"  id="navigate">광고비집행 &gt;  월별 매체 집행내역</span></TD>
	</TR>
</TABLE>
</div>

<!-- 검색 영역 -->
	<Div id='searchtag' style='margin-top:10px;'>
	<table width="1030" height="35" border="0" cellpadding="0" cellspacing="0" >
	   <tr>
		 <td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
		 <td width="80%" align="left" background="/images/bg_search.gif"><%=search_cyear2cmonth2(cyear, cmonth, cyear2, cmonth2, custcode2)%> <A HREF="#" ><img src="/images/btn_search.gif" width="39" height="20" align="absmiddle" border="0" onclick="getdata('trd_01.asp'); return false;"></A></td>
		 <td width="20%" align="right" background="/images/bg_search.gif"><img src="/images/btn_excel.gif" width="85" height="22" align="absmiddle" vspace="5" class="stylelink" onclick="get_excel_sheet();"></td>
		 <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
	   </tr>
	</table>
	</div>

<!-- 컨텐츠 영역 -->
<%
	if initpage =1 then

	sql = "select m.yearmon,c.custname,sum(case when m.med_flag = '01' then isnull(amt,0) else 0 end) as 'M01' , sum(case when m.med_flag = '02' or m.med_flag = '03' then isnull(amt,0) else 0 end) as 'M02' , sum(case when m.med_flag = 'A2' then isnull(amt,0) else 0 end) as 'M03' , sum(case when m.med_flag = '10' or m.med_flag = '20' then isnull(amt,0) else 0 end) as 'M04', sum(case when m.med_flag = 'B'  then isnull(amt,0) else 0 end) as 'M05', sum(case when m.med_flag = 'C' then isnull(amt,0) else 0 end) as 'M06', sum(case when m.med_flag = 'O' then isnull(amt,0) else 0 end) as 'M07', sum(case when m.med_flag = 'D' then isnull(amt,0) else 0 end) as 'M08' , sum(isnull(amt,0)) as 'TOTAL' from dbo.md_report_mst_v  m inner join dbo.sc_cust_hdr c on m.clientcode = c.highcustcode  where  (m.yearmon  between '"&yearmon&"' and '"&yearmon2&"')  and c.highcustcode like '" & custcode2 &"%'  group by  m.yearmon, c.custname with rollup "

	call get_recordset(objrs, sql)

	Dim cyearmon, custname2, m01, m02, m03, m04, m05, m06, m07, m08, total, cnt, prev,  mt01, mt02, mt03, mt04, mt05, mt06, mt07, mt08, subtotal
	If Not objrs.eof Then
		Set cyearmon = objrs("yearmon")
		Set custname2 = objrs("custname")
		Set m01 = objrs("m01")
		Set m02 = objrs("m02")
		Set m03 = objrs("m03")
		Set m04 = objrs("m04")
		Set m05 = objrs("m05")
		Set m06 = objrs("m06")
		Set m07 = objrs("m07")
		Set m08 = objrs("m08")
		Set total = objrs("total")
	End if

	if custcode2 = "" then custcode2 = null

%>
<div id='#contents' style='margin-top:10px;width:1030px;overflow-x:scroll;'>

<link href="/style.css" rel="stylesheet" type="text/css">
			  <table width="1014" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" >
                      <tr class="trhd">
                        <td width="90" align="center">년월</td>
						<% if isnull(custcode2) then %>
                        <td width="124" align="center">구 분</td>
						<% end if%>
                        <td width="90" align="center" >TV</td>
                        <td width="90" align="center">RD</td>
                        <td width="90" align="center">CATV</td>
                        <td width="90" align="center">지상파 DMB</td>
                        <td width="90" align="center">신문</td>
                        <td width="90" align="center">잡지</td>
                        <td width="90" align="center">인터넷</td>
                        <td width="90" align="center">옥외</td>
                        <td width="90" align="center">계</td>
                      </tr>
				<!--  -->
				<%
					cnt = 0
					do until objrs.eof
					if not (not isnull(cyearmon) and  isnull(custname2) and cnt = 0) then
				%>
				<% if isnull(custcode2) then %>
					<% if isnull(cyearmon) and isnull(custname2) and cnt=1 then%>
					  <tr  class="trbd" bgcolor="#FFC1C1" >
							<td  align="center" colspan="2">총합 </td>
					<% elseif not isnull(cyearmon) and isnull(custname2) and cnt = 1 then %>
					  <tr  class="trbd" bgcolor="#FFFFC1" >
							<td  align="center" colspan="2"><%response.write left(cyearmon,4) & "-" & right(cyearmon,2)%> 소계 </td>
					<% else %>
					  <tr  class="trbd" bgcolor="#FFFFFF" >
							<td  align="center" ><%if prev = cyearmon then Response.write "" else response.write left(cyearmon,4) & "-" & right(cyearmon,2)%> </td>
							<td  align="left" style="padding-left:10px;"><%=custname2%></td>
					<% end if %>
				<% else '구분 컬럼은 사업부 선택일 때는 보이지 않도록 해야 한다. %>
					<% if isnull(cyearmon) and isnull(custname2) and cnt=1 then%>
					  <tr  class="trbd" bgcolor="#FFC1C1" >
							<td  align="center" >총합 </td>
					<% elseif not isnull(cyearmon) and isnull(custname2) and cnt = 1 then %>
					  <tr  class="trbd" bgcolor="#FFFFC1" >
							<td  align="center" ><%response.write left(cyearmon,4) & "-" & right(cyearmon,2)%> 소계 </td>
					<% else %>
					  <tr  class="trbd" bgcolor="#FFFFFF" >
							<td  align="center" ><%if prev = cyearmon then Response.write "" else response.write left(cyearmon,4) & "-" & right(cyearmon,2)%> </td>
					<% end if %>
				<% end if %>

                        <td  align="right" ><%If m01 = "0" Then response.write "-" Else response.write FormatNumber(m01,0)%>&nbsp;&nbsp;</td>
                        <td  align="right" ><%If m02 = "0" Then response.write "-" Else response.write FormatNumber(m02,0)%>&nbsp;&nbsp;</td>
                        <td  align="right" ><%If m03 = "0" Then response.write "-" Else response.write FormatNumber(m03,0)%>&nbsp;&nbsp;</td>
                        <td  align="right" ><%If m04 = "0" Then response.write "-" Else response.write FormatNumber(m04,0)%>&nbsp;&nbsp;</td>
                        <td  align="right" ><%If m05 = "0" Then response.write "-" Else response.write FormatNumber(m05,0)%>&nbsp;&nbsp;</td>
                        <td  align="right" ><%If m06 = "0" Then response.write "-" Else response.write FormatNumber(m06,0)%>&nbsp;&nbsp;</td>
                        <td  align="right" ><%If m07 = "0" Then response.write "-" Else response.write FormatNumber(m07,0)%>&nbsp;&nbsp;</td>
                        <td  align="right" ><%If m08 = "0" Then response.write "-" Else response.write FormatNumber(m08,0)%>&nbsp;&nbsp;</td>
                        <td  align="right" ><%If total = "0" Then response.write "-" Else response.write FormatNumber(total,0)%>&nbsp;&nbsp;</td>
                  </tr>
				<%
					end if
					if prev = cyearmon then cnt = 1 else cnt = 0
					prev = cyearmon

					objrs.movenext
				loop
				objrs.close
				set objrs = nothing
				%>
              </table>
</div>
<%end if %>