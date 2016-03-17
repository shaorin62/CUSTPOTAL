
<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
	body {
		background-color:transparent;
		font-size:12px;
	}
	.trhd {
		font-size:12px;
		height: 30px;
		color: #F9F1EA;
		background-color:#9A9A9A;
		font-weight: bolder;
	}

	.trbd {
		font-size:12px;
		height: 30px;
		color: #000000;
	}
</style>
<%
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

	Response.CacheControl  = "public"
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader  "Content-Disposition" , "attachment; filename=세부 집행내역.xls"

	dim objrs, sql

	sql = " select isnull(c.custname, 'z') as custname, max(j.seqname)  as seqname , isnull(dbo.md_get_mattername_fun(m.mattercode),'z') as program_name ,  m.pub_date,  max(c2.custname) as custname2, max(m.std_step)  as std_step, max(m.std_cm) as std_cm, max(m.col_deg)  as col_deg, max(m.pub_face) pub_face, sum(isnull(m.amt,0))  as amt, max(c3.custname) as custname3, max(demandday)  as demandday , max(m.memo) as note  from dbo.md_booking_medium m  inner join dbo.sc_cust_dtl c  on c.custcode = m.timcode inner join dbo.sc_cust_dtl c2 on c2.custcode = m.medcode  left outer join dbo.sc_subseq_dtl j on j.seqno = m.subseq  left outer join dbo.sc_cust_hdr c3 on c3.highcustcode = m.exclientcode  where m.clientcode LIKE '%"&custcode2&"%'  and substring(m.pub_date,1,6) = '"&yearmon&"' and m.med_flag = 'MP01'  group by c.custname, dbo.md_get_mattername_fun(m.mattercode) , pub_date  with rollup order by isnull(c.custname, 'z'),  max(j.seqname) ,  dbo.md_get_mattername_fun(m.mattercode) desc,  convert(smalldatetime, m.pub_date) desc "

	call get_recordset(objrs, sql)

	Dim custname, seqname, program_name, pub_date, custname2, std_step, std_cm, col_deg,pub_face, code_name, amt, custname3, demandday, note,  prev_custname, prev_seqname
	If Not objrs.eof Then
		Set custname = objrs("custname")
		Set seqname = objrs("seqname")
		Set program_name = objrs("program_name")
		Set pub_date = objrs("pub_date")
		Set custname2 = objrs("custname2")
		Set std_step = objrs("std_step")
		Set std_cm = objrs("std_cm")
		Set col_deg = objrs("col_deg")
		Set pub_face = objrs("pub_face")
		Set amt = objrs("amt")
		Set custname3 = objrs("custname3")
		Set demandday = objrs("demandday")
		Set note = objrs("note")
	End if

%>
				  <table width="1260" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" >
                      <tr class="trhd">
                        <td width="130" align="center">소속사</td>
                        <td width="100" align="center">브랜드명</td>
                        <td width="180" align="center" >소재명</td>
                        <td width="60" align="center">게재일</td>
                        <td width="150" align="center">매체명</td>
                        <td width="90" align="center" colspan='4'>규격</td>
                        <td width="90" align="center">게재면</td>
                        <td width="100" align="center">광고비</td>
                        <td width="90" align="center">제작사</td>
                        <td width="60" align="center">청구일</td>
                        <td width="90" align="center">비고</td>
                      </tr>
				<!-- custname, seqname, program_name, pub_date, custname2, std_step, std_cm, col_deg, code_name, amt, custname3, demandday, note, prev_medflag, prev_custname, prev_seqname  -->
				<% do until objrs.eof 	%>
				<% If custname = "z" And program_name = "z" Then %>
                  <tr  class="trbd" bgcolor="#FFC1C1" >
                        <td >&nbsp;&nbsp;TOTAL</td>
                        <td >&nbsp;&nbsp;</td>
                        <td >&nbsp;&nbsp;</td>
                        <td align="center"></td>
                        <td >&nbsp;&nbsp;</td>
                        <td  width="30" align="right">&nbsp;</td>
                        <td  width="30" align="left">&nbsp;</td>
                        <td  width="30" align="right"> &nbsp;</td>
                        <td  width="30" align="center"></td>
                        <td  width="90" align="center"></td>
                        <td  width="100" align="right"> <%=FormatNumber(amt,0)%> &nbsp;&nbsp;</td>
                        <td width="90" align="center"></td>
                        <td width="90" align="center"></td>
                        <td width="90" align="left"></td>
                      </tr>
				<% ElseIf custname <> "z" And program_name <> "z" And IsNull(pub_date) Then %>
                  <tr  class="trbd" bgcolor="#CCFFFF" >
                        <td bgcolor="#FFFFFF">&nbsp;&nbsp;</td>
                        <td bgcolor="#FFFFFF">&nbsp;&nbsp;</td>
                        <td >&nbsp;&nbsp;<%=program_name%> 합계</td>
                        <td align="center"></td>
                        <td >&nbsp;&nbsp;</td>
                        <td  width="30" align="right">&nbsp;</td>
                        <td  width="30" align="left">&nbsp;</td>
                        <td  width="30" align="right"> &nbsp;</td>
                        <td  width="30" align="center"></td>
                        <td  width="90" align="center"></td>
                        <td  width="100" align="right"> <%=FormatNumber(amt,0)%> &nbsp;&nbsp;</td>
                        <td width="90" align="center"></td>
                        <td width="90" align="center"></td>
                        <td width="90" align="left"></td>
                      </tr>
				<% ElseIf custname <> "z" And program_name = "z" And IsNull(pub_date) Then %>
                  <tr  class="trbd" bgcolor="#FFDFDF" >
                        <td >&nbsp;&nbsp;<%=custname%> 요약</td>
                        <td >&nbsp;&nbsp;</td>
                        <td >&nbsp;&nbsp;</td>
                        <td align="center"></td>
                        <td >&nbsp;&nbsp;</td>
                        <td  width="30" align="right">&nbsp;</td>
                        <td  width="30" align="left">&nbsp;</td>
                        <td  width="30" align="right"> &nbsp;</td>
                        <td  width="30" align="center"></td>
                        <td  width="90" align="center"></td>
                        <td  width="100" align="right"> <%=FormatNumber(amt,0)%> &nbsp;&nbsp;</td>
                        <td width="90" align="center"></td>
                        <td width="90" align="center"></td>
                        <td width="90" align="left"></td>
                      </tr>
				<% Else %>
                  <tr  class="trbd" bgcolor="#FFFFFF" >
                        <td >&nbsp;&nbsp;<%If prev_custname <> custname Then response.write custname Else response.write ""%></td>
                        <td >&nbsp;&nbsp;<%If prev_seqname <> seqname Then response.write seqname Else response.write ""%></td>
                        <td >&nbsp;&nbsp;<%=program_name%></td>
                        <td align="center"><%If Not IsNull(pub_date) Then response.write Mid(pub_date,5,2) & "/" &  Right(pub_date,2) %></td>
                        <td >&nbsp;&nbsp;<%=custname2%></td>
                        <td  width="30" align="right"><%=std_step%>&nbsp;</td>
                        <td  width="30" align="left">&nbsp;단</td>
                        <td  width="30" align="right"> <%=std_cm%>&nbsp;</td>
                        <td  width="30" align="center"><%=col_deg%></td>
                        <td  width="90" align="center"><%=pub_face%></td>
                        <td  width="100" align="right"> <%=FormatNumber(amt,0)%> &nbsp;&nbsp;</td>
                        <td width="90" align="center"><%=custname3%></td>
                        <td width="90" align="center"><%=demandday%></td>
                        <td width="90" align="left"><%=note%></td>
                      </tr>
				<%
						End If
						prev_custname = custname
						prev_seqname = seqname
						objrs.movenext
					loop
					objrs.close
					set objrs = nothing
				%>
              </table>