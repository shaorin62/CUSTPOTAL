<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim custcode : custcode = request("custcode")
	dim custcode2 : custcode2 = request("custcode2")
	dim seqno : seqno = request("seqno")
	dim searchstring : searchstring = request("searchstring")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	dim custname : custname = request("custname")

	dim temp_filename : temp_filename = cyear&"."&cmonth&"("&custname&")"

	Response.CacheControl  = "public"
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader  "Content-Disposition" , "attachment; filename="&temp_filename&".xls"
%>

<%
	dim objrs, sql
	sql = "select m.title, m2.title, m2.qty, d.monthprice , m.startdate, m.enddate, c3.custname, m2.trust, m2.locate, j2.seqname, c.custname, j.thema from dbo.vw_medium_category v inner join dbo.wb_contact_md m2 on v.mdidx = m2.categoryidx inner join dbo.wb_contact_mst m on m.contidx = m2.contidx inner join dbo.sc_cust_temp c on c.custcode = m.custcode inner join dbo.wb_contact_md_dtl d on m2.contidx = d.contidx and m2.sidx = d.sidx inner join dbo.sc_cust_temp c3 on m2.custcode = c3.custcode inner join dbo.wb_jobcust j on d.jobidx = j.jobidx inner join dbo.sc_jobcust j2 on j.seqno = j2.seqno where d.cyear = '2008' and d.cmonth='12'"

	call get_recordset(objrs, sql)

	dim total_price, total_monthprice, total_expense, total_income, total_incomeratio
%>
<style type="text/css">
	body {background-color:transparent}
</style>
<body  oncontextmenu="return false">
<table width="100%" border="1" cellspacing="0" cellpadding="0">
  <tr>
    <td rowspan="2" align="center">매체명</td>
    <td rowspan="2" align="center">장소</td>
    <td rowspan="2" align="center">수량(면)</td>
    <td rowspan="2" align="center">면</td>
    <td rowspan="2" align="center">규격/재질</td>
    <td rowspan="2" align="center">최초<br>
      계약일자</td>
    <td colspan="2" align="center">계약기간</td>
    <td rowspan="2" align="center">총광고료(원)</td>
    <td rowspan="2" align="center">월광고료(원)</td>
    <td rowspan="2" align="center">월청구금액(원)</td>
    <td rowspan="2" align="center">내수액</td>
    <td rowspan="2" align="center">내수율</td>
    <td rowspan="2" align="center">광고내용</td>
    <td rowspan="2" align="center">관련부서</td>
    <td rowspan="2" align="center">매채사</td>
  </tr>
  <tr>
    <td align="center">시작일</td>
    <td align="center">종료일</td>
  </tr>
  <% do until objrs.eof %>
  <tr>
    <td><%=objrs("title")%></td>
    <td><%=objrs("locate")%></td>
    <td><%=objrs("qty")%>(<%=objrs("unit")%>)</td>
    <td><%=objrs("side")%>&nbsp;</td>
    <td><%=objrs("standard")%> (<%=objrs("quality")%>)</td>
    <td><%=objrs("firstdate")%></td>
    <td><%=objrs("startdate")%></td>
    <td><%=objrs("enddate")%></td>
    <td><%=formatnumber(objrs("totalprice"),0)%></td>
    <td><%=formatnumber(objrs("monthprice"),0)%></td>
    <td><%=formatnumber(objrs("expense"),0)%></td>
    <td><%=formatnumber(objrs("monthprice") - objrs("expense"),0)%></td>
    <td><%if objrs("monthprice") <> 0 then response.write formatnumber(objrs("monthprice") - objrs("expense")/objrs("monthprice"),2)  else response.write formatnumber(0,2)%></td>
    <td><%=objrs("seqname")%></td>
    <td><%=objrs("custname2")%></td>
    <td><%=objrs("custname3")%></td>
  </tr>
  <%
	total_price = total_price + objrs("totalprice")
	total_monthprice = total_monthprice + objrs("monthprice")
	total_expense = total_expense + objrs("expense")

	total_income = total_monthprice - total_expense
	if total_income = 0 then
		total_incomeratio = 0
	else
		total_incomeratio = total_income / total_monthprice * 100
	end if
	objrs.movenext
	Loop

	objrs.close
	set objrs = nothing




'	dim objStream: Set objStream = Server.CreateObject("ADODB.Stream")
'	objStream.Open
'	objStream.Type=1
'	objStream.LoadFromFile Server.MapPath(".") & "\" & temp_filename
'
'	dim download : download = objStream.Read
'	Response.BinaryWrite download
'	Set objStream = nothing
  %>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><%=formatnumber(total_price,0)%></td>
    <td><%=formatnumber(total_monthprice,0)%></td>
    <td><%=formatnumber(total_expense,0)%></td>
    <td><%=formatnumber(total_income,0)%></td>
    <td><%=formatnumber(total_incomeratio,2)%></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>