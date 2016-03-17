<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	'Response.AddHeader = "content-Disposition"
	dim custcode : custcode = request("custcode")
	dim custcode2 : custcode2 = request("custcode2")
	dim searchstring : searchstring = request("searchstring")
	dim s_date : s_date = request("s_date")
	dim e_date : e_date = request("e_date")

	dim temp_filename : temp_filename = "계약목록_"&year(s_date)&"_"&month(s_date)

	Response.ContentType ="application/x-msexcel"
	Response.AddHeader "Content-Disposition" , "attachment; filename="&temp_filename&".xls"
%>

<%
	dim objrs, sql
	sql = "select c.contidx, c.title,  c.firstdate, c.startdate, c.enddate, s.custname, t.monthprice, t.expense   from dbo.wb_contact_mst c inner join dbo.sc_cust_temp s on c.custcode = s.custcode   left outer  join dbo.vw_contact_totalprice t on c.contidx = t.contidx   where s.highcustcode like '" & custcode &"%' and  c.custcode like '" & custcode2 &"%' and c.title like '%" & searchstring &"%'   and c.startdate <= '" & e_date & "' and c.enddate >= '" & s_date &"'  order by c.contidx desc"
	
	call get_recordset(objrs, sql)

	dim fso : Set fso = Server.Createobject("Scripting.FileSystemObject")
	dim act : Set act = fso.CreateTextFile(Server.MapPath(".") & "\" & temp_filename,true)

act.WriteLine "<html xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
act.WriteLine "<Head>"
act.WriteLine "<!--<xml>"
act.WriteLine "<x:ExcelWorkbook>"
act.WriteLine "<x:ExcelWorksheets>"
act.WriteLine "<x:ExcelWorksheet>"
act.WriteLine "<x:Name>Members</x:Name>"
act.WriteLine "<x:worksheetOptions>"
act.WriteLine "<x:print>"
act.WriteLine "<x:validPrinterInfo/>"
act.WriteLine "</x:Print>"
act.WriteLine "</x:worksheetOption>"
act.WriteLine "</x:ExcelWorksheet>"
act.WriteLine "</x:ExcelWorksheets>"
act.WriteLine "</x:ExcelWorkbook>"
act.WriteLine "</xml>"
act.WriteLine "<-->"
act.WriteLine "</head>"
act.WriteLine "<body>"
act.WriteLine "<table>"
act.WriteLine "<tr>"
act.WriteLine "<td>매체명</td>"
act.WriteLine "<td>최초계약일</td>"
act.WriteLine "<td>시작일</td>"
act.WriteLine "<td>종료일</td>"
act.WriteLine "<td>sumamt</td>"
act.WriteLine "<td>vat</td>"
act.WriteLine "<td>semu</td>"
act.WriteLine "<td>bp</td>"
act.WriteLine "<td>demandday</td>"
act.WriteLine "<td>vendor</td>"
act.WriteLine "<td>taxyearmon</td>"
act.WriteLine "<td>taxno</td>"
act.WriteLine "<td>gbn</td>"
act.WriteLine "</tr>"
Do until objrs.Eof

		if day(startdate) = "1" then 
			period = datediff("m", startdate, enddate)+1
		else 
			period = datediff("m", startdate, enddate)
		end if

act.WriteLine "<tr>"
act.WriteLine "<td>"
act.WriteLine " "& objrs("title")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& objrs("firstdate")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& objrs("startdate")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& objrs("enddate")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& objrs("custname")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& objrs("monthprice")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& objrs("expense")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& objrs("expense")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& objrs("expense")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& objrs("expense")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& objrs("expense")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& objrs("expense")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& objrs("expense")
act.WriteLine "</td>"
act.WriteLine "</tr>"
objrs.movenext
Loop
act.WriteLine "</table>"
act.WriteLine "</body>"
act.WriteLine "</html>"
act.close
dim objStream: Set objStream = Server.CreateObject("ADODB.Stream")
objStream.Open
objStream.Type=1
objStream.LoadFromFile Server.MapPath(".") & "\" & temp_filename

dim download : download = objStream.Read
Response.BinaryWrite download 
Set objStream = nothing
%>
<script>
this.close();
</script>