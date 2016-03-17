<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<%
	Dim userid : userid = request.cookies("userid")
	Dim pcustcode : pcustcode = request("custcode")
	Dim pteamcode : pteamcode = request("teamcode")
	Dim cyear : cyear = request("cyear")
	If cyear = "" Then cyear = Year(date)


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

				sql = sql & " select c.highclasscode, "
				sql = sql & " isnull(c.middleclassname, '소계') "
				sql = sql & " ,isnull(sum(case when b.cmonth='01' then isnull(b.monthly,0) else 0 end),0) "
				sql = sql & " ,isnull(sum(case when b.cmonth='02' then isnull(b.monthly,0) else 0 end) ,0) "
				sql = sql & " ,isnull(sum(case when b.cmonth='03' then isnull(b.monthly,0) else 0 end) ,0) "
				sql = sql & " ,isnull(sum(case when b.cmonth='04' then isnull(b.monthly,0) else 0 end) ,0) "
				sql = sql & " ,isnull(sum(case when b.cmonth='05' then isnull(b.monthly,0) else 0 end) ,0) "
				sql = sql & " ,isnull(sum(case when b.cmonth='06' then isnull(b.monthly,0) else 0 end) ,0) "
				sql = sql & " ,isnull(sum(case when b.cmonth='07' then isnull(b.monthly,0) else 0 end) ,0) "
				sql = sql & " ,isnull(sum(case when b.cmonth='08' then isnull(b.monthly,0) else 0 end) ,0) "
				sql = sql & " ,isnull(sum(case when b.cmonth='09' then isnull(b.monthly,0) else 0 end)  ,0)"
				sql = sql & " ,isnull(sum(case when b.cmonth='10' then isnull(b.monthly,0) else 0 end)  ,0)"
				sql = sql & " ,isnull(sum(case when b.cmonth='11' then isnull(b.monthly,0) else 0 end) ,0) "
				sql = sql & " ,isnull(sum(case when b.cmonth='12' then isnull(b.monthly,0) else 0 end) ,0) "
				sql = sql & " ,isnull(sum(isnull(b.monthly,0)) ,0)"
				sql = sql & " from wb_contact_md a "
				sql = sql & " inner join wb_contact_exe b on a.mdidx=b.mdidx and cyear='"&cyear&"' "
				sql = sql & " inner join vw_medium_class c on a.categoryidx=c.catcode "
				sql = sql & " inner join wb_contact_mst d on a.contidx=d.contidx "
				sql = sql  & " left outer join sc_cust_dtl e on e.custcode = d.custcode "
				sql = sql  & " inner  join wb_account_cust n on e.highcustcode  = n.clientcode and n.userid='"&userid&"' and n.clientcode =  '" & objrs_1("clientcode") &"' "
				If Timcoderecord > 0 then
					sql = sql  & " inner  join wb_account_tim t on d.custcode = t.timcode and t.userid='"&userid&"' and t.clientcode = '" & objrs_1("clientcode") &"' "
				End If
				sql = sql & " inner join wb_contact_trans f on f.cyear='"&cyear&"'  and a.contidx=f.contidx and b.cmonth = f.cmonth and a.medcode = f.medcode and (f.isHold='N' or f.isHold='Y')"
				sql = sql & " group by c.highclasscode, c.middleclassname with rollup "
				sql = sql & " having highclasscode is not null "


				chkcount = chkcount +1
				objrs_1.movenext
			Loop
			'sql = sql & " order by highclasscode "
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


	sql = sql & " select c.highclasscode, "
	sql = sql & " isnull(c.middleclassname, '소계') "
	sql = sql & " ,isnull(sum(case when b.cmonth='01' then isnull(b.monthly,0) else 0 end),0) "
	sql = sql & " ,isnull(sum(case when b.cmonth='02' then isnull(b.monthly,0) else 0 end),0) "
	sql = sql & " ,isnull(sum(case when b.cmonth='03' then isnull(b.monthly,0) else 0 end),0) "
	sql = sql & " ,isnull(sum(case when b.cmonth='04' then isnull(b.monthly,0) else 0 end),0) "
	sql = sql & " ,isnull(sum(case when b.cmonth='05' then isnull(b.monthly,0) else 0 end),0) "
	sql = sql & " ,isnull(sum(case when b.cmonth='06' then isnull(b.monthly,0) else 0 end),0) "
	sql = sql & " ,isnull(sum(case when b.cmonth='07' then isnull(b.monthly,0) else 0 end),0) "
	sql = sql & " ,isnull(sum(case when b.cmonth='08' then isnull(b.monthly,0) else 0 end),0) "
	sql = sql & " ,isnull(sum(case when b.cmonth='09' then isnull(b.monthly,0) else 0 end),0) "
	sql = sql & " ,isnull(sum(case when b.cmonth='10' then isnull(b.monthly,0) else 0 end),0) "
	sql = sql & " ,isnull(sum(case when b.cmonth='11' then isnull(b.monthly,0) else 0 end),0) "
	sql = sql & " ,isnull(sum(case when b.cmonth='12' then isnull(b.monthly,0) else 0 end),0) "
	sql = sql & " ,isnull(sum(isnull(b.monthly,0)),0) "
	sql = sql & " from wb_contact_md a "
	sql = sql & " inner join wb_contact_exe b on a.mdidx=b.mdidx and cyear='"&cyear&"' "
	sql = sql & " inner join vw_medium_class c on a.categoryidx=c.catcode "
	sql = sql & " inner join wb_contact_mst d on a.contidx=d.contidx "
	sql = sql  & " left outer join sc_cust_dtl e on e.custcode = d.custcode "
	sql = sql  & " inner  join wb_account_cust n on e.highcustcode  = n.clientcode and n.userid='"&userid&"' "
	If Timcoderecord > 0 then
		sql = sql  & " inner  join wb_account_tim t on d.custcode = t.timcode and t.userid='"&userid&"' "
	End If
	sql = sql & " inner join wb_contact_trans f on f.cyear='"&cyear&"'  and a.contidx=f.contidx and b.cmonth = f.cmonth and a.medcode = f.medcode and (f.isHold='N' or f.isHold='Y')"
	sql = sql & " where d.custcode like '"&pteamcode&"%' and e.highcustcode like '"&pcustcode&"%' "
	sql = sql & " group by c.highclasscode, c.middleclassname with rollup "
	sql = sql & " having highclasscode is not null "
	'sql = sql & " order by highclasscode "

end if



	 Set rs = server.CreateObject("adodb.recordset")
	rs.activeconnection = application("connectionstring")
	rs.cursorLocation = adUseClient
	rs.cursorType = adOpenStatic
	rs.lockType = adLockOptimistic
	rs.source = sql
	rs.open

	If Not rs.eof Then
		Dim middleclassname : Set middleclassname = rs(1)
		Dim jan : Set jan = rs(2)
		Dim feb : Set feb = rs(3)
		Dim mar : Set mar = rs(4)
		Dim apr : Set apr = rs(5)
		Dim may : Set may = rs(6)
		Dim jun : Set jun = rs(7)
		Dim jul : Set jul = rs(8)
		Dim aug : Set aug = rs(9)
		Dim sep : Set sep = rs(10)
		Dim oct_ : Set oct_ = rs(11)
		Dim nov : Set nov = rs(12)
		Dim dec : Set dec = rs(13)
		Dim sum : Set sum = rs(14)
	End If

	sql = "select categoryidx, categoryname from wb_category where categorylvl is null"
	Dim cmd : set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adcmdText
	Dim rs2 : Set rs2 = cmd.execute
	If Not rs2.eof Then
		Dim highclasscode : Set highclasscode = rs2(0)
		Dim highclassname : Set highclassname = rs2(1)
	End If
	Set cmd = Nothing

	Dim s01 : s01 = 0
	Dim s02 : s02 = 0
	Dim s03 : s03 = 0
	Dim s04 : s04 = 0
	Dim s05 : s05 = 0
	Dim s06 : s06 = 0
	Dim s07 : s07 = 0
	Dim s08 : s08 = 0
	Dim s09 : s09 = 0
	Dim s10 : s10 = 0
	Dim s11 : s11 = 0
	Dim s12 : s12 = 0
	Dim total : total = 0


	Response.CacheControl  = "public"
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader  "Content-Disposition" , "attachment; filename="&cyear&"년 매체분류별  집행현황.xls"
'%>
<h2> <u>매체분류별 집행현황 ('<%=cyear%>)</u> </h2>
	  <table border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" width='2000'>
	  <thead bgcolor='#cccccc'>
		  <tr height='20'>
			<th >대분류</th>
			<th >중분류</th>
			<th >01월</th>
			<th >02월</th>
			<th >03월</th>
			<th >04월</th>
			<th >05월</th>
			<th >06월</th>
			<th >07월</th>
			<th >08월</th>
			<th >09월</th>
			<th >10월</th>
			<th >11월</th>
			<th >12월</th>
			<th >합계</th>
		  </tr>
		</thead>
		<tbody id='tbody'>
	<%
			Dim highclassname_old
			Do Until rs2.eof
			rs.Filter = "highclasscode="& highclasscode
			If rs.recordcount <> 0 Then
	%>

		<tr height='32'>
			<td  class="hd none" colspan='15' >
				<table  width='1030' border="1" style="table-layout:fixed;">
				<%
					Do Until rs.eof
				%>
					<tr height='30' <% If middleclassname = "소계" Then response.write "bgcolor=#ececec" End If %>>
						<td  width='100' class="hd none" ><%if highclassname_old <> highclassname then response.write highclassname %></td>
						<%highclassname_old = highclassname %>
						<td  width='100'  style='padding-left:3px;'><%=middleclassname%></td>
						<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(jan,0)%></td>
						<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(feb,0)%></td>
						<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(mar,0)%></td>
						<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(apr,0)%></td>
						<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(may,0)%></td>
						<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(jun,0)%></td>
						<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(jul,0)%></td>
						<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(aug,0)%></td>
						<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(sep,0)%></td>
						<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(oct_,0)%></td>
						<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(nov,0)%></td>
						<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(dec,0)%></td>
						<td  width='100' style='padding-right:10px;text-align:right;'><%=FormatNumber(sum,0)%></td>
					</tr>
				<%
				rs.movenext
				Loop

			%>
	<%
				End If
					rs2.movenext
				Loop
			%>
		</tbody>
</table>