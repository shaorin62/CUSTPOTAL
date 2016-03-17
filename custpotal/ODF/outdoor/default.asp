<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim gotopage : gotopage = request("gotopage")
	if gotopage = "" then gotopage = 1
	dim searchstring : searchstring = request("searchstring")
	dim midx, mtitle

	dim custcode : custcode = request("selcustcode")
	dim custcode2 : custcode2 = request("selcustcode2")

	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	if cyear = "" then cyear = year(date)
	if cmonth = "" then cmonth = month(date)
	dim s_date : s_date = dateserial(cyear, cmonth, "01")
	dim e_date : e_date = DateAdd("d", -1, DateAdd("m", 1, s_date))

	response.write "custcode : " & custcode
	dim pagesize : pagesize = 5

	dim objrs, sql
	sql = "select count(*) from dbo.wb_contact_mst m inner join dbo.sc_cust_dtl c on m.custcode = c.oldcustcode inner join dbo.sc_cust_hdr h on c.highcustcode = h.highcustcode left outer join (select m.contidx, max(filename) as filename, nextacceptdate from dbo.wb_contact_monitor_mst m inner join dbo.wb_contact_monitor_dtl d on m.pidx = d.pidx group by m.contidx, m.nextacceptdate) u on m.contidx = u.contidx where h.oldcustcode like '"&custcode&"%' and c.custcode like '" & custcode2 &"%'  and m.title like '%"&searchstring&"%' and m.startdate <= '"&s_date&"' and m.enddate >= '"&e_date&"' and m.cancel = 0; "
'	sql = "select count(*) from dbo.wb_contact_mst  m inner join dbo.sc_cust_temp c on m.custcode = c.custcode left outer join (select m.contidx, max(filename) as filename, nextacceptdate from dbo.wb_contact_monitor_mst m inner join dbo.wb_contact_monitor_dtl d on m.pidx = d.pidx group by m.contidx, m.nextacceptdate) u on m.contidx = u.contidx where c.highcustcode like '" & custcode &"%' and c.custcode like '" & custcode2 &"%' and m.title like '%" & searchstring &"%' and m.startdate <= '" & e_date &"' and m.enddate >= '"&s_date&"' and m.cancel = 0; "
	sql = sql & "select TOP "&pagesize&" m.contidx, m.title, m.startdate, m.enddate, m.canceldate, u.filename, u.nextacceptdate 	from dbo.wb_contact_mst m inner join dbo.sc_cust_dtl c on m.custcode = c.oldcustcode inner join dbo.sc_cust_hdr h on c.highcustcode = h.highcustcode 	left outer join 	(select m.contidx, max(filename) as filename, nextacceptdate from dbo.wb_contact_monitor_mst m inner join dbo.wb_contact_monitor_dtl d on m.pidx = d.pidx group by m.contidx, m.nextacceptdate) u on m.contidx = u.contidx 	where h.oldcustcode like '" & custcode &"%' and c.custcode like '" & custcode2 &"%' and m.title like '%" & searchstring &"%' and m.startdate <= '" & e_date &"' and m.enddate >= '"&s_date&"' and  m.contidx not in (select top "& (gotopage-1) * pagesize &" m.contidx 	from dbo.wb_contact_mst m 	inner join dbo.sc_cust_temp c on m.custcode = c.custcode 	left outer join 	(select m.contidx, max(filename) as filename, nextacceptdate from dbo.wb_contact_monitor_mst m inner join dbo.wb_contact_monitor_dtl d on m.pidx = d.pidx  group by m.contidx, m.nextacceptdate) u on m.contidx = u.contidx  order by m.contidx desc) and m.cancel =0 order by m.contidx desc"
	response.write sql
	call get_recordset(objrs, sql)

	dim totalrecord : totalrecord = objrs(0).value

	set objrs = objrs.nextrecordset

	dim contidx, title, startdate, enddate, filename, nextacceptdate, canceldate
		if not objrs.eof then
			set contidx = objrs("contidx")
			set title = objrs("title")
			set startdate = objrs("startdate")
			set enddate = objrs("enddate")
			set filename = objrs("filename")
			set nextacceptdate = objrs("nextacceptdate")
			set canceldate = objrs("canceldate")
		end if
%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"   >
<!-- <body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"   oncontextmenu="return false"> -->
<form>
<table width="1240" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td height="24" background="/images/pop_top.gif" valign="top" ><table width="700"  border="0" align="right" cellpadding="0" cellspacing="0" height="60">
      <tr style="padding-top:10">
        <td>&nbsp;</td>
        <td width="244" align="right" valign="top" ><span class="log">&nbsp;<%=request.cookies("custname")%></span> &nbsp;</td>
        <td width="1" valign="top"><img src="/images/top_vline_bg.gif" width="1" height="32"></td>
        <td width="104" align="right" valign="top" ><span class="log">&nbsp;<%=request.cookies("userid")%></span> &nbsp;</td>
        <td width="1" valign="top" ><img src="/images/top_vline_bg.gif" width="1" height="32"></td>
        <td width="164" align="right" valign="top" ><span class="log"><%=formatdatetime(request.cookies("logtime"))%>&nbsp;&nbsp;</span></td>
        <td width="1" valign="top"><img src="/images/top_vline_bg.gif" width="1" height="32"></td>
        <td width="85" align="center"valign="top" ><A HREF="/Log_out.asp"><img src="/images/btn_logout.gif" width="64" height="19" border="0"></A></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="24">&nbsp;</td>
  </tr>
  <tr>
    <td height="17"  align="center"><table width="1030" border="0" cellpadding="0" cellspacing="0" >
    <tr>
		<td><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle">옥외 모니터링 관리 </span></td>
    </tr>
    </table></td>
  </tr>
          <tr>
            <td height="15">&nbsp;</td>
          </tr>
          <tr>
            <td align="center"><table width="1030" height="35" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
                  <td  align="left" background="/images/bg_search.gif"> <img src="/images/icon_dot_search.gif" width="4" height="3" align="absmiddle"> 검색년월
				  <%call get_year(cyear)%> <%call get_month(cmonth)%> &nbsp; <img src="/images/icon_dot_search.gif" width="4" height="3" align="absmiddle"> 검색조건  <%call get_use_custcode(custcode, null, "default.asp")%><span id="custcode2"><%call get_blank_select("사업부를 선택하세요", 207)%></span><input type="text" name="txtsearchstring"> <img src="/images/btn_search.gif" width="39" height="20" align="top" class="stylelink" onclick="go_search();"></td>
                  <td align="right" background="/images/bg_search.gif">&nbsp;</td>
                  <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="26" >&nbsp;</td>
          </tr>
          <tr>
            <td valign="top"  height="600" align="center">
			<!--  -->

				<table border="0" cellpadding="0" cellspacing="0" width="1030">
			  <%	do until objrs.eof %>
				<tr>
					<td class="phbd"><a href="#" onclick='get_monitor_list(<%=contidx%>);return false;'><img src="<%if isnull(filename) then Response.write "/images/noimage.gif" else response.write "/pds/monitor/" & filename %>" width="140" height="90" border="0"  class='stylelink'></a></td>
					<td valign="top">
						<table border="0" cellspacing="3" cellpadding="0" bgcolor="#8A652B" >
						<tr>
							<td bgcolor="#FFFFFF" class="phhd"><b><%=title%></b></td>
						</tr>
						</table>
						<table border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
						</tr>
						<tr>
							<td class="tdhd" width="140"> 계약기간 </td>
							<td class="tbd" width="268"><%=startdate%> ~ <%=enddate%> </td>
							<td  class="tdhd" width="140"> 최종 모니터링 일자 </td>
							<td class="tbd" width="268"><%=nextacceptdate%></td>
						</tr>
						<tr>
							<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
						</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td colspan="4" height="10"></td>
				</tr>
			<%
				objrs.movenext
				loop

				dim page : page = "default"
			%>
				<tr>
					<td  height="20" align="center" colspan="4" class="pagesplit"><%call pagesplit(totalrecord, gotopage, pagesize, page, searchstring, custcode, custcode2, cyear, cmonth, midx,mtitle)%></td>
				</tr>
				</table>
			<!--  -->
			</td>
          </tr>
      </table>
	  </td>
    </tr>
  </table><img src="/images/pop_bottom.gif" width="1240" height="71" align="absmiddle">
</form>
<iframe id="ifrm" name="ifrm" width="0" height="0" frameborder="0" src="about:blank"></iframe>
</body>
</html>
<script language="JavaScript">
<!--
	function get_monitor_list(idx) {
		location.href = "monitor_sub_list.asp?contidx="+idx+"&gotopage=<%=gotopage%>&selcustcode=<%=custcode%>&selcustcode2=<%=custcode2%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>&searchstring=<%=searchstring%>";
	}

	function go_page(url) {
		var frm = document.getElementById("ifrm") ;
		var custcode = document.forms[0].selcustcode.options[document.forms[0].selcustcode.selectedIndex].value ;
		var custcode2 = "<%=custcode2%>" ;
		frm.src="/inc/frm_code.asp?custcode="+custcode +"&custcode2="+custcode2;
	}

	function go_search() {
		var frm = document.forms[0];
		frm.action = "default.asp";
		frm.method = "post";
		frm.submit();
	}

	window.onload = function () {
		go_page("");
	}
//-->
</script>