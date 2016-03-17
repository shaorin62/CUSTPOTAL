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

	dim pagesize : pagesize = 5

	dim objrs, sql
	sql = "select count(*) from dbo.wb_contact_mst m left outer join dbo.wb_contact_monitor_mst m2 on m.contidx = m2.contidx left outer join dbo.vw_monitor_filename f on m2.pidx = f.pidx  left outer join dbo.sc_cust_temp c on m.custcode = c.custcode where c.highcustcode like '"&custcode&"%' and m.startdate <= '"&e_date&"' and m.enddate > '"&s_date&"' and m.canceldate >= '"&s_date&"'; select  top "&pagesize&" m.contidx, m.title, m.startdate, m.enddate, f.filename, m2.nextacceptdate from dbo.wb_contact_mst m left outer join dbo.wb_contact_monitor_mst m2 on m.contidx = m2.contidx left outer join dbo.vw_monitor_filename f on m2.pidx = f.pidx  left outer join dbo.sc_cust_temp c on m.custcode = c.custcode  where c.highcustcode like '"&custcode&"%' and m.startdate <= '"&e_date&"' and m.enddate > '"&s_date&"' and m.canceldate >= '"&s_date&"' and m.contidx not in ( select  top "&(gotopage-1) * pagesize&" m.contidx  from dbo.wb_contact_mst m left outer join dbo.wb_contact_monitor_mst m2 on m.contidx = m2.contidx left outer join dbo.vw_monitor_filename f on m2.pidx = f.pidx left outer join dbo.sc_cust_temp c on m.custcode = c.custcode  where c.highcustcode like '"&custcode&"%' and m.startdate <= '"&e_date&"'  and m.enddate > '"&s_date&"' and m.canceldate >=  '"&s_date&"'   order by m.contidx desc ) order by m.contidx desc"

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
		end if
%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  oncontextmenu="return false">
<form name="F_DATE">
<!--#include virtual="/od/top.asp" -->
  <table id="Table_01" width="1240" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/od/left_outdoor_menu.asp"--></td>
      <td height="65"><img src="/images/middle_navigater_outdoor.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td  align="left" valign="top"  height="600"><table width="1030" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="19">&nbsp;</td>
          </tr>
          <tr>
            <td height="17"><TABLE  width="100%">
            <TR>
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle"> 옥외 모니터링</span></TD>
				<TD width="50%" align="right"><span class="navigator" > 옥외광고 &gt; 옥외 모니터링 </span></TD>
            </TR>
            </TABLE>
			</td>
          </tr>
          <tr>
            <td height="15">&nbsp;</td>
          </tr>
          <tr>
            <td ><table width="1030" height="35" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
                  <td  align="left" background="/images/bg_search.gif"> <img src="/images/icon_dot_search.gif" width="4" height="3" align="absmiddle"> 검색년월
				  <%call get_year(cyear)%> <%call get_month(cmonth)%> &nbsp; <img src="/images/icon_dot_search.gif" width="4" height="3" align="absmiddle"> 검색조건  <%call get_custcode_mst(custcode, null, null)%><span id="custcode2"><%'call get_blank_select("사업부를 선택하세요", 207)%></span><!-- <input type="text" name="txtsearchstring"> --> <img src="/images/btn_search.gif" width="39" height="20" align="top" class="stylelink" onclick="go_search();"></td>
                  <td align="right" background="/images/bg_search.gif">&nbsp;</td>
                  <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="26" >&nbsp;</td>
          </tr>
          <tr>
            <td valign="top"  height="600">
			<!--  -->

				<table border="0" cellpadding="0" cellspacing="0" width="1030">
			  <%	do until objrs.eof %>
				<tr>
					<td class="phbd"><img src="<%if isnull(filename) then Response.write "/images/noimage.gif" else response.write "/pds/monitor/" & filename %>" width="140" height="90" border="0" onclick='get_monitor_list(<%=contidx%>);' class='stylelink'></td>
					<td valign="top">
						<table border="0" cellspacing="3" cellpadding="0" bgcolor="#8A652B" >
						<tr>
							<td bgcolor="#FFFFFF" class="phhd"><b><%=title %></b></td>
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

				dim page : page = "monitor_list"
			%>
				<tr>
					<td  height="20" align="center" colspan="4" class="pagesplit">
<%
		dim blockpage : blockpage = Int((totalrecord-1)/pagesize)+1
		dim blockno : blockno = Int((gotopage-1)/pagesize)*pagesize+1


		if blockno <> 1 then
			response.write " <a href='monitor_list.asp?gotopage="&blockno-pagesize&"&cyear="&cyear&"&cmonth="&cmonth&"&selcustcode="&custcode&"'><img src='/images/icon_prev.gif' width='5' height='8' border='0' align='absmiddle' vspace='5'></a> "
		end if

		dim intLoop
		for intLoop = blockno to ((blockno-1)+pagesize)
			if intLoop <= int(blockpage) then
				if intLoop = int(gotopage) then
					response.write " "& intLoop & " "
				else
					response.write " <a href='monitor_list.asp?gotopage="&intLoop&"&cyear="&cyear&"&cmonth="&cmonth&"&selcustcode="&custcode&"' class='pagesplit'>"&intLoop&"</a> "
				end if
			end if
		next
		if intLoop <= blockpage then
			response.write "<a href='monitor_list.asp?gotopage="&blockno+pagesize&"&cyear="&cyear&"&cmonth="&cmonth&"&selcustcode="&custcode&"'><img src='/images/icon_next.gif' width='5' height='8' border='0' align='absmiddle' vspace='5'></a> "
		end if
%>
</td>
				</tr>
				</table>
			<!--  -->
			</td>
          </tr>
      </table>
	  </td>
    </tr>
  </table>
<!--#include virtual="bottom.asp" -->
</form>
<!-- <iframe id="ifrm" name="ifrm" width="0" height="0" frameborder="0" src="about:blank"></iframe> -->
</body>
</html>
<script language="JavaScript">
<!--
	function get_monitor_list(idx) {
		location.href = "monitor_sub_list.asp?contidx="+idx;
	}

	function go_page(url) {
//		var frm = document.getElementById("ifrm") ;
//		var custcode = document.forms[0].selcustcode.options[document.forms[0].selcustcode.selectedIndex].value ;
//		var custcode2 = "<%=custcode2%>" ;
//		frm.src="/inc/frm_code.asp?custcode="+custcode +"&custcode2="+custcode2;
	}

	function go_search() {
		var frm = document.forms[0];
		frm.action = "monitor_list.asp";
		frm.method = "post";
		frm.submit();
	}

	window.onload = function () {
		go_page("");
	}
//-->
</script>