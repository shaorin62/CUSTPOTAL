<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim contidx : contidx = request("contidx")
	dim gotopage : gotopage = request("gotopage")
	if gotopage = "" then gotopage = 1
	if contidx = "" then contidx = 6
	dim pagesize : pagesize = 8
	dim objrs, sql, searchstring, cyear, cmonth, custcode, custcode2
	dim midx, mtitle

	sql = "select title from dbo.wb_contact_mst where contidx = " & contidx
	call get_recordset(objrs, sql)

	dim title : title = objrs(0).value
	objrs.close

	sql = "select count(*) from dbo.wb_contact_monitor_mst m inner join (select pidx, max(filename) as filename from dbo.wb_contact_monitor_dtl d group by pidx) u on m.pidx = u.pidx where m.contidx="&contidx&" ;select top  "&pagesize&" m.pidx, acceptdate, filename from dbo.wb_contact_monitor_mst m inner join (select pidx, max(filename) as filename from dbo.wb_contact_monitor_dtl d group by pidx) u on m.pidx = u.pidx where m.contidx="&contidx&" and m.pidx not in (select top  "&(gotopage-1)*pagesize&" m.pidx  from dbo.wb_contact_monitor_mst m inner join (select pidx, max(filename) as filename from dbo.wb_contact_monitor_dtl d group by pidx) u on m.pidx = u.pidx where m.contidx="&contidx&" order by m.acceptdate desc) order by m.acceptdate desc"

	call get_recordset(objrs, sql)

	dim totalrecord : totalrecord = objrs(0).value

	set objrs = objrs.nextrecordset

	dim pidx, acceptdate, filename
	if not objrs.eof then
		set pidx = objrs("pidx")
		set acceptdate = objrs("acceptdate")
		set filename = objrs("filename")
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
      <td  align="left" valign="top"><table width="1030" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="19">&nbsp;</td>
          </tr>
          <tr>
            <td height="17"><TABLE  width="100%">
            <TR>
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle"> <%=title%> 모니터링</span></span></TD>
				<TD width="50%" align="right"><span class="navigator" > 옥외광고 &gt; 옥외 모니터링  &gt; <%=title%> 모니터링</span></TD>
            </TR>
            </TABLE>
			</td>
          </tr>
          <tr>
            <td height="15">&nbsp;</td>
          </tr>
          <tr>
            <td  ><table width="1030" height="35" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
                  <td width="50%" align="left" background="/images/bg_search.gif">&nbsp;</td>
                  <td width="50%" align="right" background="/images/bg_search.gif"><img src="/images/btn_monitor_reg.gif" width="100" height="18" align="absmiddle" border="0" onclick="pop_monitor_reg();" class="styleLink"> </td>
                  <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="26" >&nbsp;</td>
          </tr>
          <tr>
            <td valign="top"  height="600" >
			<!--  -->
			<table  border="0" cellpadding="0" cellspacing="5" >
			<tr>
			<%
				dim cnt : cnt = 1
				do until objrs.eof
			%>
				<td>
					<table border="0" cellpadding="9" cellspacing="0">
					<tr>
						<td bgcolor="#ECECEC" ><img src="/pds/monitor/<%=filename%>" width="220" height="140" border="0" onclick="pop_monitor_view(<%=pidx%>);" class="stylelink" alt="클릭하시면 해당 검수일의 사진목록 팝업이 나타납니다."></td>
					</tr>
					<tr>
						<td align="center" height="30" > 검수일자 : <%=acceptdate%></td>
					</tr>
					</table>
			<%
				if cnt MOD 4 = 0  then response.write "</td></tr><tr><tr><td height='1' bgcolor='#E7E9E3' colspan=4'></td></tr><tr>" else response.write "</td>"
				cnt = cnt + 1
				objrs.movenext
				loop
				dim page : page = "monitor_sub_list"
			%>
			</tr>
			</table>
			<table>
				<tr>
				<tr>
					<td  height="50"  valign="bottom" width="1002" class="pagesplit"><img src="/images/btn_list.gif" width="59" height="18" border="0" class="stylelink" onclick="history.back();"></td>
				</tr>
					<td  height="50" align="center" valign="bottom" width="1002" class="pagesplit"><%call pagesplit(totalrecord, gotopage, pagesize, page, searchstring, custcode, custcode2, cyear, cmonth, midx, mtitle)%></td>
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
</body>
</html>
<script language="JavaScript">
<!--
	function get_monitor_list() {
		location.href = "monitor_sub_list.asp";
	}

	function pop_monitor_reg() {
		var url = "pop_monitor_reg.asp?contidx=<%=contidx%>"
		var name = "pop_monitor_reg";
		var opt = "width=540, height=467, resizable=no, scrollbars=no, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}

	function pop_monitor_view(idx) {
		var url = "pop_monitor_view.asp?pidx="+idx
		var name = "pop_monitor_view";
		var opt = "width=540, height=600, resizable=no, scrollbars=yes, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}
//-->
</script>