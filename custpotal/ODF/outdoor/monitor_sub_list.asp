<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim contidx : contidx = request("contidx")
	Dim custcode : custcode = request("selcustcode")
	Dim custcode2 : custcode2 = request("selcustcode2")
	Dim cyear : cyear = request("cyear")
	Dim cmonth : cmonth = request("cmonth")
	Dim searchstring : searchstring = request("searchstring")

	dim gotopage : gotopage = request("gotopage")
	if gotopage = "" then gotopage = 1
	if contidx = "" then contidx = 6
	dim pagesize : pagesize = 8
	dim midx, mtitle, objrs

	response.write "custcode  : "  & custcode

	sql = "select title from dbo.wb_contact_mst where contidx = " & contidx
	call get_recordset(objrs, sql)

	dim title : title = objrs(0).value
	objrs.close

	Dim sql : sql = "select count(*) from dbo.wb_contact_monitor_mst m inner join (select pidx, max(filename) as filename from dbo.wb_contact_monitor_dtl d group by pidx) u on m.pidx = u.pidx where m.contidx="&contidx&" ;select top  "&pagesize&" m.pidx, acceptdate, filename from dbo.wb_contact_monitor_mst m inner join (select pidx, max(filename) as filename from dbo.wb_contact_monitor_dtl d group by pidx) u on m.pidx = u.pidx where m.contidx="&contidx&" and m.pidx not in (select top  "&(gotopage-1)*pagesize&" m.pidx  from dbo.wb_contact_monitor_mst m inner join (select pidx, max(filename) as filename from dbo.wb_contact_monitor_dtl d group by pidx) u on m.pidx = u.pidx where m.contidx="&contidx&" order by m.acceptdate desc) order by m.acceptdate desc"

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
<title>▒▒ SK M&C | Media Management System ▒▒</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"   oncontextmenu="return false">
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
		<td><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle"><%=title%> 사진 관리 </span></td>
    </tr>
    </table></td>
  </tr>
          <tr>
            <td height="15">&nbsp;</td>
          </tr>
          <tr>
            <td  align="center"><table width="1030" height="35" border="0" cellpadding="0" cellspacing="0">
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
            <td valign="top"  height="600"  align="center">
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
					<td  height="50"  valign="bottom" width="1002" class="pagesplit"><a href="#" onclick="goto_main();return false;"><img src="/images/btn_list.gif" width="59" height="18" border="0" class="stylelink" ></a></td>
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
  </table><img src="/images/pop_bottom.gif" width="1240" height="71" align="absmiddle">
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

	function goto_main() {
		location.href = "default.asp?gotopage=<%=gotopage%>&selcustocde=<%=custcode%>&selcustcode2=<%=custcode2%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>&searchstring=<%=searchstring%>"
	}
//-->
</script>