<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	if cyear = "" then cyear = year(date)
	if cmonth = "" then cmonth = month(date)
	if Len(cmonth) = 1 then cmonth = "0" & cmonth
	dim c_date : c_date = DateSerial(cyear, cmonth, "01")
	c_date = DateAdd("d", -1, c_date)

	dim medcode : medcode = request("selcustcode3")
	if medcode = "" then medcode = request.cookies("custcode")

	dim objrs, sql
	sql = "select m.contidx , m2.sidx, d.idx, m.title, m.startdate, m.enddate , m2.locate, d.side, d.standard, d.quality , m2.region, c.custname as deptname, c2.custname as custname from dbo.wb_contact_mst m inner join dbo.wb_contact_md m2 on m.contidx = m2.contidx  inner join dbo.wb_contact_md_dtl d on m2.sidx = d.sidx inner join dbo.wb_contact_md_dtl_account a on d.idx = a.idx inner join dbo.sc_cust_temp c on c.custcode = m.custcode inner join dbo.sc_cust_temp c2 on c.highcustcode = c2.custcode inner join dbo.sc_cust_temp c3 on m2.medcode = c3.custcode where m2.medcode like '"&medcode&"%' and a.cyear = '"&cyear&"' and a.cmonth = '"&cmonth&"' and m.canceldate >= '" & c_date & "'  order by m.contidx desc	"

	call get_recordset(objrs, sql)

	dim contidx, startdate, enddate, sidx, idx, title, side, locate, standard, quality, qty, photo, custname, deptname, canceldate, region

	if not objrs.eof then
		set contidx = objrs("contidx")
		set sidx = objrs("sidx")
		set idx = objrs("idx")
		set title = objrs("title")
		set title = objrs("title")
		set startdate = objrs("startdate")
		set enddate = objrs("enddate")
		set locate = objrs("locate")
		set side = objrs("side")
		set standard = objrs("standard")
		set quality = objrs("quality")
		set region = objrs("region")
		set custname = objrs("custname")
		set deptname = objrs("deptname")
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
    <td height="24" background="/images/pop_top.gif" valign="top" >
	<% if request.cookies("class") = "M" then %><table width="700"  border="0" align="right" cellpadding="0" cellspacing="0" height="60">
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
    </table>
	<% else %><table width="700"  border="0" align="right" cellpadding="0" cellspacing="0" height="60">
      <tr style="padding-top:10">
        <td height="24">&nbsp;</td>
      </tr>
    </table>
	<% end if %></td>
  </tr>
  <tr>
    <td height="24">&nbsp;</td>
  </tr>
  <tr>
    <td height="17"  align="center"><table width="1030" border="0" cellpadding="0" cellspacing="0" >
    <tr>
		<td><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle"><%=request.cookies("custname")%> 광고매체 리포트 현황 </span></td>
    </tr>
    </table></td>
  </tr>
  <tr>
    <td height="27">&nbsp;</td>
  </tr>
  <tr>
    <td valign="top" class="bdpdd" align="center">

			<table width="1030" height="35" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="13" valign="top" ><img src="/images/bg_search_left.gif" width="13" height="35" ></td>
                  <td background="/images/bg_search.gif"> <img src="/images/icon_dot_search.gif" width="4" height="3" align="absmiddle"> 검색년월
				  <%call get_year(cyear)%> <%call get_month(cmonth)%> &nbsp;<% if request.cookies("class") <> "M" then  call get_custcode_custcode3(medcode, null)%> &nbsp;     <img src="/images/btn_search.gif" width="39" height="20" align="top" class="stylelink" onClick="go_search();"> </td>
                  <td  align="right" background="/images/bg_search.gif" >&nbsp; </td>
                  <td width="13" ><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
        </table>

	</td>
  </tr>
  <tr>
    <td height="15">&nbsp;</td>
  </tr>
  <tr>
    <td height="600"  align='center'  valign="top">
	<!-- -->
				  <table width="1024" border="0" cellspacing="1" cellpadding="0" class="header" >
                      <tr height="30"  bgcolor="#ECECEC">
                        <td width="300" align="center" >매체명</td>
                        <td width="100" align="center">지역</td>
                        <td width="50" align="center">면</td>
                        <td width="300" align="center">규격 / 재질 </td>
                        <td width="100" align="center">광고주</td>
                        <td width="100" align="center">사업부서</td>
                        <td width="80" align="center">&nbsp;</td>
                      </tr>
                </tr>
				  <% do until objrs.eof %>
                      <tr height="30"  bgcolor="#FFFFFF">
                        <td  align="left" ><%=title%> <%if not isnull(locate) then response.write "<br >" & locate %></td>
                        <td  align="center">&nbsp;<%=region%></td>
                        <td  align="center">&nbsp;<%=side%></td>
                        <td  align="left">&nbsp;<%=standard%><%if not isnull(quality) then response.write " / " & quality else response.write quality%>&nbsp;</td>
                        <td  align="left">&nbsp;<%=custname%></td>
                        <td  align="left">&nbsp;<%=deptname%></td>
                        <td  align="center"><img src="/images/btn_photo_reg.gif" width="78" height="18"  vspace="5" class="stylelink" onClick="pop_photo_mng(<%=idx%>, '<%=cyear%>','<%=cmonth%>');"></td>
                      </tr>
						<tr>
							<td colspan="14" bgcolor="#E7E7DE" height="1"></td>
						</tr>
					 <%
						objrs.movenext
						loop
					 %>
        </table>
	<!--  -->
	</td>
  </tr>
  <tr>
    <td height="24">&nbsp; </td>
  </tr>
  <tr>
    <td height="24"><img src="/images/pop_bottom.gif" width="1240" height="71" align="absmiddle"></td>
  </tr>
</table>
</form>
</body>
</html>
<script language="JavaScript">
<!--
	function go_search() {
		var frm = document.forms[0];
		frm.action = "/med/";
		frm.method = "post";
		frm.submit();
	}

	function pop_photo_mng(idx, cyear, cmonth) {
		var url = "pop_contact_photo_reg.asp?idx="+idx+"&cyear="+cyear+"&cmonth="+cmonth;
		var name = "pop_contact_photo_reg";
		var opt = "width=744, height=600, resizable=no, scrollbars=yes, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}

	window.onload = function () {
		self.focus();
	}

	function pop_photo_mng1() {
		var url = "reg.asp";
		var name = "pop_photo_mng";
		var opt = "width=540, height=600, resizable=no, scrollbars=yes, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}
//-->
</script>