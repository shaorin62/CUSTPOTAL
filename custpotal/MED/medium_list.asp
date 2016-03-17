<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	if cyear = "" then cyear = year(date)
	if cmonth = "" then cmonth = month(date)
	if len(cmonth)=1 then cmonth = "0"&cmonth
	dim custcode3 : custcode3 = request("selcustcode3")
	if custcode3 = "" then custcode3 = request.cookies("custcode")

	dim objrs, sql
	'sql = "select m.contidx, m.title as mtitle, m.startdate, m.enddate, d.side, m2.standard, d.quality, m2.qty, m.custcode , d.cyear, d.cmonth, COALESCE(photo_1, photo_2, photo_3, photo_4, null) as photo, c.custname as custname2, c2.custname, c3.custname as custname3 from dbo.wb_contact_mst m inner join dbo.wb_contact_md m2 on m.contidx = m2.contidx inner join dbo.wb_contact_md_dtl d on m2.contidx = d.contidx and m2.sidx = d.sidx inner join dbo.wb_medium_mst md on m2.mdidx = md.mdidx  inner join dbo.sc_cust_temp c on m.custcode = c.custcode inner join dbo.sc_cust_temp c2 on c.highcustcode = c2.custcode	inner join dbo.sc_cust_temp c3 on c3.custcode = md.custcode	where md.custcode like '"&custcode3&"%' and d.cyear = '"&cyear&"' and d.cmonth = '"&cmonth&"'  "

	sql = "select m.contidx, m.title, m.startdate, m.enddate, m2.locate, d.idx, d.side, d.standard, d.quality, COALESCE(photo_1, photo_2, photo_3, photo_4, null) as photo, c.custname as custname2, c2.custname from dbo.wb_contact_mst m left outer join dbo.wb_contact_md m2 on m.contidx = m2.contidx left outer join dbo.wb_contact_md_dtl d on m2.sidx = d.sidx left outer join dbo.wb_contact_md_dtl_account a on d.idx = a.idx and a.cyear = '" & cyear &"' and a.cmonth = '" & cmonth & "' left outer join dbo.sc_cust_temp c on m.custcode = c.custcode left outer join dbo.sc_cust_temp c2 on c2.custcode = c.highcustcode  where m2.medcode = '" & custcode3 & "' order by m.enddate desc"

	call get_recordset(objrs, sql)

	dim contidx, startdate, enddate, sidx, locate ,mtitle,  title, side, standard, quality, qty, photo, custname2, custname, canceldate, idx

	if not objrs.eof then
		set contidx = objrs("contidx")
		set locate = objrs("locate")
		set title = objrs("title")
		set side = objrs("side")
		set standard = objrs("standard")
		set quality = objrs("quality")
		set photo = objrs("photo")
		set startdate = objrs("startdate")
		set enddate = objrs("enddate")
		set custname = objrs("custname")
		set custname2 = objrs("custname2")
		set idx = objrs("idx")
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
		<td><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle"><%=request.cookies("custname")%> 매체 사진관리 </span></td>
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
				  <%call get_year(cyear)%> <%call get_month(cmonth)%> &nbsp;<% if request.cookies("class") = "A" then  call get_custcode_custcode3(custcode3, null)%> &nbsp;     <img src="/images/btn_search.gif" width="39" height="20" align="top" class="stylelink" onClick="go_search();"> </td>
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
	<table width="1030"  border="3" cellpadding="0" cellspacing="0" bordercolor="#8D652B" >
                <tr>
                  <td>
				  <table width="1024" border="0" cellspacing="0" cellpadding="0" class="header">
                      <tr>
                        <td width="150" align="center" height="30">&nbsp;</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="300" align="center" >매체명</td>
                        <td width="3" align="center" ><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="50" align="center">면</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="125" align="center">규격 / 재질 </td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="125" align="center">계약기간</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="200" align="center">광고주</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="100" align="center">사업부서</td>
                      </tr>
                  </table>
				  </td>
                </tr>
				<!-- 데이터 처리 구간 -->
        </table>
				  <table width="1024" border="0" cellspacing="3" cellpadding="0" class="header">
				  <% do until objrs.eof %>
                      <tr>
                        <td width="150" align="center" height="100" bgcolor="#E7E7DE" valign="middle"><span class="stylelink" onclick="pop_photo_mng(<%=idx%>,'<%=cyear%>', '<%=cmonth%>');"><img src=<%if not isnull(photo) then Response.write "/pds/media/"&photo else response.write "/images/noimage.gif"%> width="130" height="80"></span></td>
                        <td width="3" align="center">&nbsp;</td>
                        <td width="300" align="center" ><%= locate %><p><%=title%></td>
                        <td width="3" align="center" >&nbsp</td>
                        <td width="50" align="center"><%=side%></td>
                        <td width="3" align="center">&nbsp</td>
                        <td width="125" align="center"><%=standard%> <p> <%=quality%></td>
                        <td width="3" align="center">&nbsp</td>
                        <td width="125" align="center"><%=startdate%> <p> <%=enddate%></td>
                        <td width="3" align="center">&nbsp</td>
                        <td width="200" align="center"><%=custname%></td>
                        <td width="3" align="center">&nbsp</td>
                        <td width="100" align="center"><%=custname2%></td>
                      </tr>
						<tr>
							<td colspan="14" bgcolor="#E7E7DE" height="1"></td>
						</tr>
					 <%
						objrs.movenext
						loop
					 %>
						<tr>
							<td colspan="14"   height="30"> *사진 등록 후 이미지가 보이지 않을 경우 검색버튼을 눌러주십시오. </td>
						</tr>
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
		frm.action = "medium_list.asp";
		frm.method = "post";
		frm.submit();
	}

	function pop_photo_mng(idx, cyear, cmonth) {
		var url = "pop_photo_mng.asp?idx="+idx+"&cyear="+cyear+"&cmonth="+cmonth;
		var name = "pop_photo_mng";
		var opt = "width=540, height=600, resizable=no, scrollbars=yes, status=yes, left=100, top=100";
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