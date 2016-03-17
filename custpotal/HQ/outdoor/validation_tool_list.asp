<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	if cyear = "" then cyear = year(date)
	if cmonth = "" then cmonth = month(date)

	dim objrs, sql
	sql = "select c.tidx, g.mgroupname, m.title, c.[avg], c.class, t.validclass, sum(d.monthprice) as monthprice from dbo.wb_medium_mst m inner join dbo.wb_validation_class c on m.mdidx = c.mdidx inner join dbo.wb_contact_md m2 on m.mdidx = m2.mdidx inner join dbo.wb_contact_md_dtl d on m2.contidx = d.contidx and m2.sidx = d.sidx  inner join dbo.vw_medium_category g on m.categoryidx = g.mdidx left outer join dbo.wb_validation_tool t on c.tidx = t.tidx where d.cyear = '"&cyear&"' and d.cmonth = '"&cmonth&"' group by c.tidx, g.mgroupname, m.title, c.[avg], c.class, t.validclass order by mgroupname desc"
	call get_recordset(objrs, sql)

	dim tidx, mgroupname, title, avg, mclass, validclass, monthprice, cnt

	if not objrs.eof Then
		set tidx = objrs("tidx")
		set mgroupname = objrs("mgroupname")
		set title = objrs("title")
		set avg = objrs("avg")
		set mclass = objrs("class")
		set validclass = objrs("validclass")
		set monthprice = objrs("monthprice")
		cnt = objrs.recordcount
	end if
%>
<html>
<head>
<title>▒SK MARKETING EXCELLENT▒</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="../../style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="">
<!--#include virtual="/hq/top.asp" -->
  <table width="1240" height="652" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" id="Table_01">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/hq/left_outdoor_menu.asp"--></td>
      <td height="65"><img src="/images/middle_navigater_outdoor.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td height="600" align="left" valign="top"><table width="1030" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="19">&nbsp;</td>
          </tr>
          <tr>
            <td height="17"><TABLE  width="100%">
            <TR>
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle"> 옥외관리 &gt; 효용성 Tool </span></TD>
				<TD width="50%" align="right"><span class="navigator" > 옥외관리 &gt; 효용성 Tool</span></TD>
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
                  <td width="13" valign="top" ><img src="/images/bg_search_left.gif" width="13" height="35" ></td>
                  <td background="/images/bg_search.gif"> <img src="/images/icon_dot_search.gif" width="4" height="3" align="absmiddle"> 검색년월
				  <%call get_year(cyear)%> <%call get_month(cmonth)%>   <img src="/images/btn_search.gif" width="39" height="20" align="top" class="stylelink" onclick="go_search();">
				 </td>
                  <td  align="right" background="/images/bg_search.gif" >&nbsp;</td>
                  <td width="13" ><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="26" >&nbsp;</td>
          </tr>
          <tr>
            <td ><table width="1030"  border="3" cellpadding="0" cellspacing="0" bordercolor="#8D652B">
                <tr>
                  <td>
				  <table width="1024" border="0" cellspacing="0" cellpadding="0" class="header">
                      <tr>
                        <td width="40" align="center" height="30">No</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="74" align="center" >매체분류</td>
                        <td width="3" align="center" ><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="360" align="center">매체명</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="130" align="center">매체평점</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="130" align="center">매체등급</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="130" align="center">효용성등급</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="160" align="center">월광고료</td>
                      </tr>
                  </table></td>
                </tr>
              </table>
              <table width="1024" height="31" border="0" cellpadding="0" cellspacing="0" >
	     <%
			do until objrs.eof
		%>
                  <tr onClick="pop_vaildation_view('<%=mgroupname%>', <%=tidx%>,<%=monthprice%>)" class="styleLink" >
                    <td width="40" align="center"  height="30"><%=cnt%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="74" align="center"><%=mgroupname%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="360" align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=title%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="130" align="center"><%=avg%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="130" align="center"><%=mclass%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="130" align="center"><%=validclass%></td>
                    <td width="3"align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="160" align="right"><%If Not IsNull(monthprice) Then response.write formatnumber(monthprice,0) Else response.write "0"%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                  </tr>
                  <tr>
                    <td height="1" bgcolor="#E7E9E3" colspan="30"></td>
                  </tr>
				<%
						cnt = cnt -1
						objrs.movenext
					loop
					objrs.close
					set objrs = nothing
				%>
              </table></td>
          </tr>
      </table></td>
    </tr>
  </table>
<!--#include virtual="/bottom.asp" -->
</form>
<iframe id="ifrm" name="ifrm" width="0" height="0" frameborder="0" src="about:blank"></iframe>
</body>
</html>
<script language="JavaScript">
<!--
	function pop_vaildation_view(category, idx, price) {
		switch (category) {
			case "옥탑":
				var url = "pop_validation_neon.asp?tidx=" + idx+"&monthprice="+price;
				break;
			case "야립":
				var url = "pop_validation_board.asp?tidx=" + idx+"&monthprice="+price;
				break;
			case "LED":
				var url = "pop_validation_led.asp?tidx=" + idx+"&monthprice="+price;
				break;
			default :
				var url = "pop_validation_etc.asp?tidx=" + idx+"&monthprice="+price;
				break;
		}

		var name = "pop_validation_tool" ;
		var opt = "width=893,resizable=no, scrollbars=yes, status=no, , menubar=no, top=100, left=100";
		window.open(url, name, opt);
	}

	function pop_contact_reg() {
		var url = "pop_contact_reg.asp"
		var name = "pop_contact_reg";
		var opt = "width=540, height=577, resizable=no, scrollbars=no, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}

	function go_search() {
		var frm = document.forms[0];
		frm.action = "validation_tool_list.asp";
		frm.method = "post";
		frm.submit();
	}

	window.onload = function () {
	}
//-->
</script>

