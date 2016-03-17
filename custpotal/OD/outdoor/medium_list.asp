<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim gotopage : gotopage = request.querystring("gotopage")
	if gotopage = "" then gotopage = 1
	dim searchstring : searchstring = request("txtsearchstring")				' 매체명
	dim categoryidx : categoryidx = request("selcategory")					' 매체분류
	if categoryidx = "" then categoryidx = null
	dim custcode : custcode = request("selcustcode")							' 매체사코드
	if custcode = "" then custcode = null

	dim objrs, sql
	sql = "select count(mdidx) from dbo.wb_medium_mst; select m.mdidx, g.mdname, m.title, m.region, c.custname,  m.locate  from dbo.wb_medium_mst m left outer join dbo.sc_cust_temp c on m.custcode = c.custcode  left outer join dbo.vw_medium_category g on m.categoryidx = g.mdidx where m.title like '%"&searchstring&"%'  "
	if not isnull(custcode) then sql = sql & " and m.custcode = '" & custcode &"' "
	if not isnull(categoryidx)  then  sql = sql & " and  g.mgroupidx = "& categoryidx
	sql = sql & " order by g.ggroupidx , g.mdname"

	call get_recordset(objrs, sql)

	dim cnt : cnt = objrs(0)
	set objrs = objrs.nextrecordset
	dim mdidx, mediumname, categoryname, custname, locate, region
	if not objrs.eof then
		set mdidx = objrs("mdidx")					' 매체일련번호
		set mediumname = objrs("title")			' 매체명
		set categoryname = objrs("mdname")	' 분류명
		set custname = objrs("custname")			' 매체사명
		set locate = objrs("locate")					' 설치위치
		set region = objrs("region")					' 매체지역
	end if
%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  oncontextmenu="return false">
<form name="form1" method="post" action="">
<!--#include virtual="/od/top.asp" -->
  <table id="Table_01" width="1240"  border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/od/left_outdoor_menu.asp"--></td>
      <td height="65"><img src="/images/middle_navigater_outdoor.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td  height="600" align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="19">&nbsp;</td>
          </tr>
          <tr>
            <td height="17"><TABLE  width="100%">
            <TR>
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle"> 매체현황 </span></TD>
				<TD width="50%" align="right"><span class="navigator" > 옥외관리 &gt; 매체관리 &gt; 매체현황  </span></TD>
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
                  <td align="left" background="/images/bg_search.gif"> <%call get_middle_categoty(categoryidx)%>  <%call get_medium_custcode(custcode, null) %><input type="text" name="txtsearchstring"> <img src="/images/btn_search.gif" width="39" height="20" align="absmiddle" onClick="get_serch();" class="styleLink" ></td>
                  <td align="right" background="/images/bg_search.gif"><img src="/images/btn_medium_reg.gif" width="78" height="18" align="absmiddle" border="0" class="stylelink" onclick="get_medium_reg();"> </td>
                  <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="26" >&nbsp;</td>
          </tr>
          <tr>
            <td ><table width="1030" height="31" border="3" cellpadding="0" cellspacing="0" bordercolor="#8D652B">
                <tr>
                  <td><table width="1024" border="0" cellspacing="0" cellpadding="0" class="header">
                      <tr>
                        <td width="40" align="center" class="header">No</td>
                        <td width="3" align="center" ><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="120" align="center" >구분</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="64" align="center">지역</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="300" align="center" >매체명 </td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="150" align="center">매체사</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="350" align="center">설치위치</td>
                      </tr>
                  </table></td>
                </tr>
              </table>
                <table width="1024" height="31" border="0" cellpadding="0" cellspacing="0" bordercolor="#8D652B">
				<% do until objrs.eof %>
                  <tr onclick="go_medium_view('<%=mdidx%>')" class="styleLink">
                    <td width="40" height="31" align="center"><%=cnt%></td>
                    <td width="3" align="left">&nbsp;</td>
                    <td width="120" align="left" style="padding-left:10px;"> <%=categoryname%></td>
                    <td width="3">&nbsp;</td>
                    <td width="64" align="left" style="padding-left:10px;"> <%=region%></td>
                    <td width="3">&nbsp;</td>
                    <td width="300" align="left" style="padding-left:10px;"> &nbsp; <%=mediumname%></td>
                    <td width="3">&nbsp;</td>
                    <td width="150" align="left" style="padding-left:10px;"> &nbsp; <%=custname%></td>
                    <td width="3">&nbsp;</td>
                    <td width="350" align="left" style="padding-left:10px;">&nbsp; <%=locate%></td>
                  </tr>
                  <tr>
                    <td height="1" bgcolor="#E7E9E3" colspan="11"></td>
                  </tr>
				<%
						cnt = cnt - 1
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
<!--#include virtual="bottom.asp" -->
</form>
</body>
</html>
<script language="JavaScript">
<!--
	function go_medium_view(idx) {
		location.href="medium_view.asp?mdidx=" + idx + "&gotopage=<%=gotopage%>&searchstring=<%=searchstring%>&categoryidx=<%=categoryidx%>&custcode=<%=custcode%>";
	}

	function get_medium_reg() {
		var url = "pop_medium_reg.asp"
		var name = "pop_medium_reg";
		var opt = "width=540, height=372, resizable=no, scrollbars=no, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}

	function get_serch() {
		var frm = document.forms[0];
		frm.action = "medium_list.asp";
		frm.method = "post";
		frm.submit() ;
	}

	window.onload = function () {
	}
//-->
</script>