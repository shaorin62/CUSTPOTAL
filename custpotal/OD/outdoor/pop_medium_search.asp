<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim searchstring : searchstring = request("txtsearchstring")
	dim searchcategory : searchcategory = request("selcategory")
	if searchcategory = "" then searchcategory = 0

	dim objrs, sql
	sql = "select d.sidx, m.mdidx, m.title, d.side, d.standard, d.quality, m.locate, c.mdidx as categoryidx, c.mdname as categoryname, m.unit, m.custcode, t.custname, d.unitprice, m.regionmemo, m.mediummemo, m.map " &_
			" from dbo.wb_medium_mst m inner join dbo.wb_medium_dtl d on m.mdidx = d.mdidx " &_
			" left join dbo.vw_medium_category c on m.categoryidx = c.mdidx  " &_
			" left  join dbo.sc_cust_temp t on m.custcode = t.custcode " &_
			" where m.title like '%" & searchstring &"%' or c.mgroupidx = " & searchcategory & " order by title "
			'response.write sql
	call get_recordset(objrs, sql)

	dim mdidx, sidx, title, side, standard, quality, locate, categoryidx, categoryname, unit, custcode, unitprice, custname, regionmemo, mediummemo, map
	if not objrs.eof then
		set mdidx = objrs("mdidx")
		set sidx = objrs("sidx")
		set title = objrs("title")
		set categoryidx = objrs("categoryidx")
		set categoryname = objrs("categoryname")
		set side = objrs("side")
		set unit = objrs("unit")
		set standard = objrs("standard")
		set quality = objrs("quality")
		set locate = objrs("locate")
		set custcode = objrs("custcode")
		set unitprice = objrs("unitprice")
		set custname = objrs("custname")
		set regionmemo = objrs("regionmemo")
		set mediummemo = objrs("mediummemo")
		set map = objrs("map")
	end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<title>▒▒ SK M&C | Media Management System ▒▒  </title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<link href="/style.css" rel="stylesheet" type="text/css">
</head>

<body  oncontextmenu="return false">
<form>
	<table  width="718"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="22"><img src="/images/pop_left_top_bg.gif" width="22" height="102"></td>
    <td background="/images/pop_center_top.gif" align="left" valign="top" style="padding-left:2px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> 매체 검색 <p><%'call get_middle_categoty()%> <span style="font-size:12px;color:#333333">검색할 매체명을 입력하세요. </font><input type="text" name="txtsearchstring" class="kor" > <img src="/images/btn_search.gif" width="39" height="20" align="absmiddle" class="stylelink" onClick="get_medium_category();"> </td>
    <td width="121"><img src="/images/pop_right_top_bg.gif" width="121" height="102"></td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" width="718" >
  <tr>
    <td width="22"><img src="/images/pop_left_body_top.gif" width="22" height="16"></td>
    <td background="/images/pop_center_top_bg.gif"> &nbsp;</td>
    <td width="23"><img src="/images/pop_right_body_top.gif" width="23" height="16"></td>
  </tr>
  <tr>
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
    <td>
<!-- -->
  <table cellpadding="0" cellspacing="0">
          <tr >
            <td class="thd2" width="200" style="padding-left:5px;">매체명</td>
            <td class="thd2" width="30">면</td>
            <td class="thd2" width="170">규격(M)/재질</td>
            <td class="thd2" width="318">설치위치</td>
          </tr>
          <tr>
            <td colspan="4">
			<div style="overflow-y:scroll;height:510">
			<table border="0" border="0" cellpadding="0" cellspacing="0">
			  <% do until objrs.eof %>
              <tr height="30" onclick="put_medium_data('<%=mdidx%>','<%=sidx%>', '<%=replace(title, "'", "/")%>','<%=categoryidx%>','<%=categoryname%>', '<%=side%>', '<%=replace(standard, """", "-")%>', '<%=quality%>','<%=custcode%>', '<%=unit%>','<%=locate%>','<%=unitprice%>')" class="stylelink">
                <td width="200" style="padding-left:5px;"><%=title%>&nbsp;</td>
                <td width="30" style="padding-left:5px;"><%=side%>&nbsp;</td>
                <td width="170"><%=standard%> <% if not isnull(quality) then response.write "/" &quality &"" %> &nbsp;</td>
                <td width="300"><%=locate%>&nbsp;</td>
              </tr>
			  <tr>
				<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
			  </tr>
			  <%
					objrs.movenext
					loop
					objrs.close
					set objrs = nothing
			  %>
            </table>
			</div>
			</td>
          </tr>
      </table>
<!-- -->
	</td>
    <td background="/images/pop_right_middle_bg.gif">&nbsp;</td>
  </tr>
  <tr>
    <td><img src="/images/pop_left_bottom_bg.gif" width="22" height="25"></td>
    <td background="/images/pop_center_bottom_bg.gif">&nbsp;</td>
    <td><img src="/images/pop_right_bottom_bg.gif" width="23" height="25"></td>
  </tr>
</table>
</form>
</body>
</html>
<script language="JavaScript">
<!--
	function put_medium_data(mdidx, sidx, title, categoryidx, categoryname, side, standard, quality, custcode, unit, locate, unitprice, custname, map) {
		var frm = window.opener.document.forms[0];
		frm.mdidx.value = mdidx ;
		frm.sidx.value = sidx ;
		frm.txttitle.value = title.replace("/","'") ;
		frm.txtcategoryidx.value = categoryidx;
		frm.txtcategoryname.value = categoryname;
		frm.selside.value = side;
		frm.txtunit.value = unit;
		frm.txtunitprice.value = unitprice;
		frm.txtstandard.value = standard.replace(/\-/g,'"') ;
		frm.selquality.value = quality;
		frm.selcustcode.value = custcode;
//		frm.txtcustname.value = custname;
//		frm.selcustcode.value = custcode
		frm.txtunitprice.value = unitprice;
		frm.txtlocate.value = locate;
		frm.txtmap.value = map;
		this.close();
	}

	function get_medium_category() {
		var frm = document.forms[0];
		frm.action = "pop_medium_search.asp";
		frm.mehtod = "post";
		frm.submit();
	}

	window.onload = function () {
		self.focus();
	}
//-->
</script>
