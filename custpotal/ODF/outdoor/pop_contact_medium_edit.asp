<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim contidx : contidx = request("contidx")
	dim sidx : sidx = request("sidx")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	if cyear = "" then cyear = year(date)
	if cmonth = "" then cmonth = month(date)

	dim objrs, sql
	sql = "select highcustcode, title from dbo.wb_contact_mst m inner join dbo.sc_cust_temp c on m.custcode = c.custcode  where contidx = " & contidx

	call get_recordset(objrs, sql)

	dim cont_title : cont_title = objrs("title").value
	dim cont_custcode : cont_custcode  = objrs("highcustcode").value

	objrs.close

	sql = "select m.contidx, m.sidx, m.title, m.locate, m.categoryidx, m.side, m.unit, m.unitprice, m.standard, m.quality, m.qty, m.trust, m.map, m.custcode, d.monthprice, d.expense, d.jobidx,  v.mdname , s.custname  from dbo.wb_contact_md m inner join dbo.wb_contact_md_dtl d on m.contidx = d.contidx and m.sidx = d.sidx and d.cyear = '"&cyear&"'  and d.cmonth = '"&cmonth&"' inner join dbo.vw_medium_category v on m.categoryidx = v.mdidx left outer join dbo.wb_jobcust j on d.jobidx = j.jobidx inner join dbo.sc_cust_temp s on m.custcode = s.custcode inner join dbo.wb_contact_mst m2 on m.contidx = m2.contidx where m.contidx = "&contidx&" and m.sidx = "&sidx

	call get_recordset(objrs, sql)

	dim title, locate, categoryidx, side, unit, unitprice, standard, quality, qty, trust, custcode, monthprice, expense, jobidx, comment, categoryname, custname3, map

	if not objrs.eof then
		title = objrs("title").value
		locate = objrs("locate").value
		categoryidx = objrs("categoryidx").value
		side = objrs("side").value
		unit = objrs("unit").value
		unitprice = objrs("unitprice").value
		standard = objrs("standard").value
		quality = objrs("quality").value
		qty = objrs("qty").value
		trust = objrs("trust").value
		custcode = objrs("custcode").value
		monthprice = objrs("monthprice").value
		expense = objrs("expense").value
		jobidx = objrs("jobidx").value
		categoryname = objrs("mdname").value
		custname3 = objrs("custname").value
	end if
	objrs.close
	set objrs = nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></head>

<body  oncontextmenu="return false">
<form>
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> 계약 매체 수정
		</td>
    <td background="/images/pop_logo.gif" align="right" width="121" height="51"><img src="/images/pop_logo.gif" width="121" height="51"></td>
  </tr>
</table>
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="22"><img src="/images/pop_left_body_top.gif" width="22" height="16"></td>
    <td background="/images/pop_center_top_bg.gif">&nbsp;</td>
    <td width="23"><img src="/images/pop_right_body_top.gif" width="23" height="16"></td>
  </tr>
  <tr>
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
    <td>
	<!--  -->

	  <input type="hidden" name="sidx" value="<%=sidx%>">
	  <input type="hidden" name="contidx" value="<%=contidx%>">
	  <input type="hidden" name="txtlocate" value="<%=locate%>">
	  <input type="hidden" name="txtmap" value="<%=map%>">
<input type="hidden" name="cyear" value="<%=cyear%>">
<input type="hidden" name="cmonth" value="<%=cmonth%>">
	  <table border="0" cellpadding="0" cellspacing="0" align="center">
          <tr>
            <td class="hw ph">매체명</td>
            <td colspan="3" class="bw"><input type="text" name="txttitle" readonly style="width:330px;" value="<%=title%>">  &nbsp; </td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw ph">분류</td>
            <td  class="bw pbd"><input type="text" name="txtcategoryname" readonly  value="<%=categoryname%>">&nbsp;<input type="hidden" name="txtcategoryidx"  value="<%=categoryidx%>"></td>
            <td  class="hw ph">면</td>
            <td  class="bw pbd"><% call get_side_code(side)%>&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw ph">단위</td>
            <td class="thbd s6"><input type="text" name="txtunit" readonly style="width:42px;" class="number"  value="<%=unit%>">&nbsp;</td>
            <td class="hw ph">단가</td>
            <td class="thbd s6"><input type="text" name="txtunitprice" readonly style="width:42px;" class="number"  value="<%=unitprice%>">&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw ph">규격</td>
            <td class="bw pbd"><input type="text" name="txtstandard" readonly  value="<%=standard%>">&nbsp;</td>
            <td class="hw ph">재질</td>
            <td class="bw pbd"><% call  get_quality_code(quality) %>&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw ph">수량*</td>
            <td class="bw pbd"><input name="txtqty" type="text" size="5"  class="number"   value="<%=qty%>"></td>
            <td class="hw ph">등급*</td>
            <td  class="bw pbd"><input name="rdotrust" type="radio" value="일반" <%if trust = "일반" then response.write "checked"%>>
              일반
              <input name="rdotrust" type="radio" value="정책" <%if trust = "정책" then response.write "checked"%>>
              정책</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw ph">매체사</td>
            <td colspan="3"  class="bw"><% call get_medium_custcode(custcode, null)%>&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw ph">월광고료*</td>
            <td colspan="3"  class="bw"><input type="text" name="txtmonthprice" id="txtmonthprice" maxlength="17"  class="number"  onfocus="this.select();return false;"   onkeyup="comma(document.getElementById('txtmonthprice'));" value="<%=formatnumber(monthprice, 0)%>"> &nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw ph">월지급액*</td>
            <td colspan="3"  class="bw"><input type="text" name="txtexpense" id="txtexpense" maxlength="17" class="number"   onfocus="this.select();return false;"  onblur="calculation_income(this);" onkeyup="comma(document.getElementById('txtexpense'));" value="<%=formatnumber(expense, 0)%>">&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw ph">내수액*</td>
            <td colspan="3"  class="bw"><input type="text" name="txtincome"  class="number" readonly value="<%=formatnumber(monthprice-expense, 0)%>">&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw ph">내수율*</td>
            <td colspan="3"  class="bw"><input type="text" name="txtincomeratio"  class="number" readonly value="<%if monthprice <> 0 then response.write formatnumber((monthprice-expense)/monthprice*100) else response.write "0.00"%>">&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw ph">소재명*</td>
            <td colspan="3"  class="bw"><%call get_jobcust_subject(cont_custcode, null, null, jobidx) %></td>
          </tr>
		  <tr>
				<td colspan="4"  height="50" valign="bottom" align="right"> <img src="/images/btn_save.gif" width="59" height="18" vspace="5" style="cursor:hand" onClick="check_submit();" hspace="10" ><img src="/images/btn_init.gif" width="67" height="18" vspace="5" style="cursor:hand" onClick="set_reset();"><img src="/images/btn_close.gif" width="57" height="18" vspace="6" style="cursor:hand" onClick="set_close();" hspace="10" >
	</td>
      </table>
	<!--  -->
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

	<script language="JavaScript" src="/js/calendar.js"></script>
	<script language="JavaScript" src="/js/script.js"></script>
	<script language="JavaScript">
	<!--
		function pop_medium_search() {
			var url = "pop_medium_search.asp";
			var name = "pop_medium_search";
			var opt = "width=718, height=680, resizable=no, top=100, left=660;"
			window.open(url, name, opt);
		}
		function check_submit() {
			var frm = document.forms[0];

			if (frm.txtqty.value == "") {
				alert("수량을 입력하셔야 합니다.");
				frm.txtqty.focus();
				return false;
			}
			frm.action = "contact_medium_edit_proc.asp";
			frm.method = "post";
			frm.submit();

		}

		function set_reset() {
			document.forms[0].reset();
		}

		function set_close() {
			this.close();
		}

		function go_search() {
			var frm = document.forms[0];
			frm.action = "pop_contact_medium_edit.asp";
			frm.method = "post";
			frm.submit();
		}

		window.onload=function () {
			self.focus();
		}
	//-->
	</script>