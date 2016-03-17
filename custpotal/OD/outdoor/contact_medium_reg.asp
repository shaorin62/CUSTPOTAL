<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim contidx : contidx = request("contidx")

	dim objrs, sql
	sql = "select title, highcustcode from dbo.wb_contact_mst m inner join dbo.sc_cust_temp c on m.custcode = c.custcode  where contidx = " & contidx

	call get_recordset(objrs, sql)

	dim title : title = objrs("title").value
	dim custcode  : custcode = objrs("highcustcode").value

	objrs.close
	set objrs = nothing

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
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
</style></head>

<body  oncontextmenu="return false">
<form>
	  <input type="hidden" name="sidx" >
	  <input type="hidden" name="contidx" value="<%=contidx%>">
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> <<%=title%>> 매체 등록 </td>
    <td background="/images/pop_bg.gif" align="right"><img src="/images/pop_logo.gif" width="121" height="51"></td>
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
	  <table border="0" cellpadding="0" cellspacing="0" align="center">
          <tr>
            <td class="tdhd s4">매체명</td>
            <td colspan="3" class="tdbd s7"><input type="text" name="txttitle" readonly style="width:330px;" >  &nbsp;<img src="/images/btn_find.gif" width="39" height="20" align="absmiddle" class="stylelink" onClick="pop_medium_search();"> </td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">분류</td>
            <td  class="thbd s5"><input type="text" name="txtcategoryname" readonly  >&nbsp;<input type="hidden" name="txtcategoryidx" ></td>
            <td  class="tdhd s4">면</td>
            <td  class="tdbd s6"><% call get_side_code(null)%> &nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">단위</td>
            <td class="thbd s6"><input type="text" name="txtunit" readonly style="width:42px;" class="number" >&nbsp;</td>
            <td class="tdhd s4">단가</td>
            <td class="thbd s6"><input type="text" name="txtunitprice" readonly style="width:42px;" class="number" value="0">&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">규격</td>
            <td class="thbd s5"><input type="text" name="txtstandard" readonly >&nbsp;</td>
            <td class="tdhd s4">재질</td>
            <td class="thbd s6"><% call  get_quality_code(null) %>&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">수량*</td>
            <td class="thbd s5"><input name="txtqty" type="text" size="5"  class="number"   value="1"></td>
            <td class="tdhd s4">등급*</td>
            <td  class="thbd s6"><input name="rdotrust" type="radio" value="일반" checked>
              일반
              <input name="rdotrust" type="radio" value="정책" >
              정책</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">매체사</td>
            <td colspan="3"  class="tdbd s7"><% call get_medium_custcode(custcode, null)%><input type="hidden" name="txtcustname"><input type="hidden" name="txtlocate">&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">월광고료*</td>
            <td colspan="3"  class="tdbd s7"><input type="text" name="txtmonthprice" id="txtmonthprice" maxlength="17"  class="number"  onfocus="this.select();return false;"   onkeyup="comma(document.getElementById('txtmonthprice'));" value="0"> &nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">월지급액*</td>
            <td colspan="3"  class="tdbd s7"><input type="text" name="txtexpense" id="txtexpense" maxlength="17" class="number"   onfocus="this.select();return false;"  onblur="calculation_income(this);" onkeyup="comma(document.getElementById('txtexpense'));" value="0">&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">내수액*</td>
            <td colspan="3"  class="tdbd s7"><input type="text" name="txtincome"  class="number" readonly value="0">&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">내수율*</td>
            <td colspan="3"  class="tdbd s7"><input type="text" name="txtincomeratio"  class="number" readonly value="0.00">&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">소재명*</td>
            <td colspan="3"  class="tdbd s7"><%call get_jobcust_subject(custcode, null, null, null) %></td>
          </tr>
		  <tr>
				<td colspan="4"  height="50" valign="bottom" align="right"> <img src="/images/btn_save.gif" width="59" height="20" vspace="5" style="cursor:hand" onClick="check_submit();" hspace="10" ><img src="/images/btn_init.gif" width="67" height="20" vspace="5" style="cursor:hand" onClick="set_reset();"><img src="/images/btn_close.gif" width="57" height="18" vspace="6" style="cursor:hand" onClick="set_close();" hspace="10" >
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
			if (frm.selsubject.value == "") {
				alert("소재를 선택하세요");
				frm.selsubject.focus();
				return false;
			}
			frm.action = "contact_medium_reg_proc.asp";
			frm.method = "post";
			frm.submit();

		}

		function set_reset() {
			document.forms[0].reset();
		}

		function set_close() {
			this.close();
		}

		window.onload=function () {
			self.focus();
		}
	//-->
	</script>