<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim contidx	 											' 계약번호

	dim objrs, sql

	dim title													' 계약 명
	dim clientSubCode									' 계약 체결 부서 코드

	contidx = request("contidx")

	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")

	' 해당 계약의 제목과 담당 사업부서를 가져온다.
	' 사업부서는 관련 브랜드를 가져오기 위해서 필요하다.
	sql = "select title, custcode from dbo.wb_contact_mst where contidx = " & contidx

	call get_recordset(objrs, sql)

	title = objrs("title").value
	clientSubCode = objrs("custcode").value

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
<form enctype="multipart/form-data">
<INPUT TYPE="hidden" NAME="cyear" value="<%=cyear%>">
<INPUT TYPE="hidden" NAME="cmonth" value="<%=cmonth%>">
<input type="hidden" name="contidx" value="<%=contidx%>"><!-- 계약 일련번호 코드 -->
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top">  <%=title%> 매체 등록 </td>
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
            <td class="hw">설치지역</td>
            <td class="bw"><%call get_region_code(null, null) %> </td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">설치위치</td>
            <td class="bw"><input type="text" name="txtlocate"  style="width:370px;" style="ime-mode:active"></td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">매체분류</td>
            <td  class="bw" ><input type="text" name="txtcategoryname" style="width:240px;" readonly>&nbsp;<input type="hidden" name="txtcategoryidx" > <img src="/images/btn_find.gif" width="39" height="20" align="absmiddle" class="stylelink" onClick="pop_medium_category();"></td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">매체사</td>
            <td  class="bw"><% call get_medium_custcode(null, null)%>&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">매체약도</td>
            <td  class="bw" ><input type="file" name="txtmap" style="width:370px;"></td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">등급*</td>
            <td  class="bw"><input name="rdotrust" type="radio" value="일반" checked> 일반 <input name="rdotrust" type="radio" value="정책" > 정책</td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td  class="hw">면</td>
            <td  class="bw"><% call get_side_code(null)%> &nbsp;</td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">단가</td>
            <td class="bw"><input type="text" name="txtunitprice" class="number"  onkeyup="getFormatNumber(document.getElementById('txtunitprice'));"  id="txtunitprice" >&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">수량/단위</td>
            <td class="bw"><input name="txtqty" type="text" size="5"  class="number"   value="1"> / <input type="text" name="txtunit" style="width:42px;ime-mode:active" maxlength="4" >&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">규격</td>
            <td class="bw"><input type="text" name="txtstandard" style="width:370px;" >&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">재질</td>
            <td class="bw"><% call  get_quality_code(null) %>&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">월광고료</td>
            <td  class="bw"><input type="text" name="txtmonthprice" id="txtmonthprice" maxlength="17"  class="number"   onkeyup="getFormatNumber(document.getElementById('txtmonthprice'));" > &nbsp;</td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">월지급액</td>
            <td  class="bw"><input type="text" name="txtexpense" id="txtexpense" maxlength="17" class="number" onkeyup="getFormatNumber(document.getElementById('txtexpense'));" onblur="calculation_income();">&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">내수액</td>
            <td  class="bw"><input type="text" name="txtincome"  class="number" readonly >&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">내수율</td>
            <td  class="bw"><input type="text" name="txtincomeratio"  class="number" readonly>&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">브랜드</td>
            <td  class="bw"><% call get_jobcust(clientsubcode, null, null, "get_jobcust.asp")%></td>
          </tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
          <tr>
            <td class="hw">소재명</td>
            <td  class="bw"><span id="thema"><%call get_jobcust_subject(clientsubcode, null, null, null) %></span></td>
          </tr>
		  <tr>
				<td colspan="2"  height="50" valign="bottom" align="right"> <img src="/images/btn_save.gif" width="59" height="18" vspace="5" style="cursor:hand" onClick="check_submit();" ><img src="/images/btn_init.gif" width="67" height="18" vspace="5" style="cursor:hand" onClick="set_reset();" hspace="10" ><img src="/images/btn_close.gif" width="57" height="18" vspace="6" style="cursor:hand" onClick="set_close();" >
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
<iframe src="get_jobcust.asp" width="500" height="500" frameborder="0" name="scriptFrame" id="scriptFrame"></iframe>
</body>
</html>
	<script language="JavaScript">
	<!--
		function pop_medium_search() {
			var url = "pop_medium_search.asp";
			var name = "pop_medium_search";
			var opt = "width=718, height=680, resizable=no, top=100, left=660;"
			window.open(url, name, opt);
		}

		function pop_medium_category() {
			var url = "pop_medium_category.asp";
			var name = "pop_medium_category";
			var opt = "width=540, height=525, resziable=no, scrollbars = yes, status=yes, top=100, left=660";
			window.open(url, name, opt);
		}

		function check_submit() {
			var frm = document.forms[0];

			if (frm.txtcategoryidx.value == "") {
				alert("매체의 분류를  조회하세요");
				return false;
			}

			if (frm.selcustcode.selectedIndex == 0) {
				alert("매체사를 선택하세요.");
				frm.selcustcode.focus();
				return false;
			}

			if (frm.txtunit.value == "") {
				alert("단위를 입력하세요.");
				frm.txtunit.focus();
				return false;
			}

			if (frm.txtstandard.value == "") {
				alert("규격을 입력하세요.");
				frm.txtstandard.focus();
				return false;
			}

			if (frm.txtqty.value == "") {
				alert("수량을 하세요.");
				frm.txtqty.focus();
				return false;
			}
			frm.action = "contact_medium_reg_proc.asp";
			frm.method = "post";
			frm.submit();

		}

		function go_page(url) {
			var frm = document.forms[0];
			var seqno = frm.selseqno.options[frm.selseqno.selectedIndex].value;

			scriptFrame.location.href = url+"?seqno=" + seqno ;
		}

		function getFormatNumber(element) {
			var val = Number(String(element.value).replace(/[^\d]/g,"")).toLocaleString().toLocaleString().slice(0,-3);
			if (val == 0) element.value = "0";
			else element.value = val ;
		}

		function calculation_income() {
			var frm = document.forms[0];
			var monthprice = frm.txtmonthprice.value.replace(/[^\d]/g, "") ;
			if (monthprice == "") 		monthprice = 0;
			var expense = frm.txtexpense.value.replace(/[^\d]/g, "") ;
			if (expense == "") 		expense = 0 ;
			var income =parseInt(monthprice) - parseInt(expense) ;
			var ratio=0 ;
			if (income <= 0) ratio = "0.00";
			else ratio = income/monthprice*100 ;
frm.txtincome.value = Number(String(income).replace(/[^\d-]/g,"")).toLocaleString().slice(0,-3);
//			if (Income < 0) {			{
//			var nagative = Number(String(income).replace(/[^\d-]/g,"")).toLocaleString().slice(0,-3)
//			frm.txtincome.value ="-"+nagative;
//			} else {
//			frm.txtincome.value = Number(String(income).replace(/[^\d]/g,"")).toLocaleString().slice(0,-3);
//			}
			frm.txtincomeratio.value = Number(String(ratio)).toLocaleString();
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