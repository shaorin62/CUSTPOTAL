<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim contidx : contidx = request("contidx")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")

	dim objrs, sql
	sql = "select m.title, m.firstdate, m.startdate, m.enddate, m.comment, m.regionmemo, m.mediummemo, c.custcode, c.custname, c2.custcode as custcode2, c2.custname as custname2 from dbo.wb_contact_mst m  inner join dbo.sc_cust_temp c on m.custcode = c.custcode inner join dbo.sc_cust_temp c2 on c.highcustcode = c2.custcode where contidx = " & contidx
	call get_recordset(objrs, sql)

	dim title, firstdate, startdate, enddate, comment, regionmemo, mediummemo , custcode, custcode2, custname, custname2
	if not objrs.eof then
		title = objrs("title")
		firstdate = objrs("firstdate")
		startdate = objrs("startdate")
		enddate = objrs("enddate")
		comment = objrs("comment")
		regionmemo = objrs("regionmemo")
		mediummemo = objrs("mediummemo")
		custcode = objrs("custcode")
		custcode2 = objrs("custcode2")
		custname = objrs("custname")
		custname2 = objrs("custname2")
	end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<title>�Ƣ� SK M&C | Media Management System �Ƣ�  </title>
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
<input type="hidden" name="contidx" value="<%=contidx%>">
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> <%=title%> </td>
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

			<table border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td class="hw">����</td>
				<td class="bw"><input name="txttitle" type="text" id="txttitle"maxlength="100" style="width:350px" value="<%=title%>"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hw">���ʰ����</td>
				<td class="bw"><input name="txtfirstdate" type="text" id="txtfirstdate" maxlength="10"  value="<%=firstdate%>"> <img src="/images/calendar.gif" width="39" height="20" border="0" align="absmiddle" onclick="Calendar_D(document.all.txtfirstdate)"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hw">������</td>
				<td class="bw"><input name="txtstartdate" type="text" id="txtstartdate" > <img src="/images/calendar.gif" width="39" height="20" border="0" align="absmiddle" onclick="Calendar_D(document.all.txtstartdate)" ></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hw">������</td>
				<td class="bw"><input name="txtenddate" type="text" id="txtenddate"  > <img src="/images/calendar.gif" width="39" height="20" border="0" align="absmiddle" onclick="Calendar_D(document.all.txtenddate)"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hw">�����</td>
				<td class="bw"><input type="text" name="txtdeptname" id="txtdeptname" readonly size="35"  value="<%=custname%>">  <input name="txtdeptcode" type="hidden" id="txtdeptcode" size="10" readonly  value="<%=custcode%>"> <img src="/images/btn_find.gif" width="39" height="20" align="absmiddle" class="stylelink" onClick="get_pop_custcode();"> </td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hw">������</td>
				<td class="bw"><input type="text" name="txtcustname" id="txtcustname" readonly size="35"  value="<%=custname2%>"> <input name="txtcustcode" type="hidden" id="txtcustcode" size="10" readonly  value="<%=custcode2%>">
                  </td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
 			<tr>
				<td class="hw">����Ư��</td>
				<td class="bw" style="padding-top:3px; padding-bottom:3px;"><textarea name="txtregionmemo" rows="5"  style="width:350px" id="txtregionmemo"><%if not isnull(regionmemo) then response.write regionmemo%></textarea></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hw">��üƯ��</td>
				<td class="bw" style="padding-top:3px; padding-bottom:3px;"><textarea name="txtmediummemo" rows="5"  style="width:350px" id="txtmediummemo"><%if not isnull(mediummemo) then response.write mediummemo%></textarea></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hw">Ư�̻���</td>
				<td class="bw" style="padding-top:3px; padding-bottom:3px;"><textarea name="txtcomment" rows="5"  style="width:350px" id="txtcomment"><%if not isnull(comment) then response.write comment%></textarea></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
                  <td width="50%" height="50" align="left" valign="bottom"><img src="/images/space.gif" width="59" height="18" border="0"></td>
                  <td width="50%" align="right" valign="bottom"><img src="/images/btn_save.gif" width="59" height="18" hspace="10" vspace="5" style="cursor:hand" onClick="check_submit();"><img src="/images/btn_init.gif" width="67" height="18" vspace="5" style="cursor:hand" onClick="set_reset();"><img src="/images/btn_close.gif" width="57" height="18" vspace="6" style="cursor:hand" onClick="set_close();" hspace="10" ></td>
                </tr>
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

<script type="text/javascript" src="/js/calendar.js"></script>
<script type="text/javascript" src="/js/script.js"></script>
<script language="javascript">
<!--
	function check_submit() {
		var frm = document.forms[0];
		if (frm.txttitle.value == "") {
			alert("������ �Է��ϼ���.") ;
			frm.txttitle.focus();
			return false;
		}
		if (frm.txtstartdate.value == "") {
			alert("���������� �Է��ϼ���.") ;
			frm.txtstartdate.focus();
			return false;
		}
		if (frm.txtenddate.value == "") {
			alert("����������� �Է��ϼ���.");
			frm.txtenddate.focus();
			return false;
		}
//		if (frm.txttotalprice.value == "") {
//			alert("�� ������ �Է��ϼ���.") ;
//			frm.txttotalprice.focus();
//			return false;
//		}
//		if (frm.txtmonthprice.value == "") {
//			alert("�������� �Է��ϼ���.");
//			frm.txtmonthprice.focus();
//			return false;
//		}
//		if (frm.txtexpense.value == "") {
//			alert("�����ֺ� �Է��ϼ���.");
//			frm.txtexpense.focus();
//			return false;
//		}
		if (frm.txtcustcode.value == "") {
			alert("����θ�  �Է��ϼ���.");
			frm.txtcustcode.focus();
			return false;
		}

		frm.method = "POST";
		frm.action = "contact_extention_proc.asp";
		frm.submit();
	}

	function get_pop_custcode() {
		var url = "pop_custcode.asp";
		var name = "pop_custcode";
		var opt = "width=540, height=500, resziable=no, scrollbars = yes, status=yes, top=100, left=600";
		window.open(url, name, opt);
	}

	function set_reset() {
		document.forms[0].reset();
	}

	function set_trust_text() {
		var frm = document.forms[0];
		var bln = frm.chktrust.checked ;
		var account = document.getElementById("account")
		var ratio = document.getElementById("ratio")
		if (bln) {
			frm.txtmonthpay.disabled = bln ;
			account.innerHTML = "������(��)";
			ratio.innerHTML = "��������(%)" ;

		} else {
			frm.txtmonthpay.disabled = bln ;
			account.innerHTML = "������(��)";
			ratio.innerHTML = "������(%)" ;
		}
		frm.txtmonthpay.value = "0";
		frm.txtincome.value = "0";
		frm.txtincomeratio.value = "0";
	}

	function set_close() {
		this.close();
	}

	window.onload = function () {

	}
//-->
</script>