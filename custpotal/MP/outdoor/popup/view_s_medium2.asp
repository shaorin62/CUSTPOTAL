<!--#include virtual="/MP/outdoor/inc/Function.asp" -->
<%
	Dim pcontidx : pcontidx = request("contidx")
	Dim pcrud : pcrud = request("crud")
	Dim pcyear : pcyear = request("cyear")
	If pcyear = "" then pcyear =  Year(Now) '추가부분 ( 년도가 null 일시에 오류 ....)
	Dim pcmonth : pcmonth = request("cmonth")
	Dim pmdidx : pmdidx = request("mdidx")
	If pmdidx = "" Then pmdidx = 0

	If pcrud = "d" Then server.transfer "/MP/outdoor/process/db_s_medium.asp"

	Dim sql : sql = "select b.standard, b.quality, a.region, a.locate, a.unit, a.medcode, a.empid , a.categoryidx, isnull(c.qty,1) as qty, isnull(c.monthly,0) as monthly, isnull(c.expense,0) as expense, d.thmno, e.categoryname, f.thmname, g.title, h.highcustcode, d.seq as subseq  from wb_contact_mst g left outer join wb_contact_md a on g.contidx=a.contidx and a.mdidx = ? left outer join  wb_contact_md_dtl b on b.seq = (select max(seq) from wb_contact_md_dtl where mdidx=? and cyear+cmonth <= ?)  left outer join wb_contact_exe c on a.mdidx=c.mdidx and c.cyear=? and c.cmonth=? left outer join wb_subseq_exe d on a.mdidx=d.mdidx and d.seq = (select max(seq) from wb_subseq_exe where cyear+cmonth <= ? and mdidx=?) left outer join wb_category e on a.categoryidx=e.categoryidx left outer join wb_subseq_dtl f on d.thmno=f.thmno left outer join sc_cust_dtl h on g.custcode=h.custcode where g.contidx=?"
'	response.write sql

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandType = adCmdText
	cmd.commandText = sql
	cmd.parameters.append cmd.createparameter("mdidx", adInteger, adParamInput)
	cmd.parameters.append cmd.createparameter("mdidx2", adInteger, adParamInput)
	cmd.parameters.append cmd.createparameter("yearmon", adChar, adParamInput, 6)
	cmd.parameters.append cmd.createparameter("cyear", adChar, adParamInput, 4)
	cmd.parameters.append cmd.createparameter("cmonth", adChar, adParamInput, 2)
	cmd.parameters.append cmd.createparameter("yearmon2", adChar, adParamInput, 6)
	cmd.parameters.append cmd.createparameter("mdidx3", adInteger, adParamInput)
	cmd.parameters.append cmd.createparameter("contidx", adInteger, adParamInput)
	cmd.parameters("mdidx").value = pmdidx
	cmd.parameters("mdidx2").value = pmdidx
	cmd.parameters("yearmon").value = pcyear&pcmonth
	cmd.parameters("cyear").value = pcyear
	cmd.parameters("cmonth").value = pcmonth
	cmd.parameters("yearmon2").value = pcyear&pcmonth
	cmd.parameters("mdidx3").value = pmdidx
	cmd.parameters("contidx").value = pcontidx
	Dim rs : Set rs = cmd.execute
	clearparameter(cmd)

	Dim standard : standard = rs(0)
	Dim quality : quality = rs(1)
	Dim region : region = rs(2)
	Dim locate : locate = rs(3)
	Dim unit : unit =rs(4)
	Dim medcode : medcode = rs(5)
	Dim empid : empid = rs(6)
	Dim categoryidx : categoryidx = rs(7)
	Dim qty : qty = rs(8)
	Dim monthly : monthly = rs(9)
	Dim expense : expense = rs(10)
	Dim thmno : thmno = rs(11)
	Dim categoryname : categoryname = rs(12)
	Dim thmname : thmname = rs(13)
	Dim title : title = rs(14)
	Dim highcustcode : highcustcode = rs(15)
	Dim subseq : subseq = rs(16)

%>
<html>
<head>
	<title>▒▒ SK M&C | Media Management System ▒▒  </title>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<link href="/MP/outdoor/style.css" rel="stylesheet" type="text/css">
	<script type='text/javascript' src='/js/ajax.js'></script>
	<script type='text/javascript' src='/js/script.js'></script>
	<script type="text/javascript">
	<!--
		function submitchange() {
			var frm = document.forms[0];

			if (frm.txtstandard.value.replace(/\s/g, "") == "") {
				alert("매체 규격을 입력하세요");
				frm.txtstandard.focus();
				return false;
			}

			frm.action = "/MP/outdoor/process/db_s_medium2.asp";
			frm.method = "post";
			frm.submit();
		}


		window.onload = function () {
			self.focus();
		}

	//-->
	</script>
</head>
<body>
<form>
<input type="hidden" id="contidx" name="contidx" value="<%=pcontidx%>" />
<input type="hidden" id="cyear" name="cyear" value="<%=pcyear%>" />
<input type="hidden" id="cmonth" name="cmonth" value="<%=pcmonth%>" />
<input type="hidden" id="mdidx" name="mdidx" value="<%=pmdidx%>" />
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="absmiddle"> <%=title%> : 매체 수정  </td>
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
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h" style="width:80px;">규격 </td>
				<td class="sc"><input name="txtstandard" type="text" id="txtstandard"maxlength="100" style="width:370px" value="<%=standard%>"> </td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">재질 </td>
				<td class="sc"><% Call getquality(quality)%></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			</table>
			<p>
			<table border="0" cellpadding="0" cellspacing="0">
			<tr>
                  <td  align="right" valign="bottom" width='495'><a href="#" onclick="submitchange(); return false;"><img src="/images/btn_save.gif" width="59" height="18" border=0 hspace='10'></a><a href="#" onclick="window.close(); return false;"><img src="/images/btn_close.gif" width="57" height="18" border=0 ></a></td>
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