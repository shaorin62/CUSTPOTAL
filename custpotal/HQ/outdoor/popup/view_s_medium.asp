<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<%
'	Dim item
'	For Each item In request.querystring
'		response.write item &  " : " & request.querystring(item) & "<br>"
'	Next
	Dim pcontidx : pcontidx = request("contidx")
	Dim pcrud : pcrud = request("crud")
	Dim pcyear : pcyear = request("cyear")
	If pcyear = "" then pcyear =  Year(Now) '추가부분 ( 년도가 null 일시에 오류 ....)
	Dim pcmonth : pcmonth = request("cmonth")
	Dim pmdidx : pmdidx = request("mdidx")
	If pmdidx = "" Then pmdidx = 0

	If pcrud = "d" Then server.transfer "/hq/outdoor/process/db_s_medium.asp"

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
	<link href="/hq/outdoor/style.css" rel="stylesheet" type="text/css">
	<script type='text/javascript' src='/js/ajax.js'></script>
	<script type='text/javascript' src='/js/script.js'></script>
	<script type="text/javascript">
	<!--
		function submitchange() {
			var frm = document.forms[0];
			if (frm.txtqty.value == 0) {
				alert("수량은 1개 이상을 입력하세요");
				frm.txtqty.value = "1";
				return false;
			}

//			if (frm.txtstandard.value.replace(/\s/g, "") == "") {
//				alert("매체 규격을 입력하세요");
//				frm.txtstandard.focus();
//				return false;
//			}

			if (frm.txtlocate.value.replace(/\s/g, "") == "") {
				alert("매체 설치위치를 입력하세요");
				frm.txtlocate.focus();
				return false;
			}

			if (!frm.cmbmed.value) {
				alert("매체사를 선택하세요");
				frm.cmbmed.focus();
				return false;
			}

//			if (!frm.cmbemp.value) {
//				alert("매체 담당을 선택하세요");
//				frm.cmbemp.focus();
//				return false;
//			}

			if (frm.hdncategoryidx.value.replace(/\s/g, "") == "") {
				alert("매체 분류를 입력하세요");
				frm.txtcategoryname.focus();
				return false;
			}

			frm.action = "/hq/outdoor/process/db_s_medium.asp";
			frm.method = "post";
			frm.submit();
		}

		function calculation() {
			// 내수액(율) 자동 계산, 월광고료 선입력 필수 체크
			var monthly = parseFloat(document.getElementById("txtmonthly").value.replace(/,/g, ""));
			var expense = parseFloat(document.getElementById("txtexpense").value.replace(/,/g, ""));

			var income = (monthly - expense).toLocaleString().slice(0, -3) ;
			if (monthly > 0 ) var rate =  ((monthly - expense)/monthly*100).toLocaleString();
			else rate = "0.00" ;

			document.getElementById("incomeview").innerHTML = income + " ("+rate+")";
		}

		function get_list_medname() {
			if (event.keyCode == "13") {
			var medname = document.getElementById("txtmedname").value;
			var medcode = "<%=medcode%>";
			var params = "medname="+encodeURIComponent(medname)+"&medcode="+medcode;
			sendRequest("/hq/outdoor/inc/getmedname.asp", params, _get_list_medname, "GET");
			}
		}

		function _get_list_medname() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
						var medview = document.getElementById("medlist");
						medview.innerHTML = xmlreq.responseText ;
						document.getElementById("cmbmed").style.width = '220px';
						document.getElementById("cmbmed").attachEvent("onchange", getmedemployee);
						getmedemployee();
				}
			}
		}

		function getmedemployee() {
			var medcode = document.getElementById("cmbmed").value;
			var empid = "<%=empid%>";
			var params = "medcode="+medcode+"&empid="+empid ;
			sendRequest("/hq/outdoor/inc/getmedemployee.asp", params, _getmedemployee, "GET");
		}

		function _getmedemployee() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
						var employeeview = document.getElementById("employeeview");
						employeeview.innerHTML = xmlreq.responseText ;
						var cmbemp = document.getElementById("cmbemp");
						cmbemp.style.width = "80px";
				}
			}
		}

		function setclose() {
			document.getElementById("themeLayer").style.display='none';
			document.getElementById("categoryLayer").style.display='none';
		}

		window.onload = function () {
			self.focus();
			var crud = "<%=pcrud%>";
			if (crud == "d") {
				submitchange();
			}
			document.getElementById("txtmedname").attachEvent("onkeyup", get_list_medname);

			document.getElementById("txttheme").attachEvent("onfocus", function() {
				document.getElementById("themeLayer").style.display='block';
				document.getElementById("categoryLayer").style.display='none';
			});
			document.getElementById("txtcategoryname").attachEvent("onfocus", function() {
				document.getElementById("categoryLayer").style.display='block';
				document.getElementById("themeLayer").style.display='none';
			});
//			_sendRequest("/hq/outdoor/inc/getmedname.asp", "medname="+ document.getElementById("txtmedname").value, _get_list_medname, "GET");
			calculation();
			_sendRequest("/hq/outdoor/inc/getmedemployee.asp", "medcode="+document.getElementById("cmbmed").value+"&empid=<%=empid%>", _getmedemployee, "GET");
						document.getElementById("cmbmed").attachEvent("onchange", getmedemployee);
			document.getElementById("cmbmed").style.width = '220px';
		}

		function debug() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
					var debugConsole = document.getElementById("debugConsole");
					debugConsole.innerHTML = xmlreq.responseText ;
				}
			}
		}
	//-->
	</script>
</head>
<body>
<form>
<input type="hidden" id="contidx" name="contidx" value="<%=pcontidx%>" />
<input type="hidden" id="crud" name="crud" value="<%=pcrud%>" />
<input type="hidden" id="cyear" name="cyear" value="<%=pcyear%>" />
<input type="hidden" id="cmonth" name="cmonth" value="<%=pcmonth%>" />
<input type="hidden" id="mdidx" name="mdidx" value="<%=pmdidx%>" />
<input type="hidden" id="subseq" name="subseq" value="<%=subseq%>" />
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="absmiddle"> <%=title%> : 매체<%if contidx ="" then %> 등록 <% else %> <%=getside(pside)%>  수정 <% end if %> </td>
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
				<td class="hdr h">수량/단위 </td>
				<td class="sc"><input name="txtqty" type="text" id="txtqty" maxlength="10"  class="number"  value="<%=qty%>">  <input type="text" id="txtunit" name="txtunit" value="<%=unit%>" style="width:41px;" />  <span style="margin-left:175px;"> <input type="radio" id="common" name='rdotrust' value="C" checked/> 일반  <input type="radio" id="policy" name='rdotrust' value="P" /> 정책 </span></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">규격/재질 </td>
				<td class="sc"><input name="txtstandard" type="text" id="txtstandard"maxlength="100" style="width:220px" value="<%=standard%>"> <% Call getquality(quality)%></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">매체위치 </td>
				<td class="sc"> <%Call getregion(region)%> <input name="txtlocate" type="text" id="txtlocate"maxlength="100" style="width:320px" style="ime-mode:active;" value="<%=locate%>"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">매체사/담당</td>
				<td class="sc" > <INPUT TYPE="text" NAME="txtmedname" id="txtmedname" style="width:70px;"> <span id="medlist"><%call getmedcombo(medcode)%></span> <span id='employeeview'><select style='width:100px;'><option value=''>매체 담당</option></select></span></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h"> 월광고료</td>
				<td class="sc"> <input name="txtmonthly" type="text" id="txtmonthly" maxlength="20"  class="currency"  onclick='comma(this);'  onkeyup="comma(this); calculation();" value="<%=FormatNumber(monthly,0)%>" ></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h"> 월지급액</td>
				<td class="sc"> <input name="txtexpense" type="text" id="txtexpense" maxlength="20"  class="currency"  value="<%=FormatNumber(expense,0)%>" onkeyup="comma(this); calculation();" onclick='comma(this);' ></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h"> 내수액(율) </td>
				<td class="sc"> <div id="incomeview">0 (0.00)</div></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h" width='95'> 매체분류 </td>
				<td class="sc" width='400'> <input name="txtcategoryname" type="text" id="txtcategoryname"maxlength="100" style="width:370px" style="ime-mode:active;"  value="<%=categoryname%>">  <input type="hidden" id="hdncategoryidx" name="hdncategoryidx"  value="<%=categoryidx%>"/></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h" width='95'> 집행소재 </td>
				<td class="sc" width='400'><input name="txttheme" type="text" id="txttheme" style="width:370px" style="ime-mode:active;" value="<%=thmname%>"> <input type='hidden' name='hdnthmno' id='hdnthmno' value="<%=thmno%>"> </td>
			</tr>
			</table>
			<table border="0" cellpadding="0" cellspacing="0">
			<tr>
                  <td  align="right" valign="bottom" width='495'><% if pmdidx <> "" then %> * 변경된 내역은 선택된 년월 이후 모든 내역에 반영됩니다. <% end if %><br><a href="#" onclick="submitchange(); return false;"><img src="/images/btn_save.gif" width="59" height="18" border=0 hspace='10'></a><a href="#" onclick="window.close(); return false;"><img src="/images/btn_close.gif" width="57" height="18" border=0 ></a></td>
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
<div id="themeLayer"style="LEFT:10px; TOP:67px; width:520px; height:297px;position:absolute; z-index:10; background-color:#CCCCCC;border: 1px solid #333333;color:#000000;padding:10 10 10 10;font-weight:bolder;display:none;overflow:hidden;"><!--#include virtual="/hq/outdoor/inc/getsubseq.asp" --></div>

<div id="categoryLayer"style="LEFT:10px; TOP:67px; width:520px; height:297px;POSITION:absolute; z-index:11; background-color:#CCCCCC;border: 1px solid #333333;color:#000000;padding:10 10 10 10;font-weight:bolder;display:none;"><!--#include virtual="/hq/outdoor/inc/getMediumClass.asp" --></div>
<div id='debugConsole'></div>
