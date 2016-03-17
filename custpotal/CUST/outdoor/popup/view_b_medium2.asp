<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<%
	Dim mdidx : mdidx = request("mdidx")
	Dim contidx : contidx = request("contidx")
	Dim crud : crud = "U"
	If mdidx = "" Then
		mdidx = 0
		crud = "C"
	End If

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandType = adCmdText

	Dim sql : sql = "select title from wb_contact_mst where contidx = ?"
	cmd.commandText = sql
	cmd.parameters.append cmd.createparameter("contidx", adInteger, adParamInput)
	cmd.parameters("contidx").value = contidx
	Dim rs : Set rs = cmd.execute
	If Not rs.eof Then
		Dim title : title = rs(0)
	End If
	rs.close
	clearparameter(cmd)

	sql="select count(mdidx) from wb_contact_md_dtl where mdidx =?"
	cmd.parameters.append cmd.createparameter("mdidx", adInteger, adParamInput)
	cmd.parameters("mdidx").value = mdidx
	cmd.commandText = sql
	Set rs = cmd.execute
	If Not rs.eof Then
		Dim isCnt : isCnt = rs(0)
	End If
	rs.close
	clearparameter(cmd)

	sql = "select mdidx, categoryidx, region, locate, unit, map, medcode, trust, medclass, validclass, empid  from wb_contact_md where mdidx = ?"
	cmd.commandText = sql
	cmd.parameters.append cmd.createparameter("mdidx", adInteger, adParamInput)
	cmd.parameters("mdidx").value = mdidx
	Set rs = cmd.execute

	If Not rs.eof Then
		Dim categoryidx : categoryidx = rs(1)
		Dim region : region = rs(2)
		Dim locate : locate = rs(3)
		Dim unit : unit = rs(4)
		Dim map : map = rs(5)
		Dim medcode : medcode = rs(6)
		Dim trust : trust = rs(7)
		Dim medclass : medclass = rs(8)
		Dim validclass : validclass = rs(9)
		Dim empid : empid = rs(10)
	End If

	clearParameter(cmd)
%>
<html>
<head>
	<title>▒▒ SK M&C | Media Management System ▒▒  </title>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<link href="/cust/outdoor/style.css" rel="stylesheet" type="text/css">
	<script type='text/javascript' src='/js/ajax.js'></script>
	<script type='text/javascript' src='/js/script.js'></script>
	<script type="text/javascript">
	<!--
		function checkForm() {
			var frm = document.forms[0];
			if (frm.hdncategoryidx.value.replace(/\s/g, "") == "") {
				alert("매체분류를 입력하세요");
				return false;
			}
			if (frm.txtunit.value.replace(/\s/g, "") == "") {
				alert("매체 단위를 입력하세요");
				frm.txtunit.focus();
				return false;
			}
			if (frm.cmbmed.value == "") {
				alert("매체사를 선택하세요");
				frm.cmbmed.focus();
				return false;
			}
//			if (frm.cmbemp.value == "") {
//				alert("담당직원을 선택하세요");
//				frm.cmbemp.focus();
//				return false;
//			}
//			if(frm.rdotrust[0].checked || frm.rdotrust[1].checked) {
//				alert("매체분류를 선택하세요");
//				frm.rdotrust[0].focus();
//				return false;
//			}
			submitchange();
		}

		function checkDelete() {
			if (confirm("매체정보를 삭제하시겠습니까?")) {
				var isCnt = "<%=isCnt%>";
				if (parseInt(isCnt) > 0)  {
					alert ("면 정보가 등록된 매체는 삭제할 수 없습니다.");
					return false;
				} else {
					document.getElementById("crud").value = "D";
					submitchange();
				}
			}
		}

		function submitchange() {
			var frm = document.forms[0];
			frm.action = "/cust/outdoor/process/db_b_medium2.asp";
			frm.method = "post";
			frm.submit();
		}

		function getmedemployee() {
			var medcode = document.getElementById("cmbmed").value;
			var empid = "<%=empid%>";
			var params = "medcode="+medcode+"&empid="+empid ;
			sendRequest("/cust/outdoor/inc/getmedemployee.asp", params, _getmedemployee, "GET");
		}

		function _getmedemployee() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
						var employeeview = document.getElementById("employeeview");
						employeeview.innerHTML = xmlreq.responseText ;
						var cmbemp = document.getElementById("cmbemp");
						cmbemp.style.width = "100px";
				}
			}
		}

		function get_list_medname() {
			if (event.keyCode == "13") {
			var medname = document.getElementById("txtmedname").value;
			var params = "medname="+encodeURIComponent(medname);
			sendRequest("/cust/outdoor/inc/getmedname.asp", params, _get_list_medname, "GET");
			}
		}

		function _get_list_medname() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
						var medview = document.getElementById("medlist");
						medview.innerHTML = xmlreq.responseText ;
						document.getElementById("cmbmed").style.width = '270px';
						document.getElementById("cmbmed").attachEvent("onchange", getmedemployee);
						getmedemployee();
				}
			}
		}

		function getClear(idx) {
			document.forms[0].txtfile.value = "";
			document.forms[0].file.select();
			document.selection.clear();
		}

		function setclose() {
			document.getElementById("classLayer").style.display='none';
		}

		window.onload = function () {
			self.focus();
			var crud = "<%=pcrud%>";
			if (crud == "d") submitchange();
			document.getElementById("txtmedname").attachEvent("onkeyup", get_list_medname);
			document.getElementById("txtcategoryname").attachEvent("onfocus",
			function() {document.getElementById("classLayer").style.display='block';});
//			_sendRequest("/cust/outdoor/inc/getmedname.asp", "medname="+ document.getElementById("txtmedname").value, _get_list_medname, "GET");
			document.getElementById("cmbmed").attachEvent("onchange", getmedemployee);
			document.getElementById("cmbmed").style.width = '270px';
			document.getElementById("cmbregion").style.width='100px';
			_sendRequest("/cust/outdoor/inc/getmedemployee.asp", "medcode="+document.getElementById("cmbmed").value+"&empid=<%=empid%>", _getmedemployee, "GET");

		}
	//-->
	</script>
</head>
<body>
<form enctype='multipart/form-data'>
<input type="hidden" id="mdidx" name="mdidx" value="<%=mdidx%>" />
<input type="hidden" id="contidx" name="contidx" value="<%=contidx%>" />
<input type="hidden" id="crud" name="crud" value="<%=crud%>" />
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="absmiddle"> <%=title%> 매체 관리 </td>
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
				<td class="hdr h" width='95'>설치지역 </td>
				<td class="sc" width='400'><%Call getregion(region)%></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">설치위치 </td>
				<td class="sc"><input name="txtlocate" type="text" id="txtlocate"maxlength="100" style="width:370px" value="<%=locate%>" style="ime-mode:active;"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">매체분류 </td>
				<td class="sc"><input name="txtcategoryname" type="text" id="txtcategoryname"maxlength="100" style="width:370px"  style="ime-mode:active;"  value="<%=getmediumname(categoryidx)%>"><input type="hidden" id="hdncategoryidx" name="hdncategoryidx"  value="<%=categoryidx%>"/></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">매체단위 </td>
				<td class="sc"><input type="text" id="txtunit" name="txtunit" value="<%=unit%>" style="width:100px;" maxlength='10'/></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h"> 매체사</td>
				<td class="sc"><INPUT TYPE="text" NAME="txtmedname" id="txtmedname" style="width:70px;"> <span id="medlist"><%call getmedcombo(medcode)%></span></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h"> 담당직원</td>
				<td class="sc"><span id='employeeview'><select style='width:100px;'><option value=''>매체 담당</option></select></span></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h"> 광고구분 </td>
				<td class="sc"><input type="radio" id="common" name='rdotrust' value="C" checked <%If trust = "C" Then response.write "checked"%>/> 일반  <input type="radio" id="policy" name='rdotrust' value="P" <%If trust = "P" Then response.write "checked"%>/> 정책</td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h" width='95'> 매체약도 </td>
				<td class="sc" width='400'><div id="extra_div" style="position:absolute;z-index:100;">
				<input name="file" type="file" id="file" style="filter:alpha(opacity:0);width:335;height:25;" onChange="this.blur(); document.forms[0].txtfile.value=this.value; " ></div><input name="txtfile" type="text" id="txtfile" style="width:245" readonly value="<%=map%>"><input type='button' value='찾아보기...' style='margin-left:3px;width:86px;''><input type='button' value='삭제' onclick='getClear();'><INPUT TYPE="hidden" NAME="orgfile" value="<%=map%>"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			</table>
			<table border="0" cellpadding="0" cellspacing="0">
			<tr>
                  <td  valign="bottom"  height='50' width='200'><%If mdidx>0 Then %><a href="#" onclick="checkDelete(); return false;"><img src="/images/btn_delete.gif" width="59" height="18"  vspace="5" border=0 hspace='5'></a><%End If %></td>
                  <td  align="right" valign="bottom"  height='50' width='295'><a href="#" onclick="checkForm(); return false;"><img src="/images/btn_save.gif" width="59" height="18"  vspace="5" border=0 hspace='5'></a><a href="#" onclick="window.close(); return false;" ><img src="/images/btn_close.gif" width="57" height="18" vspace="5" border=0 ></a></td>
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
<div id="classLayer"style="LEFT:10px; TOP:92px; width:520px; height:270px;POSITION:absolute; z-index:11; background-color:#CCCCCC;border: 1px solid #333333;color:#000000;padding:10 10 10 10;font-weight:bolder;display:none;"><!--#include virtual="/cust/outdoor/inc/getMediumClass.asp" --></div>