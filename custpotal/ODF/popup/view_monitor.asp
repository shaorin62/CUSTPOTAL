<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<%
'	Dim item
'	For Each item In request.querystring
'		response.write item &  " : " & request.querystring(item) & "<br>"
'	Next

	Dim ptitle : ptitle = request("title")
	Dim pcrud : pcrud = request("crud")
	Dim pcyear : pcyear = request("cyear")
	Dim pcmonth : pcmonth = request("cmonth")
	Dim pmdidx : pmdidx = request("mdidx")
	Dim pside : pside = request("side")
	Dim pnum : pnum = request("num")
	Dim pcustcode : pcustcode = request("custcode")
	Dim pteamcode : pteamcode = request("teamcode")
	If UCase(pcrud) = "C" Then 	pnum = pnum + 1

	Dim sql : sql = "select num, cdate, cname , status, comment, img01, img02, img03, img04 from wb_contact_monitor where mdidx=? and side=? and cyear=? and cmonth=? and num=?"
	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	cmd.parameters.append cmd.createparameter("mdidx", adInteger, adparaminput)
	cmd.parameters.append cmd.createparameter("side", adchar, adparaminput, 1)
	cmd.parameters.append cmd.createparameter("cyear", adchar, adparaminput, 4)
	cmd.parameters.append cmd.createparameter("cmonth", adchar, adparaminput, 2)
	cmd.parameters.append cmd.createparameter("num", adUnsignedTinyInt, adparaminput)
	cmd.parameters("mdidx").value = pmdidx
	cmd.parameters("side").value = pside
	cmd.parameters("cyear").value = pcyear
	cmd.parameters("cmonth").value = pcmonth
	cmd.parameters("num").value = pnum
	Dim rs : Set rs = cmd.execute

	If Not rs.eof Then
		Dim num : num = rs("num")
		Dim c_date : c_date = rs("cdate")
		Dim cname : cname = rs("cname")
		Dim status : status = rs("status")
		Dim comment : comment = rs("comment")
		Dim img01 : img01 = rs("img01")
		Dim img02 : img02 = rs("img02")
		Dim img03 : img03 = rs("img03")
		Dim img04 : img04 = rs("img04")
	End If
	If cname = "" Then cname = request.cookies("custname")
%>
<html>
<head>
	<title>▒▒ SK M&C | Media Management System ▒▒  </title>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<link href="/hq/outdoor/style.css" rel="stylesheet" type="text/css">
	<script type='text/javascript' src='/js/ajax.js'></script>
	<script type='text/javascript' src='/js/script.js'></script>
    <script type="text/javascript" src="/js/calendar.js"></script>
<!--
	<SCRIPT LANGUAGE="VBS" for="FileUploadManager" event="OnTransfer_Click()">
		winstyle="height=355,width=445, status=no,toolbar=no,menubar=no,location=no"
		window.open "/odf/process/FileUploadMonitor.htm",null,winstyle
    </SCRIPT>
	<script For="FileUploadManager" Event="OnError(nCode, sMsg, sDetailMsg)" Language="javascript">
		OnFileManagerError(nCode, sMsg, sDetailMsg);
	</script> -->
	<script type="text/javascript">
	<!--
		var bln = true ;
		function validElement() {
			var frm = document.forms[0];
			if (frm.txtcdate.value == "") {
				alert('검수일자는 필수입력입니다.');
				return false;
			}

			if (frm.txtcname.value.replace(/\s/g, "") == "") {
				alert("검수자 이름을 입력하세요");
				frm.txtcname.focus();
				return false;
			}
			submitchange();
		}

		function submitchange() {
			var frm = document.forms[0];
			frm.action = "/odf/process/db_monitor.asp";
			frm.method = "post";
			frm.submit();
		}

		window.onload = function () {
			self.focus();
			var crud = "<%=pcrud%>";
			if (crud == "d") {submitchange();}
			if (status == "") document.getElementById("rdofine").checked = true;
			else {
				if (status == "1") document.getElementById("rdofine").checked = true;
				else document.getElementById("rdobad").checked = true;
			}

		}

		window.onunload = function () {
			window.opener.viewmonitor = null ;
			var crud = "<%=pcrud%>";
			if (crud == "d") submitchange();
		}


		function getClear(idx) {
			document.forms[0].txtfile[idx].value = "";
			document.forms[0].file[idx].select();
			document.selection.clear();
		}

        function OnFileManagerError(nCode, sMsg, sDetailMsg) {
			alert(sMsg);
			return false;
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
<form action='/odf/process/db_monitor.asp' enctype="multipart/form-data">
<input type="hidden" id="crud" name="crud" value="<%=pcrud%>" />
<input type="hidden" id="cyear" name="cyear" value="<%=pcyear%>" />
<input type="hidden" id="cmonth" name="cmonth" value="<%=pcmonth%>" />
<input type="hidden" id="mdidx" name="mdidx" value="<%=pmdidx%>" />
<input type="hidden" id="side" name="side" value="<%=pside%>" />
<input type="hidden" id="custcode" name="custcode" value="<%=pcustcode%>" />
<input type="hidden" id="teamcode" name="teamcode" value="<%=pteamcode%>" />
<input type="hidden" id="orgnum" name='orgnum' value='<%=num%>' />
<table width="550" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:14px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="absmiddle"> <%=ptitle%> (<%=getside(pside)%>) </td>
    <td background="/images/pop_bg.gif" align="right"><img src="/images/pop_logo.gif" width="121" height="51"></td>
  </tr>
</table>
<table width="550" border="0" cellspacing="0" cellpadding="0">
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
				<td class="hdr h">검수일자 </td>
				<td class="sc"><input name="txtcdate" type="text" id="txtcdate" class="dt" maxlength="10" readonly value="<%=c_date%>"> <a href="#" onclick="Calendar_D(document.all.txtcdate); return false;"><img src="/images/calendar.gif" width="39" height="20"  align="absmiddle"></a></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">검수회수 </td>
				<td class="sc">
				<select id='selnum' name='selnum' onchange='validNum();'>
					<option value='1' <%If pnum = 1 Then response.write "selected "%>> 1회차 </option>
					<option value='2' <%If pnum = 2 Then response.write "selected "%>> 2회차 </option>
					<option value='3' <%If pnum = 3 Then response.write "selected "%>> 3회차 </option>
					<option value='4' <%If pnum = 4 Then response.write "selected "%>> 4회차 </option>
					<option value='5' <%If pnum = 5 Then response.write "selected "%>> 5회차 </option>
				</select></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h" width='95'>검수상태 </td>
				<td class="sc" width='400'><input type="radio" id="rdofine" name='rdostatus' value="1" checked/> 양호 <input type="radio" id="rdobad" name='rdostatus' /> 불량 </td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">검수자명 </td>
				<td class="sc"><input type="text" id="txtcname" name="txtcname" style='width:370px' value="<%=cname%>"/></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h" >비&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;고 </td>
				<td class="sc" style='padding-top:5px; padding-bottom:5px;'><textarea id="txtcomment" name='txtcomment' style='width:370px; height:42px;'><%=comment%></textarea></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">모니터링 </td>
				<td class="sc" ><div id="extra_div" style="position:absolute;z-index:100;">
				<input name="file" type="file" id="file" style="filter:alpha(opacity:0);width:335;height:25;" onChange="this.blur(); document.forms[0].txtfile[0].value=this.value; " ></div><input name="txtfile" type="text" id="txtfile" style="width:245" readonly value="<%=img01%>"><input type='button' value='찾아보기...' style='margin-left:3px;width:86px;''><input type='button' value='삭제' onclick='getClear(0);'><INPUT TYPE="hidden" NAME="orgfile" value="<%=img01%>">
				</td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">모니터링 </td>
				<td class="sc" ><div id="extra_div" style="position:absolute;z-index:100;">
				<input name="file" type="file" id="file" style="filter:alpha(opacity:0);width:335;height:25;" onChange="this.blur(); document.forms[0].txtfile[1].value=this.value; " ></div><input name="txtfile" type="text" id="txtfile" style="width:245" readonly value="<%=img02%>"><input type='button' value='찾아보기...' style='margin-left:3px;width:86px;''><input type='button' value='삭제' onclick='getClear(1);'><INPUT TYPE="hidden" NAME="orgfile" value="<%=img02%>">
				</td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">모니터링 </td>
				<td class="sc" ><div id="extra_div" style="position:absolute;z-index:100;">
				<input name="file" type="file" id="file" style="filter:alpha(opacity:0);width:335;height:25;" onChange="this.blur(); document.forms[0].txtfile[2].value=this.value; " ></div><input name="txtfile" type="text" id="txtfile" style="width:245" readonly value="<%=img03%>"><input type='button' value='찾아보기...' style='margin-left:3px;width:86px;''><input type='button' value='삭제' onclick='getClear(2);'><INPUT TYPE="hidden" NAME="orgfile" value="<%=img03%>">
				</td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">모니터링 </td>
				<td class="sc" ><div id="extra_div" style="position:absolute;z-index:100;">
				<input name="file" type="file" id="file" style="filter:alpha(opacity:0);width:335;height:25;" onChange="this.blur(); document.forms[0].txtfile[3].value=this.value; " ></div><input name="txtfile" type="text" id="txtfile" style="width:245" readonly value="<%=img04%>"><input type='button' value='찾아보기...' style='margin-left:3px;width:86px;''><input type='button' value='삭제' onclick='getClear(3);'><INPUT TYPE="hidden" NAME="orgfile" value="<%=img04%>">
				</td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			</table>
			<table border="0" cellpadding="0" cellspacing="0">
			<tr>
                  <td  align="right" valign="bottom" width='550' height='50' style='padding-right:38px;'><a href="#" onclick="validElement(); return false;"><img src="/images/btn_save.gif" width="59" height="18"  vspace="5" border=0></a> <a href="#" onclick="window.close(); return false;"><img src="/images/btn_close.gif" width="57" height="18" vspace="6" border=0 ></a></td>
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
<%
	rs.close
	Set rs = Nothing
	Set cmd = Nothing
%>