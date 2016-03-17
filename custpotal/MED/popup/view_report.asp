<!--#include virtual="/hq/outdoor/inc/function.asp" -->
<%

	If request.cookies("userid") = "" Then
		response.write "<script> try {this.close();} catch(e) {window.close();} </script>"
		response.end
	end if
	Dim pmdidx : pmdidx = request("mdidx")
	Dim ptitle : ptitle = request("title")
	Dim plocate : plocate = request("locate")
	Dim pcyear : pcyear = request("cyear")
	Dim pcmonth : pcmonth = request("cmonth")
	Dim pcrud : pcrud = request("crud")


	Dim sql : sql = "select filename from wb_report_dtl where mdidx=? and cyear=? and cmonth=? and empid=?"
	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeConnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput)
	cmd.parameters.append cmd.createparameter("cyear", adchar, adparaminput, 4)
	cmd.parameters.append cmd.createparameter("cmonth", adchar, adparaminput, 2)
	cmd.parameters.append cmd.createparameter("empid", adchar, adparaminput, 9)
	cmd.parameters("mdidx").value = pmdidx
	cmd.parameters("cyear").value = pcyear
	cmd.parameters("cmonth").value = pcmonth
	cmd.parameters("empid").value = request.cookies("userid")
	Dim rs : Set rs = cmd.Execute
	clearparameter(cmd)
	If Not rs.eof Then
		Dim filename : filename = rs(0)
		txtname = "<font color='#990000'><b>이전에 등록된 파일이 있습니다.</b></font>"
	Else
		filename = Null
		txtname = "등록된 파일이 없습니다."
	End If
	Set cmd = Nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/hq/outdoor/style.css" rel="stylesheet" type="text/css">
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
</head>
<script type="text/javascript">
<!--

	var crud = "<%=pcrud%>";

	window.onload = function () {
		if (crud == 'd') submitchange();
		self.focus();
	}

	function submitchange() {
		var frm = document.forms[0];
		if (crud != 'd') {
			if (frm.file.value.replace(/\s/g, "") == "") {
				alert("등록할 보고서 파일을 선택하세요");
				frm.file.focus();
				return false;
			}
		}
		frm.action = "/med/process/db_report.asp";
		frm.method = "post";
		frm.submit();
	}

	function checkfile(p) {
		var file = p.value ;
		var ext = file.substring(file.lastIndexOf(".")+1, file.length);
		if ((ext !="ppt") && (ext != "pptx")) {
			if (p.value != "")	alert('보고서는 파워포인트 형식으로만 등록됩니다');
			document.getElementById("file").select();
			document.selection.clear();
			document.getElementById("file").blur();

			return false;
		}
	}

//-->
</script>
<body  background="/images/pop_bg.gif" >
<form onsubmit="return submitchange();"  enctype="multipart/form-data" >
<input type="hidden" id="crud" name='crud' value='<%=pcrud%>' />
<input type="hidden" id="filename" name='filename' value='<%=filename%>' />
<table width="600" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> <%=ptitle%> <%if not isnull(plocate) then response.write  "(" &plocate&")" %> </td>
    <td background="/images/pop_bg.gif" align="right"><img src="/images/pop_logo.gif" width="121" height="51"></td>
  </tr>
</table>
<table width="600" border="0" cellspacing="0" cellpadding="0"  bgcolor="#FFFFFF">
  <tr>
    <td width="22" ><img src="/images/pop_left_body_top.gif" width="22" height="16"></td>
    <td background="/images/pop_center_top_bg.gif">&nbsp;</td>
    <td width="23"><img src="/images/pop_right_body_top.gif" width="23" height="16"></td>
  </tr>
  <tr>
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
    <td>
	<!--  -->
	  <input type="hidden" name="mdidx" value="<%=pmdidx%>">
	  <input type="hidden" name="cyear" value="<%=pcyear%>">
	  <input type="hidden" name="cmonth" value="<%=pcmonth%>">
	  <table border="0" cellpadding="0" cellspacing="0" align="center" >
		  <tr>
			<td  valign='top' height='25'> <img src='/images/m_ppt.gif' width='16' height='16' align='absbottom' > <%=txtname%> </td>
		  </tr>
		  <tr>
			<td  valign='top'><input type='file' name='file' id='file' style='width:540px;' onchange='checkfile(this);'><input type='hidden' name='orgfile' value="<%=filename%>"></td>
		  </tr>
		  <tr>
			<td height='30' align='right' valign='bottom'> <input type='image'  src='/images/btn_save.gif' width='57' height='18'> <a href="#" onclick='window.close(); return false;'><img src='/images/btn_close.gif' width='57' height='18'> </td>
		  </tr>
		  <tr>
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