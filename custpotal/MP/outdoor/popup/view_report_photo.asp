<!--#include virtual="/MP/outdoor/inc/Function.asp" -->
<%
	'Call getquerystringparameter
	Dim contidx : contidx = request("contidx")
	Dim cyear : cyear = request("cyear")
	Dim cmonth : cmonth = request("cmonth")
	Dim noimages : noimages = "/images/noimage.gif"

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandType = adCmdText
	Dim sql :  sql = "select photo1, photo2, photo3, photo4 from wb_report_photo where seq = (select max(seq) from wb_report_photo where contidx=? and cyear+cmonth<='"&cyear&cmonth&"') "
'	response.write sql
'	response.write contidx
	cmd.parameters.append cmd.createparameter("contidx", adInteger, adParaminput)
	cmd.parameters("contidx").value = contidx
	cmd.commandText = sql
	Dim rs : Set rs = cmd.Execute
	If Not rs.eof Then
		If rs(0) = "" Or IsNull(rs(0)) Then photo1 = noimages Else photo1 ="/pds/media/"&rs(0) End If
		If rs(1) = "" Or IsNull(rs(1)) Then photo2 = noimages Else photo2 = "/pds/media/"&rs(1) End If
		If rs(2) = "" Or IsNull(rs(2)) Then photo3 = noimages Else photo3 ="/pds/media/"&rs(2) End If
		If rs(3) = "" Or IsNull(rs(3)) Then photo4 = noimages Else photo4 ="/pds/media/"&rs(3) End If
	Else
		photo1 = noimages
		photo2 = noimages
		photo3 = noimages
		photo4 = noimages
	End If
	rs.close
%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
			<link href="/MP/outdoor/style.css" rel="stylesheet" type="text/css">
			<title>▒▒ SK M&C | Media Management System ▒▒  </title>
			<script type='text/javascript' src='/js/script.js'></script>
			<script type="text/javascript">
			<!--
				function getpreview() {
					var element = event.srcElement;
					document.getElementById(element.className).src = element.value;
				}

				function resetphoto() {
					if (confirm("선택한 관리사진을 삭제하시겠습니까?")) {
						var noimages = "/images/noimage.gif";
						var element = event.srcElement;
						var img = document.getElementById(element.className);
						img.select();
						document.selection.clear();
						document.getElementById(document.getElementById(element.className).className).src=noimages;
						document.getElementById("crud").value = 'D';
						document.getElementById("no").value = img.className ;
						submitchange();
					}
				}

				function submitchange() {
					var frm = document.forms[0];
					frm.target="processFrame";
					frm.action = "/MP/outdoor/process/db_report_photo.asp";
					frm.method = "post";
					frm.submit();

				}
			//-->
		</script>
	</head>

<body>
<form enctype="multipart/form-data" >
<input type='hidden' name="contidx" value="<%=contidx%>">
<input type='hidden' name="crud">
<input type="hidden" id="no" name="no"/>
<table width="720" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="absmiddle"> 보고서 사진 관리 </td>
    <td background="/images/pop_bg.gif" align="right"><img src="/images/pop_logo.gif" width="121" height="51"></td>
  </tr>
</table>
<table width="720" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="22"><img src="/images/pop_left_body_top.gif" width="22" height="16"></td>
    <td background="/images/pop_center_top_bg.gif">&nbsp;</td>
    <td width="23"><img src="/images/pop_right_body_top.gif" width="23" height="16"></td>
  </tr>
  <tr>
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
    <td height="458" valign='top'>
<!--  -->
		<table border="0" cellpadding="0" cellspacing="0" align='center'>
			<tr>
				<td style='padding-left:30px;' colspan='2'><%Call getyear(cyear)%> <%Call getmonth(cmonth)%></td>
			</tr>
			<tr>
				<td align='center'  style='height:200px;width:300px;'><img src="<%=photo1%>" width="235" height="170" alt="&nbsp;" class='photo' id='photo1' /><br /><input type='file' id='file1' name='file1' class='photo1' onchange='getpreview();'> <a href="#" onclick='resetphoto();'><img src='/images/m_delete.gif' width='16' height='15' alt="관리 사진 삭제" align='absmiddle'  class='file1'></a></td>
				<td style='width:300px;' align='center' ><img src="<%=photo2%>" width="235" height="170" alt="&nbsp;" class='photo' id='photo2' /><br /><input type='file'  id='file2' class='photo2' onchange='getpreview();' name='file2'> <a href="#" onclick='resetphoto();'><img src='/images/m_delete.gif' width='16' height='15' alt="관리 사진 삭제" align='absmiddle' class='file2' ></a></td>
			</tr>
			<tr>
				<td style='height:200px;' align='center' ><img src="<%=photo3%>" width="235" height="170" alt="&nbsp;" class='photo' id='photo3' /><br /><input type='file'  id='file3' class='photo3' onchange='getpreview();' name='file3'> <a href="#" onclick='resetphoto();'><img src='/images/m_delete.gif' width='16' height='15' alt="관리 사진 삭제" align='absmiddle'  class='file3' ></a></td>
				<td align='center' ><img src="<%=photo4%>" width="235" height="170" alt="&nbsp;" class='photo' id='photo4' /><br /><input type='file'  id='file4' class='photo4' onchange='getpreview();' name='file4'> <a href="#" onclick='resetphoto();'><img src='/images/m_delete.gif' width='16' height='15' alt="관리 사진 삭제" align='absmiddle'  class='file4' ></a></td>
			</tr>
		</table>
		<table border="0" cellpadding="0" cellspacing="0">
			<tr height='30'>
				<td width='300' valign='top'>&nbsp;</td>
				<td width='375' align='right' valign='bottom'>  <strong> <a href="#" onclick="submitchange(); return false;">저장</a>  </strong> |  <strong> <a href="#" onclick="window.close(); return false;">닫기</a>  </strong> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
			</tr>
		</table>
<!--  -->
	</td>
    <td background="/images/pop_right_middle_bg.gif"></td>
  </tr>
  <tr>
    <td><img src="/images/pop_left_bottom_bg.gif" width="22" height="25"></td>
    <td background="/images/pop_center_bottom_bg.gif">&nbsp;</td>
    <td><img src="/images/pop_right_bottom_bg.gif" width="23" height="25"></td>
  </tr>
</table>
</form>
<iframe src='about:blank' width='600' height='500' name="processFrame" frameborder=0></iframe>
</body>
</html>