<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim contidx : contidx = request("contidx")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	Dim custname : custname = request("custname")
	Dim deptname : deptname = request("deptname")

	dim objrs, sql


	sql = "select m.title, d.locate from dbo.wb_contact_mst m inner join dbo.wb_contact_md d on m.contidx = d.contidx where m.contidx = " & contidx
		
	call get_recordset(objrs, sql)

	dim title : title = objrs("title")
	dim locate : locate = objrs("locate")

	objrs.close

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
</style>

	<script language="JavaScript">
	<!--

		function check_submit() {

			var frm = document.forms[0];
			if (frm.file1.value == "") {
				alert("업로드할 보고서양식(ppt, pptx)을 선택하세요");
				return false;
			}

			frm.action = "pop_contact_report_reg_proc.asp";
			frm.method = "post";
			frm.submit();

		}
	//-->
	</script>

</head>

<body  background="/images/pop_bg.gif"  oncontextmenu="return false">
<form enctype="multipart/form-data" onSubmit="return check_submit();">
<table width="686" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top">  <%=cmonth%>/<%=cyear%> &nbsp;&nbsp;  <%=title%> <%if not isnull(locate) then response.write  " : " & locate %></td>
    <td background="/images/pop_bg.gif" align="right"><img src="/images/pop_logo.gif" width="121" height="51"></td>
  </tr>
</table>
<table width="686" border="0" cellspacing="0" cellpadding="0"  bgcolor="#FFFFFF">
  <tr>
    <td width="22" ><img src="/images/pop_left_body_top.gif" width="22" height="16"></td>
    <td background="/images/pop_center_top_bg.gif">&nbsp;</td>
    <td width="23"><img src="/images/pop_right_body_top.gif" width="23" height="16"></td>
  </tr>
  <tr>
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
    <td>
	<!--  -->
	  <table border="0" cellpadding="0" cellspacing="0" align="center">
		<tr>
			<td class="bw"  valign="top" >
				<table border="0" cellpadding="0" cellspacing="0" align="center">
					<tr>
						<td align="center" height="31" ><input type='file' name='file1' id='file1' style='width:500px;'></td>
					</tr>
					<tr>
						<td  height="50" align="right" valign="bottom"><input type="image" src="/images/btn_save.gif" width="59" height="18"   style="cursor:hand"  ><img src="/images/btn_close.gif" width="57" height="18"style="cursor:hand" onClick="window.close();" hspace="5" > </td>
				    </tr>
				</table>
			</td>
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
<input type="hidden" name="contidx" value="<%=contidx%>">
<input type="hidden" name="cyear" value="<%=cyear%>">
<input type="hidden" name="cmonth" value="<%=cmonth%>">
<input type="hidden" name="custname" value="<%=custname%>">
<input type="hidden" name="deptname" value="<%=deptname%>">
<input type="hidden" name="title" value="<%=title%>">
</form>
</body>
</html>