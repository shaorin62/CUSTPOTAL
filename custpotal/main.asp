<html>
<head>
<title>▒SK MARKETING EXCELLENT▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table id="Table_01" width="1240" height="643" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td rowspan="3" valign="top">
			<img src="/images/main_01.gif" width="210" height="240" alt=""></td>
		<td background="/images/top_02.gif">		<table width="600" height="39" border="0" align="right" cellpadding="0" cellspacing="0">
      <tr>
        <td>&nbsp;</td>
        <td width="244" align="right"><span class="log">&nbsp;<%=request.cookies("custname2")%></span> &nbsp;</td>
        <td width="1" valign="top"><img src="/images/top_vline_bg.gif" width="1" height="32"></td>
        <td width="104" align="right"><span class="log">&nbsp;<%=request.cookies("userid")%></span> &nbsp;</td>
        <td width="1" valign="top"><img src="/images/top_vline_bg.gif" width="1" height="32"></td>
        <td width="164" align="right"><span class="log">&nbsp;<%=formatdatetime(request.cookies("logtime"))%></span></td>
        <td width="1" valign="top"><img src="/images/top_vline_bg.gif" width="1" height="32"></td>
        <td width="85" align="center"><A HREF="/Log_out.asp"><img src="/images/btn_logout.gif" width="64" height="19" border="0"></A></td>
      </tr>
    </table></td>
	</tr>
	<tr>
		<td>
		<table id="Table_01" width="1030" height="48" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td>
			<img src="/images/menu_01.gif" width="300" height="48" alt="" border="0"></td>
		<td>
			<A HREF="/hq/trans/public_01_list.asp?menuNum=1"><img src="/images/menu_02.gif" width="100" height="48" alt="" border="0"></A></td>
		<td>
			<a href="/hq/board/list.asp"><img src="/images/menu_03.gif" width="140" height="48" alt="" border="0"></a></td>
		<td>
			<a href="/hq/outdoor/contact_list.asp?menuNum=1"><img src="/images/menu_04.gif" width="100" height="48" alt="" border="0"></a></td>
		<td><% If request.cookies("class") = "A" Or request.cookies("class") = "N" Then %><a href="/hq/admin/acc_list.asp?menuNum=1" ><img src="/images/menu_05.gif" width="110" height="48" alt="" border="0"></a><% End if%></td>
		<td>
			<img src="/images/menu_06.gif" width="280" height="48" alt="" border="0"></td>
	</tr>
</table>
		</td>
	</tr>
	<tr>
		<td>
			<img src="/images/main_04.gif" width="1030" height="152" alt=""></td>
	</tr>
	<tr>
		<td colspan="2">
			<img src="/images/main_05.gif" width="1240" height="310" alt=""></td>
	</tr>
	<tr>
		<td colspan="2">
			<img src="/images/bottom_bg.gif" width="1240" height="88" alt=""></td>
	</tr>
</table>
</body>
</html>

