<%@CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%

	Dim pcyear : pcyear = request("cyear")
	Dim pcmonth : pcmonth = request("cmonth")
	Dim pmdidx : pmdidx = request("mdidx")
	Dim pside : pside = request("side")
	If pmdidx = "" Then pmdidx = 0
	If pside = "" Then pside = "F"
	Dim sql : sql = "select desc_01, desc_02, desc_03, desc_04 from wb_contact_photo "
	sql = sql & " where seq = (select max(seq) from wb_contact_photo where cyear+cmonth <= '"&pcyear & pcmonth &"' and mdidx="&pmdidx&" and side='"&pside&"') "


	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adcmdText
	Dim rs : Set rs = cmd.execute
	Set cmd = Nothing
	If  rs.eof Then 
%>
<table width="1024" align="center" style="padding: 10 10 0 10 " cellpadding=0 cellspacing=0 border=0>
	<tr>
		<td><img src='/images/noimage.gif' width='240' height='180' class='noimage'></td>
		<td><img src='/images/noimage.gif' width='240' height='180' class='noimage'></td>
		<td><img src='/images/noimage.gif' width='240' height='180' class='noimage'></td>
		<td><img src='/images/noimage.gif' width='240' height='180' class='noimage'></td>
	</tr>
</table>
<% Else %>
<table width="1024" align="center" style="margin-top:20px;" >
	<tr>
		<td><%=getimage(rs(0))%></td>
		<td><%=getimage(rs(1))%></td>
		<td><%=getimage(rs(2))%></td>
		<td><%=getimage(rs(3))%></td>
	</tr>
</table>
<%
	End if
	Function getimage(photo)
		If IsNull(photo) Then 
			getimage = "<img src='/images/noimage.gif' width='240' height='180'  class='noimage'>"
		Else 
			getimage = "<a href='#'  onclick='preview();'><img src='/pds/media/"&photo&"' width=240' height='180' class='photo' ></a>"
		End If 
	End Function 

%>