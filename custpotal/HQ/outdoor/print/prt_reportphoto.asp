<%

	Dim pcyear : pcyear = request("cyear")
	Dim pcmonth : pcmonth = request("cmonth")
	Dim pcontidx : pcontidx = request("contidx")
	If pcontidx = "" Then pcontidx = 0
	Dim sql : sql = "select photo1, photo2, photo3, photo4 from wb_report_photo "
	sql = sql & " where seq = (select max(seq) from wb_report_photo where cyear+cmonth <= '"&pcyear & pcmonth &"' and contidx="&pcontidx&") "


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
		<td><img src='http://10.110.10.86:6666/images/noimage.gif' width='240' height='180' class='noimage'></td>
		<td><img src='http://10.110.10.86:6666/images/noimage.gif' width='240' height='180' class='noimage'></td>
		<td><img src='http://10.110.10.86:6666/images/noimage.gif' width='240' height='180' class='noimage'></td>
		<td><img src='http://10.110.10.86:6666/images/noimage.gif' width='240' height='180' class='noimage'></td>
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
			getimage = "<img src='http://10.110.10.86:6666/images/noimage.gif' width='240' height='180'  class='noimage'>"
		Else
			getimage = "<img src='http://10.110.10.86:6666/pds/media/"&photo&"' width=240' height='180' class='photo' >"
		End If
	End Function

%>