<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<%
'	For Each item In request.form
'		response.write item & " : "& request.form(item) & "<br>"
'	Next
'	response.End
	Dim sql, rs , filename
	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandType = adCmdText
	cmd.parameters.append cmd.createparameter("contidx", adInteger, adParamInput)
	For intLoop = 1 To request("contidx").count
		sql = "select custcode, categoryidx, title, flag, medcode from wb_contact_mst a inner join wb_contact_md b on a.contidx=b.contidx where a.contidx=?"
		cmd.commandText = sql
		cmd.parameters("contidx").value = request("contidx")(intLoop)
		Set rs = cmd.execute


		If rs("flag") = "B" Then
			MakeUrlToPrint("http://mms.raed.co.kr/cust/outdoor/print/prt_b_contact.asp?contidx="&request("contidx")(intLoop)&"&cyear="&request("cyear")&"&cmonth="&request("cmonth"))
		Else
			MakeUrlToPrint("http://mms.raed.co.kr/cust/outdoor/print/prt_s_contact.asp?contidx="&request("contidx")(intLoop)&"&cyear="&request("cyear")&"&cmonth="&request("cmonth"))
		End If

'
'
'		If rs("flag") = "B" Then
'			MakeUrlToPrint("http://10.110.10.86:6666/cust/outdoor/print/prt_b_contact.asp?contidx="&request("contidx")(intLoop)&"&cyear="&request("cyear")&"&cmonth="&request("cmonth"))
'		Else
'			MakeUrlToPrint("http://10.110.10.86:6666/cust/outdoor/print/prt_s_contact.asp?contidx="&request("contidx")(intLoop)&"&cyear="&request("cyear")&"&cmonth="&request("cmonth"))
'		End If


		rs.close
		response.write "<br style='page-break-before:always'>"
	Next
	Set rs = Nothing
	Set cmd = Nothing
%>
<script type="text/javascript">
<!--
	self.print();
//-->
</script>

