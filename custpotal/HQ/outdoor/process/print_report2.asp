<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<%
'    For Each item In request.form
'		response.write item & " : "& request.form(item) & "<br>"
'	Next
'	response.End

	Dim sql, rs , filename
	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	Dim intcontidx
	Dim firstarr , item
	Dim arrreal()
	Dim intRtn , intcount, intCnt
	cmd.activeconnection = application("connectionstring")
	cmd.commandType = adCmdText
	cmd.parameters.append cmd.createparameter("contidx", adInteger, adParamInput)
	
	intcount =  request("contidx").count -1
	ReDim arrreal(intcount)
		
		intRtn = 0
		intcount = request("contidx").count -1

		firstarr = Split(request("contidx"),",")

		For Each item in firstarr
			If  Len(item) <= 5 Then
					arrreal(intRtn) = item
					intRtn  = intRtn + 1

			End If 
		Next
		
	For intLoop = 1 To request("contidx").count
		sql = "select custcode, categoryidx, title, flag, medcode from wb_contact_mst a inner join wb_contact_md b on a.contidx=b.contidx where a.contidx=?"
		cmd.commandText = sql
		
		intCnt = intLoop - 1

		cmd.parameters("contidx").value =arrreal(intCnt)
		
		Set rs = cmd.execute

		If rs("flag") = "B" Then
			MakeUrlToPrint("http://10.110.10.86:6666/hq/outdoor/print/prt_b_contact2.asp?contidx="&arrreal(intCnt)&"&cyear="&request("cyear")&"&cmonth="&request("cmonth"))
		Else
			MakeUrlToPrint("http://10.110.10.86:6666/hq/outdoor/print/prt_s_contact2.asp?contidx="&arrreal(intCnt)&"&cyear="&request("cyear")&"&cmonth="&request("cmonth"))
		End If
		rs.close
		response.write "<br style='page-break-before:always'>"
	Next
	Set rs = Nothing
	Set cmd = Nothing
%>


