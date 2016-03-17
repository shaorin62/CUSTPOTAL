<!--#include virtual="/mp/outdoor/inc/Function.asp" -->
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
			filename = request("cyear")&request("cmonth")&"_"&getcustname(rs("custcode"))&"_"&getteamname(rs("custcode"))&"_"&getmediumname(rs("categoryidx"))&"_"&rs("title")&"_"&getmedname(rs("medcode"))&".html"
			Call MakeUrlToFile("http://mms.raed.co.kr/mp/outdoor/print/prt_b_contact.asp?contidx="&request("contidx")(intLoop)&"&cyear="&request("cyear")&"&cmonth="&request("cmonth"), filename,request("contidx")(intLoop),request("cyear"),request("cmonth"))
		Else
			filename = request("cyear")&request("cmonth")&"_"&getcustname(rs("custcode"))&"_"&getteamname(rs("custcode"))&"_00_"&rs("title")&"_00.html"
			Call MakeUrlToFile("http://mms.raed.co.kr/mp/outdoor/print/prt_s_contact.asp?contidx="&request("contidx")(intLoop)&"&cyear="&request("cyear")&"&cmonth="&request("cmonth"), filename,request("contidx")(intLoop),request("cyear"),request("cmonth"))

		End If
		rs.close
	Next

'	Dim filename : filename = request("filename")
'	Call MakeUrlToFile("http://10.110.10.86:6666/mp/outdoor/print/prt_s_contact.asp?contidx="&request("contidx")&"&cyear="&request("cyear")&"&cmonth="&request("cmonth"), filename,request("contidx"),request("cyear"),request("cmonth"))

		response.write "<script> alert('현황 파일을 생성하였습니다.'); parent.location.replace('/mp/outdoor/list_report.asp?cmbcustcode="&request("cmbcustcode")&"&cmbteamcode="&request("cmbteamcode")&"&cyear="&request("cyear")&"&cmonth="&request("cmonth")&"&menunum="&request("menunum")&"'); </script>"
%>

