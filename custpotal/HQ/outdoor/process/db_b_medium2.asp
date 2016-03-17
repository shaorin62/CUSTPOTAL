<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '보안관련라이브러리 %>
<%

	Dim uploadform : Set uploadform = Server.CreateObject ("DEXT.FileUpload")
	uploadform.defaultpath = "C:\pds\media"

	Dim atag : atag = ""
	Dim crud : crud = uploadform("crud")
	Dim contidx : contidx = uploadform("contidx")
	Dim mdidx : mdidx = uploadform("mdidx")
	Dim region : region = uploadform("cmbregion")
	Dim locate : locate = uploadform("txtlocate")
	Dim categoryidx : categoryidx = uploadform("hdncategoryidx")
	Dim unit : unit = uploadform("txtunit")
	Dim medcode : medcode = uploadform("cmbmed")
	Dim empid : empid = uploadform("cmbemp")
	Dim trust : trust = uploadform("rdotrust")
	Dim map : map = uploadform("file")
	Dim orgmap : orgmap = uploadform("orgfile")
	Dim txtfile : txtfile = uploadform("txtfile")
	If empid = "" Then empid = Null

'			response.write "crud : " & crud & "<br>"
'			response.write "contidx : " & contidx & "<br>"
'			response.write "mdidx : " & mdidx & "<br>"
'			response.write "region : " & region & "<br>"
'			response.write "locate : " & locate & "<br>"
'			response.write "categoryidx : " & categoryidx & "<br>"
'			response.write "unit : " & unit & "<br>"
'			response.write "medcode : " & medcode & "<br>"
'			response.write "empid : " & empid & "<br>"
'			response.write "trust : " & trust & "<br>"
'			response.write "map : " & map & "<br>"
'			response.write "orgmap : " & orgmap & "<br>"
'			response.write "txtfile : " & txtfile & "<br>"
'			response.End



	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandType = adCmdText

	Select Case UCase(crud)
	Case "C"
		If uploadform("file") <> "" Then
			map = uploadform("file").saveAs (,false)
			map = Right(map, Len(map)-InstrRev(map,"\"))
			txtmap = "<a href='#' onclick='viewMap(); return false;'><img src='/pds/media/"&map&"' width='230' height='190'></a>"
		Else
			map = Null
			txtmap = "<img src='/images/noimage.gif' width='230' height='190'>"
		End If
		sql = "insert into wb_contact_md(categoryidx, region, locate, unit, map, medcode, trust, contidx, empid, cuser, cdate, uuser, udate) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
		cmd.parameters.append cmd.createparameter("categoryidx", adInteger, adParaminput)
		cmd.parameters.append cmd.createparameter("region", advarchar, adParaminput, 10)
		cmd.parameters.append cmd.createparameter("locate", advarchar, adParaminput, 200)
		cmd.parameters.append cmd.createparameter("unit", advarchar, adParaminput, 10)
		cmd.parameters.append cmd.createparameter("map", advarchar, adParaminput, 100)
		cmd.parameters.append cmd.createparameter("medcode", adChar, adParaminput, 6)
		cmd.parameters.append cmd.createparameter("trust", adVarchar, adParaminput, 10)
		cmd.parameters.append cmd.createparameter("contidx", adInteger, adParaminput)
		cmd.parameters.append cmd.createparameter("empid", adChar, adParaminput, 9)
		cmd.parameters.append cmd.createparameter("cuser", adVarChar, adParaminput, 12)
		cmd.parameters.append cmd.createparameter("cdate", adDBTimeStamp, adParaminput)
		cmd.parameters.append cmd.createparameter("uuser", adVarChar, adParaminput, 12)
		cmd.parameters.append cmd.createparameter("udate", adDBTimeStamp, adParaminput, 12)
		cmd.parameters("categoryidx").value = categoryidx
		cmd.parameters("region").value = region
		cmd.parameters("locate").value = locate
		cmd.parameters("unit").value = unit
		cmd.parameters("map").value = map
		cmd.parameters("medcode").value = medcode
		cmd.parameters("trust").value = trust
		cmd.parameters("contidx").value = contidx
		cmd.parameters("empid").value = empid
		cmd.parameters("cuser").value = session("userid")
		cmd.parameters("cdate").value = date
		cmd.parameters("uuser").value = null
		cmd.parameters("udate").value = Null
		cmd.commandText = sql
		cmd.execute ,, adExecuteNoRecords

		clearparameter(cmd)
		sql = "select @@identity from wb_contact_md "
		cmd.commandText = sql
		Set rs = cmd.execute

		mdidx = rs(0)
		rs.close
		Set rs = Nothing
	Case "U"
		If txtfile = "" Then
			If uploadform.FileExists(uploadform.defaultpath&"\"&orgmap) Then uploadform.deleteFile(uploadform.defaultpath&"\"&orgmap)
			map = Null
			txtmap = "<img src='/images/noimage.gif' width='230' height='190'>"
		ElseIf Trim(txtfile) <> Trim(orgmap) Then
			If uploadform.FileExists(uploadform.defaultpath&"\"&orgmap) Then uploadform.deleteFile(uploadform.defaultpath&"\"&orgmap)
			map = uploadform("file").saveAs (,false)
			map = Right(map, Len(map)-InstrRev(map,"\"))
			txtmap = "<a href='#' onclick='viewMap(); return false;'><img src='/pds/media/"&map&"' width='230' height='190'></a>"
		ElseIf Trim(txtfile) = Trim(orgmap) Then
			map = orgmap
			txtmap = "<a href='#' onclick='viewMap(); return false;'><img src='/pds/media/"&orgmap&"' width='230' height='190'></a>"
		End If

		sql = "update wb_contact_md set categoryidx =?, region =?, locate=?, unit=?, map=?, medcode=?, trust=?, empid=?, uuser=?, udate=?  where mdidx=?"
		cmd.parameters.append cmd.createparameter("categoryidx", adInteger, adParaminput)
		cmd.parameters.append cmd.createparameter("region", advarchar, adParaminput, 10)
		cmd.parameters.append cmd.createparameter("locate", advarchar, adParaminput, 200)
		cmd.parameters.append cmd.createparameter("unit", advarchar, adParaminput, 10)
		cmd.parameters.append cmd.createparameter("map", advarchar, adParaminput, 100)
		cmd.parameters.append cmd.createparameter("medcode", adChar, adParaminput, 6)
		cmd.parameters.append cmd.createparameter("trust", adVarchar, adParaminput, 10)
		cmd.parameters.append cmd.createparameter("empid", adChar, adParaminput, 9)
		cmd.parameters.append cmd.createparameter("uuser", adVarChar, adParaminput, 12)
		cmd.parameters.append cmd.createparameter("udate", adDBTimeStamp, adParaminput, 12)
		cmd.parameters.append cmd.createparameter("mdidx", adInteger, adParamInput)
		cmd.parameters("categoryidx").value = categoryidx
		cmd.parameters("region").value = region
		cmd.parameters("locate").value = locate
		cmd.parameters("unit").value = unit
		cmd.parameters("map").value = map
		cmd.parameters("medcode").value = medcode
		cmd.parameters("trust").value = trust
		cmd.parameters("empid").value = empid
		cmd.parameters("uuser").value = session("userid")
		cmd.parameters("udate").value = date
		cmd.parameters("mdidx").value = mdidx
		cmd.commandText = sql
		cmd.execute ,, adExecuteNoRecords

	Case "D"
		If uploadform.FileExists(uploadform.defaultpath&"\"&map) Then uploadform.deleteFile(uploadform.defaultpath&"\"&map)
		sql = "delete from wb_contact_md where mdidx =?"
		cmd.parameters.append cmd.createparameter("mdidx", adInteger, adParamInput)
		cmd.parameters("mdidx").value = mdidx
		cmd.commandText = sql
		cmd.execute ,, adExecuteNoRecords
		mdidx = ""
		txtmap = "<img src='/images/noimage.gif' width='230' height='190'>"
	End Select


%>
<script type="text/javascript">
<!--
	var crud = "<%=crud%>";
	if (crud == "d") {alert(opener.document.URL);}
	window.opener.document.getElementById("mdidx").value = "<%=mdidx%>";
	window.opener.document.getElementById("map").innerHTML = "<%=txtmap%>";
	window.opener.getcontact();
	window.close();
//-->
</script>