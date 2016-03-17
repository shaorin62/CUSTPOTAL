<!--#include virtual="/inc/getdbcon.asp"-->
<!--#include virtual="/inc/func.asp"-->
<%
	dim uploadform : set uploadform = server.createobject("DEXT.FileUpload")
	uploadform.defaultpath = "c:\pds\monitor"

	dim contidx : contidx = uploadform("contidx")
	dim acceptdate : acceptdate = uploadform("txtacceptdate")
	If acceptdate = "" Then acceptdate = null
	dim acceptweek : acceptweek = uploadform("selweek")
	If acceptweek = "" Then acceptweek = null
	dim status : status = uploadform("rdostatus")
	If status = "" Then status = null
	dim nextacceptdate : nextacceptdate = uploadform("txtnextacceptdate")
	If nextacceptdate = "" Then nextacceptdate = null
	dim file : file = uploadform("txtfile")
	dim comment : comment = uploadform("txtcomment")
	If comment = "" Then comment = null
	dim userid : userid = uploadform("txtuserid")

	dim filecount : filecount = uploadform("txtfile").count

	dim objrs, sql
	sql = "select pidx, contidx,  acceptdate, status, acceptweek, nextacceptdate, comment, cuser, cdate, uuser, udate from dbo.wb_contact_monitor_mst where contidx = " & contidx
	call set_recordset(objrs, sql)

	objrs.addnew
		objrs.fields("contidx").value = contidx
		objrs.fields("acceptdate").value = acceptdate
		objrs.fields("status").value = status
		objrs.fields("acceptweek").value = acceptweek
		objrs.fields("nextacceptdate").value = nextacceptdate
		objrs.fields("comment").value = comment
		objrs.fields("cuser").value = userid
		objrs.fields("cdate").value = date
		objrs.fields("uuser").value = userid
		objrs.fields("udate").value = date
	objrs.update
	dim pidx : pidx = objrs.fields("pidx").value
	objrs.close

	sql = "select didx, pidx, typical, filename from dbo.WB_CONTACT_MONITOR_DTL where pidx = " & pidx
	call set_recordset(objrs, sql)

	dim intLoop , temp, flag
	flag = true
	for intLoop = 1 to filecount
			if uploadform("txtfile")(intLoop) <> "" then 
				objrs.addnew
				objrs.fields("pidx").value = pidx
				objrs.fields("typical").value = flag
				temp = uploadform("txtfile")(intLoop).Save (, false)
				objrs.fields("filename").value = uploadform("txtfile")(intLoop).filename
				objrs.update
				if flag then flag = false
			end if
	next
	objrs.close
%>
<script language="JavaScript">
<!--
	window.opener.location.reload();
	this.close();
//-->
</script>