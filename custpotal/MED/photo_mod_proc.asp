<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim uploadform : set uploadform = server.createobject("DEXT.FileUpload")
	uploadform.defaultpath = "C:\pds\media"

	dim file : file = uploadform("txtfile")
	dim idx : idx = uploadform("idx")
	dim cyear : cyear = uploadform("cyear")
	dim cmonth : cmonth = uploadform("cmonth")
	dim num : num = uploadform("num")


	dim objrs, sql
	sql = "select photo_1, photo_2, photo_3, photo_4 from dbo.wb_contact_md_dtl_account where idx="&idx&" and cyear="&cyear&" and cmonth="&cmonth
	call set_recordset(objrs, sql)

	dim filename : filename = uploadform("txtfile").Save (, false)
		filename = right(filename, len(filename)-InStrRev(filename, "\"))


	response.write filename & "<BR>"
	response.write idx & "<BR>"
	response.write cyear & "<BR>"
	response.write cmonth & "<BR>"

	select case cstr(num)
	case "1"
		objrs.fields("photo_1").value = filename
	case "2"
		objrs.fields("photo_2").value = filename
	case "3"
		objrs.fields("photo_3").value = filename
	case "4"
		objrs.fields("photo_4").value = filename
	end select
	objrs.update

	objrs.close
	set objrs = nothing
%>
<script language="JavaScript">
<!--
	window.opener.location.reload();
	this.close();
//-->
</script>