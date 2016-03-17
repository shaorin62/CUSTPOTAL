<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim num : num = request("num")
	dim idx : idx = request("idx")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	dim attachFile
	dim fso : set fso = server.createobject("scripting.filesystemobject")
	dim defaultpath : defaultpath = "C:\pds\media"

	dim objrs, sql
	sql = "select photo_1, photo_2, photo_3, photo_4 from dbo.wb_contact_md_dtl_account where idx ="&idx&" and  cyear="&cyear&" and cmonth="&cmonth
	call set_recordset(objrs, sql)



	select case cstr(num)
	case "1"
		objrs.fields("photo_1").value = null
	attachFile = defaultpath & "\" & objrs.fields("photo_1").value
	case "2"
		objrs.fields("photo_2").value = null
	attachFile = defaultpath & "\" & objrs.fields("photo_2").value
	case "3"
		objrs.fields("photo_3").value = null
	attachFile = defaultpath & "\" & objrs.fields("photo_3").value
	case "4"
		objrs.fields("photo_4").value = null
	attachFile = defaultpath & "\" & objrs.fields("photo_4").value
	end select
	if fso.fileexists(attachFile) then	fso.deletefile(attachFile)
	objrs.update

	objrs.close
	set objrs = nothing
%>
<script language="JavaScript">
<!--
	location.href="pop_photo_mng.asp?idx=<%=idx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
//-->
</script>