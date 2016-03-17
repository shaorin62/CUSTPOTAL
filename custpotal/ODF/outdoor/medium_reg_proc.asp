<!--#include virtual="/inc/getdbcon.asp"-->
<!--#include virtual="/inc/func.asp"-->
<%
	dim uploadform : set uploadform = server.createobject("DEXT.FileUpload")
	uploadform.defaultpath = server.mappath("../../") & "\pds\media"

	dim title : title =  uploadform("txttitle")
	dim custcode : custcode = uploadform("selcustcode")
	dim categoryidx : categoryidx = uploadform("txtcategoryidx")
	dim region : region = uploadform("selregion")
	dim locate : locate =  uploadform("txtlocate")
	dim unit : unit = uploadform("rdounit")
	dim map : map = uploadform("txtmap")
	dim filename : filename = uploadform("txtmap").filename
	dim regionmemo : regionmemo = uploadform("txtregionmemo")
	dim mediummemo : mediummemo = uploadform("txtmediummemo")

	if region = "" then region = null
	if locate = "" then locate = null
	if map = "" then
		map = null
		filename = null
	else
		dim tmp : tmp = uploadform("txtmap").save(, false)
	end if
	if unit = "±âÅ¸" then unit = uploadform("txtunit")
	if regionmemo = "" then regionmemo = null
	if mediummemo = "" then mediummemo = null


	dim objrs, mdidx
	dim sql : sql = "select top 1 mdidx, title, custcode, categoryidx, unit, region, locate, map, regionmemo, mediummemo, cuser, cdate, uuser, udate from dbo.WB_MEDIUM_MST"
	call set_recordset(objrs, sql)

	objrs.addnew
	objrs.fields("title").value = title
	objrs.fields("custcode").value = custcode
	objrs.fields("categoryidx").value = categoryidx
	objrs.fields("unit").value = unit
	objrs.fields("region").value = region
	objrs.fields("locate").value = locate
	objrs.fields("map").value = filename
	objrs.fields("regionmemo").value = regionmemo
	objrs.fields("mediummemo").value = mediummemo
	objrs.fields("cuser").value = request.cookies("userid")
	objrs.fields("cdate").value = date
	objrs.fields("uuser").value = request.cookies("userid")
	objrs.fields("udate").value = date

	objrs.update

	mdidx = objrs.fields("mdidx").value
	objrs.close
	set objrs = nothing

%>
<script language="JavaScript">
<!--
	window.opener.location.replace("/hq/outdoor/medium_view.asp?mdidx=<%=mdidx%>");
	this.close();
//-->
</script>