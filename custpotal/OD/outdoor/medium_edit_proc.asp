<!--#include virtual="/inc/getdbcon.asp"-->
<!--#include virtual="/inc/func.asp"-->
<%
	dim uploadform : set uploadform = server.createobject("DEXT.FileUpload")
	uploadform.defaultpath = "C:\pds\media"

	dim mdidx : mdidx = uploadform("mdidx")
	dim title : title =  uploadform("txttitle")
	dim custcode : custcode = uploadform("selcustcode")
	dim categoryidx : categoryidx = uploadform("txtcategoryidx")
	dim region : region = uploadform("selregion")
	dim locate : locate =  uploadform("txtlocate")
	dim unit : unit = uploadform("rdounit")
	dim map : map = uploadform("txtmap")
	dim filename : filename = uploadform("txtmap").filename

	if region = "" then region = null
	if locate = "" then locate = null
	if unit = "±âÅ¸" then unit = uploadform("txtunit")
	if filename = "" then filename = null
	if map = "" then map = null

	dim objrs
	dim sql : sql = "select mdidx, title, custcode, categoryidx, unit, region, locate, map, uuser, udate from dbo.WB_MEDIUM_MST where mdidx = " & mdidx
	call set_recordset(objrs, sql)

	if  not isnull(map)  then
		dim tmp : tmp = uploadform("txtmap").save(, false)
		objrs.fields("map").value = filename
		dim attachFile : attachFile = server.mappath("..")&"\pds\media" & "\"& objrs("map")
		dim fso : set fso = server.createobject("scripting.filesystemobject")
		if fso.fileexists(attachFile) then fso.deletefile(attachFile)
		set fso = nothing
	end if


	objrs.fields("title").value = title
	objrs.fields("custcode").value = custcode
	objrs.fields("categoryidx").value = categoryidx
	objrs.fields("unit").value = unit
	objrs.fields("region").value = region
	objrs.fields("locate").value = locate
	objrs.fields("uuser").value = request.cookies("userid")
	objrs.fields("udate").value = date
	objrs.update
	objrs.close
	set objrs = nothing

%>
<script language="JavaScript">
<!--
	window.opener.location.href = "medium_view.asp?mdidx=<%=mdidx%>";
	this.close();
//-->
</script>