<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '보안관련라이브러리 %>
<!--#include virtual="/inc/head.asp"-->			<% '초기 설정 페이지(에러 메세지 미출력) %>

<%
	dim fso : set fso = server.createobject("scripting.filesystemobject")
	dim uploadform : set uploadform = server.createobject("DEXT.FileUpload")
	uploadform.defaultpath = "C:\map"

	dim idx : idx = uploadform("idx")
	dim cyear : cyear = uploadform("cyear")
	dim cmonth : cmonth = uploadform("cmonth")
	dim contidx : contidx = uploadform("contidx")
	Dim atag

	dim region : region = uploadform("selregion")
	dim locate : locate = clearXSS( uploadform("txtlocate"), atag)
	dim categoryidx : categoryidx = uploadform("txtcategoryidx")
	dim medcode : medcode = uploadform("selcustcode")
	dim map : map = uploadform("txtmap")
	dim trust : trust = uploadform("rdotrust")
	dim side : side = uploadform("selside")
	dim unitprice : unitprice = uploadform("txtunitprice")
	dim qty : qty = uploadform("txtqty")
	dim unit : unit = uploadform("txtunit")
	dim standard : standard = clearXSS( uploadform("txtstandard"), atag)
	dim quality : quality = uploadform("selquality")
	dim monthprice : monthprice = uploadform("txtmonthprice")
	dim expense : expense = uploadform("txtexpense")
	dim thema : thema = uploadform("selsubject")

	dim totalprice, totalqty, attachFile, tmp

	' 첨부파일에 등록가능 여부 판단
	Dim strFileChk
	If map = "" Then
		map = Null
	Else
		strFileChk = Check_Ext(map,"JPG,GIF,PNG")

		If strFileChk  = "error" Then
			Response.write "<script>"
			Response.write "alert('등록할 수 없는 파일입니다.\n\n이미지 파일(JPG,GIF,PNG)만 등록하십시오.');"
			Response.write " this.close();"
			Response.write "</script>"
			Response.End
		End if
	End if

	if region = "" then region = null
	if locate = "" then locate = null
	if side = "" then side = null
	if unitprice = "" then unitprice = 0 else unitprice = replace(unitprice, ",","")
	if quality = "" then quality = null
	if monthprice = "" then monthprice = 0 else monthprice = replace(monthprice, ",","")
	if expense = "" then expense = 0 else expense = replace(expense, ",","")
	if thema = "" then thema = null
	if Len(cmonth) = 1 then cmonth = "0" & cmonth

	dim objrs, sql
	sql = "select locate, categoryidx, medcode, region, unit, trust, map, uuser, udate from dbo.wb_contact_md m inner join dbo.wb_contact_md_dtl d on m.sidx = d.sidx where idx = " & idx
	call set_recordset(objrs, sql)

		objrs("locate") = locate
		objrs("categoryidx") = categoryidx
		objrs("medcode") = medcode
		objrs("region") = region
		objrs("unit") = unit
		objrs("trust") = trim(trust)
		if map <> "" then
			attachFile = uploadform.defaultpath & "\" & objrs("map")
			if fso.fileexists(attachFile) then	fso.deletefile(attachFile)

			tmp = uploadform("txtmap").save(, false)
			map = right(tmp, len(tmp)-InStrRev(tmp, "\"))
			objrs("map") = map
		end if
		objrs("uuser") = request.cookies("userid")
		objrs("udate") = date

		objrs.update

	objrs.close

	sql = "select side, unitprice, standard, quality from dbo.wb_contact_md_dtl where idx = " & idx

	call set_recordset(objrs, sql)

		objrs("side") = side
		objrs("unitprice") = unitprice
		objrs("standard") = standard
		objrs("quality") = quality

		objrs.update

	objrs.close

	sql = "select qty, jobidx, monthprice, expense from dbo.wb_contact_md_dtl_account where idx = " & idx & " and cyear = '" & cyear & "' and cmonth = '" & cmonth & "' "

	call set_recordset(objrs, sql)

		objrs("qty") = qty
		objrs("jobidx") = thema
		objrs("monthprice") = monthprice
		objrs("expense") = expense

		objrs.update

	objrs.close

	sql = "select qty, jobidx from dbo.wb_contact_md_dtl_account where cyear+cmonth >= '"&cyear&cmonth&"' and  idx = " & idx

	call set_recordset(objrs, sql)

	if not objrs.eof then
		do until objrs.eof
			objrs("qty") = qty
			objrs("jobidx") = thema
			objrs.update
		objrs.movenext
		Loop
	end if

	objrs.close

	set objrs = nothing
%>
<script language="JavaScript">
<!--
	window.opener.location.href="pop_contact_view.asp?contidx=<%=contidx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>&idx=<%=idx%>";
	this.close();
//-->
</script>