<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
'	Dim item
'	For Each item In request.Form
'		response.write item & " : " & request.Form(item) & "<br>"
'	Next
'	response.end
	
	dim contidx : contidx = request("contidx")
	dim org_contidx : org_contidx = contidx 
	dim title : title = request("txttitle")
	dim startdate : startdate = request("txtstartdate")
	dim mediummemo : mediummemo = request("txtmediummemo")
	dim firstdate : firstdate = request("txtfirstdate")
	dim enddate : enddate = request("txtenddate")
	dim custcode : custcode = request("txtdeptcode")
	dim comment : comment = request("txtcomment")
	dim regionmemo : regionmemo = request("txtregionmemo")

	if mediummemo = "" then mediummemo = null
	if regionmemo = "" then regionmemo = null
	if comment = "" then comment = null


	dim objrs, sql

	sql = "select top 1 contidx, custcode, title, firstdate, startdate, enddate,  regionmemo, mediummemo, comment, cuser, cdate, uuser, udate from dbo.wb_contact_mst "

	call set_recordset(objrs, sql)

	if comment =  "" then comment = null

	objrs.addnew
	objrs.fields("custcode").value = custcode
	objrs.fields("title").value = title
	objrs.fields("firstdate").value = firstdate
	objrs.fields("startdate").value = startdate
	objrs.fields("enddate").value = enddate
	objrs.fields("comment").value = comment
	objrs.fields("mediummemo").value = mediummemo
	objrs.fields("regionmemo").value = regionmemo
	objrs.fields("cuser").value = request.cookies("userid")
	objrs.fields("cdate").value = date
	objrs.fields("uuser").value = request.cookies("userid")
	objrs.fields("udate").value = date
	objrs.update

	contidx = objrs.fields("contidx").value

	objrs.close


	sql = "select sidx, monthprice, expense, jobidx from dbo.wb_contact_tmp where contidx = " & org_contidx
	call get_recordset(objrs, sql)

	dim sidx, monthprice, expense, jobidx
	if not objrs.eof then 
		set sidx = objrs("sidx")
		set monthprice = objrs("monthprice")
		set expense = objrs("expense")
		set jobidx = objrs("jobidx")
	end if
	
	dim objrs2

	dim intLoop, lastmonth, photo_1, photo_2, photo_3, photo_4, M_startdate
	dim m_mdidx, m_title, m_locate, m_categoryidx, m_side, m_unit, m_unitprice, m_standard, m_quality, m_custcode, m_qty, m_trust, m_map

	do until objrs.eof 	
	
	sql = "select contidx, sidx, mdidx, title, locate, categoryidx, side, unit, unitprice, standard, quality, custcode, qty, trust, map, cuser, cdate, uuser, udate from dbo.WB_CONTACT_MD where contidx = "&org_contidx &" and sidx = " & sidx
	call set_recordset(objrs2, sql)

	m_mdidx = objrs2.fields("mdidx").value
	m_title = objrs2.fields("title").value
	m_locate = objrs2.fields("locate").value
	m_categoryidx = objrs2.fields("categoryidx").value
	m_side = objrs2.fields("side").value
	m_unit = objrs2.fields("unit").value
	m_unitprice = objrs2.fields("unitprice").value
	m_standard = objrs2.fields("standard").value
	m_quality=objrs2.fields("quality").value
	m_custcode = objrs2.fields("custcode").value
	m_qty = objrs2.fields("qty").value
	m_trust = objrs2.fields("trust").value
	m_map = objrs2.fields("qty").value

	objrs2.addnew			'입력할 데이터가 없는 경우에는 신규입력
		objrs2.fields("contidx").value = contidx
		objrs2.fields("mdidx").value = m_mdidx
		objrs2.fields("sidx").value = sidx
		objrs2.fields("title").value = m_title
		objrs2.fields("locate").value = m_locate
		objrs2.fields("categoryidx").value = m_categoryidx
		objrs2.fields("side").value = m_side
		objrs2.fields("unit").value = m_unit
		objrs2.fields("unitprice").value = m_unitprice
		objrs2.fields("standard").value = m_standard
		objrs2.fields("quality").value = m_quality
		objrs2.fields("custcode").value = m_custcode
		objrs2.fields("qty").value = m_qty
		objrs2.fields("trust").value = m_trust
		objrs2.fields("map").value = m_map
		objrs2.fields("cuser").value = request.cookies("userid")
		objrs2.fields("cdate").value = date
		objrs2.fields("uuser").value = request.cookies("userid")
		objrs2.fields("udate").value = date
		objrs2.update
		objrs2.close
		
		sql = "select contidx, sidx, cyear, cmonth, monthprice, expense, jobidx, photo_1, photo_2, photo_3, photo_4, perform from dbo.WB_CONTACT_MD_DTL where contidx = " & org_contidx &" and sidx="&sidx
		call set_recordset(objrs2, sql)
		M_startdate = startdate 
		lastmonth = DateDiff("m", M_startdate, enddate)
		
		for intLoop = 0 to lastmonth
			objrs2.addnew
			objrs2.fields("contidx").value = contidx
			objrs2.fields("sidx").value = sidx
			objrs2.fields("cyear").value = YEAR(M_startdate)
			objrs2.fields("cmonth").value = MONTH(M_startdate)
			objrs2.fields("jobidx").value = objrs2.fields("jobidx").value		
			objrs2.fields("photo_1").value = null
			objrs2.fields("photo_2").value = null
			objrs2.fields("photo_3").value = null
			objrs2.fields("photo_4").value = null
			objrs2.fields("perform").value = 0

			if intLoop = 0 then					' 계약 시작일자가 15일 이전인경우에는 청구금액을 입력하지 않는다.
				if day(M_startdate) > 1 and day(M_startdate) < 16 then
					objrs2.fields("monthprice").value = 0
					objrs2.fields("expense").value = 0
				else
					objrs2.fields("monthprice").value = monthprice
					objrs2.fields("expense").value = expense
				end if
			elseif intLoop = lastmonth then					' 계약 종료일자가 15일 이상인경우에는 청구금액을 입력하지 않는다.
				if day(enddate) <= 15 then
					objrs2.fields("monthprice").value = 0
					objrs2.fields("expense").value = 0
				else
					objrs2.fields("monthprice").value = monthprice
					objrs2.fields("expense").value = expense
				end if
			else
				objrs2.fields("monthprice").value = monthprice
				objrs2.fields("expense").value = expense
			end if

			objrs2.update
			M_startdate = DateAdd("m", 1, M_startdate)
		next
		objrs2.close

		sql = "select contidx, sidx, monthprice, expense, jobidx from dbo.wb_contact_tmp where contidx="&contidx&" and sidx="&sidx
		call set_recordset(objrs2, sql)

		objrs2.addnew
		objrs2.fields("contidx").value = contidx
		objrs2.fields("sidx").value = sidx
		objrs2.fields("monthprice").value = monthprice
		objrs2.fields("expense").value = expense
		objrs2.fields("jobidx").value = jobidx
		objrs2.update
		objrs2.close
	objrs.movenext
	loop

	set objrs2 = nothing

	set objrs = nothing
%>
<script language="JavaScript">
<!--
	window.opener.opener.location.reload();
	this.close();
//-->
</script>