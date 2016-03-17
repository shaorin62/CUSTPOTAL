<%@ language="vbscript" codepage="65001" %>
<!--#include virtual="/inc/func.asp" -->

<%
	response.cookies("cookiemidx") = request("midx")
	response.cookies("cookiecustcode") = request("custcode")
	response.cookies("cookiesearchstring") = request("searchstring")
	response.cookies("cookiehighcategory") = request("highcategory")
	response.cookies("cookiecategory") = request("category")
	response.cookies("cookieattr02") = request("attr02")

	dim objrs, objrs2, sql

	dim gotopage : gotopage = request.QueryString("gotopage")

	if gotopage = "" then gotopage = 1

	dim searchstring : searchstring = request("searchstring")

	dim pagesize : pagesize = 10

	dim flag : flag = request("flag")

	dim midx : midx = request("midx")
	dim highcategory : highcategory = request("highcategory")
	dim category : category = request("category")

	dim strcustcode : strcustcode = request("ccustcode")
	dim strtitle
	dim stryear : stryear = request("cyear")
	dim strmon : strmon = request("cmonth")

	if midx = "" then
		sql = "select min(midx) from dbo.wb_menu_mst where custcode is null"
		call get_recordset(objrs, sql)
		if not objrs.eof then
			midx = objrs(0)
		else
			midx = 0
		end if
		objrs.close
	end if

	sql = "select title from dbo.wb_menu_mst where midx=" & midx
	call get_recordset(objrs, sql)
	if not objrs.eof then
		strtitle = objrs("title")
	end if
	objrs.close
	response.cookies("cookietitlename") = escape(strtitle)

	Dim conhighcategory
	Dim concategory
	Dim constrcustcode
	Dim constryear
	Dim constrmon

	If highcategory <> "" Then
		conhighcategory =  " and r.highcategory = " & highcategory
	Else
		conhighcategory =  " "
	End If

	If category <> "" Then
		concategory =  " and r.category = " & category
	Else
		concategory =  " "
	End If

	If strcustcode <> "" Then
		constrcustcode =  " and r.custcode = '" & strcustcode & "'"
	Else
		constrcustcode =  " "
	End If

	If stryear <> "" Then
		constryear = " and cyear = '" & stryear & "' "
	Else
		constryear = ""
	End if

	If strmon <> "" Then
		constrmon = " and cmonth = '" & strmon & "' "
	Else
		constrmon = ""
	End If

	'과거
	'sql = "select count(*) from dbo.wb_report r inner join dbo.wb_menu_mst m on r.midx = m.midx left outer join dbo.sc_cust_temp c on m.custcode = c.custcode where r.midx="&midx&" and r.title like '%" & searchstring & "%' ; select top "&pagesize&" r.ridx, r.title, r.cuser, r.cdate from dbo.wb_report r inner join dbo.wb_menu_mst m on r.midx = m.midx left outer join dbo.sc_cust_temp c on m.custcode = c.custcode where r.midx="&midx&" and r.title like '%" & searchstring & "%'  and r.ridx not in (select top "&(gotopage-1)*pagesize&" r.ridx from dbo.wb_report r inner join dbo.wb_menu_mst m on r.midx = m.midx left outer join dbo.sc_cust_temp c on m.custcode = c.custcode where r.midx="&midx&" and r.title like '%" & searchstring & "%' order by r.ridx desc) order by r.ridx desc"

    '수정본  (midx가있어서  cust_temp조인 뺐음 )
	'sql = "select count(*) from dbo.wb_report r inner join dbo.wb_menu_mst m on r.midx = m.midx where r.midx="&midx&" and r.title like '%" & searchstring & "%' ; select top "&pagesize&" r.ridx, r.title, r.cuser, r.cdate from dbo.wb_report r inner join dbo.wb_menu_mst m on r.midx = m.midx where r.midx="&midx&" and r.title like '%" & searchstring & "%'  and r.ridx not in (select top "&(gotopage-1)*pagesize&" r.ridx from dbo.wb_report r inner join dbo.wb_menu_mst m on r.midx = m.midx  where r.midx="&midx&" and r.title like '%" & searchstring & "%' order by r.ridx desc) order by r.ridx desc"

	If midx = 1 Then
		sql = " select count(*)   "
		sql = sql & " from (   "
		sql = sql & " 	select ridx, title, cuser, cdate    "
		sql = sql & " 	from (   "
		sql = sql & " 		select ridx, title, cuser, cdate    "
		sql = sql & " 		from dbo.wb_report   "
		sql = sql & " 		where cdate  >= dateadd(d,-7,getdate())   "
		sql = sql & " 		and title like '%" & searchstring & "%'   "
		sql = sql & " "		& conhighcategory
		sql = sql & " "		& concategory
		sql = sql & " 		union all   "
		sql = sql & " 		select r.ridx, r.title, r.cuser, r.cdate    "
		sql = sql & " 		from dbo.wb_report r   "
		sql = sql & " 		inner join dbo.wb_menu_mst m on r.midx = m.midx      "
		sql = sql & " 		where r.midx= 1   "
		sql = sql & " "		& conhighcategory
		sql = sql & " "		& concategory
		sql = sql & " 		and r.title like '%" & searchstring & "%'   "
		sql = sql & " 	) a group by ridx, title, cuser, cdate    "
		sql = sql & " ) b;     "

		sql = sql & " select top "&pagesize&"  ridx, title, cuser, cdate, cnt "
		sql = sql & " from (   "
		sql = sql & " 	select ridx, title, cuser, cdate, cnt    "
		sql = sql & " 	from (   "
		sql = sql & " 		select ridx, title, cuser, cdate, cast(isnull(cnt,0) as varchar(20)) + '/' + cast(isnull(downcnt,0) as varchar(20)) cnt    "
		sql = sql & " 		from dbo.wb_report   "
		sql = sql & " 		where cdate  >= dateadd(d,-7,getdate())   "
		sql = sql & " 		and title like '%" & searchstring & "%'   "
		sql = sql & " "		& conhighcategory
		sql = sql & " "		& concategory
		sql = sql & " 		union all   "
		sql = sql & " 		select r.ridx, r.title, r.cuser, r.cdate, cast(isnull(r.cnt,0) as varchar(20)) + '/' + cast(isnull(r.downcnt,0) as varchar(20)) cnt    "
		sql = sql & " 		from dbo.wb_report r   "
		sql = sql & " 		inner join dbo.wb_menu_mst m on r.midx = m.midx      "
		sql = sql & " 		where r.midx= 1   "
		sql = sql & " "		& conhighcategory
		sql = sql & " "		& concategory
		sql = sql & " 		and r.title like '%" & searchstring & "%'   "
		sql = sql & " 	) a group by ridx, title, cuser, cdate, cnt    "
		sql = sql & " ) b    "
		sql = sql & " where ridx not in(   "
		sql = sql & " 	select top  "&(gotopage-1)*pagesize&" ridx   "
		sql = sql & " 	from (   "
		sql = sql & " 		select ridx   "
		sql = sql & " 		from dbo.wb_report   "
		sql = sql & " 		where cdate  >= dateadd(d,-7,getdate())   "
		sql = sql & " 		and title like '%" & searchstring & "%'   "
		sql = sql & " "		& conhighcategory
		sql = sql & " "		& concategory
		sql = sql & " 		union all   "
		sql = sql & " 		select r.ridx   "
		sql = sql & " 		from dbo.wb_report r   "
		sql = sql & " 		inner join dbo.wb_menu_mst m on r.midx = m.midx      "
		sql = sql & " 		where r.midx= 1   "
		sql = sql & " "		& conhighcategory
		sql = sql & " "		& concategory
		sql = sql & " 		and r.title like '%" & searchstring & "%'   "
		sql = sql & " 	) a group by ridx   "
		sql = sql & " 	order by ridx desc    "
		sql = sql & " )     "
		sql = sql & " order by ridx desc ;   "

	Else
		sql = " select count(*)  "
		sql = sql & " from dbo.wb_report r  "
		sql = sql & " inner join dbo.wb_menu_mst m on r.midx = m.midx  "
		sql = sql & " where r.midx="&midx&" "
		sql = sql & " and r.title like '%" & searchstring & "%'   "
		sql = sql & " " & conhighcategory
		sql = sql & " " & concategory
		sql = sql & " " & constrcustcode
		sql = sql & " " & constryear
		sql = sql & " " & constrmon & ";"

		sql = sql & " select top "&pagesize&" r.ridx, r.title, r.cuser, r.cdate, cast(isnull(r.cnt,0) as varchar(20)) + '/' + cast(isnull(r.downcnt,0) as varchar(20)) cnt "
		sql = sql & " from dbo.wb_report r  "
		sql = sql & " inner join dbo.wb_menu_mst m on r.midx = m.midx  "
		sql = sql & " where r.midx="&midx&" "
		sql = sql & " " & conhighcategory
		sql = sql & " " & concategory
		sql = sql & " " & constrcustcode
		sql = sql & " " & constryear
		sql = sql & " " & constrmon
		sql = sql & " and r.title like '%" & searchstring & "%'   "
		sql = sql & " and r.ridx not in ( "
		sql = sql & " 	select top "&(gotopage-1)*pagesize&" r.ridx  "
		sql = sql & " 	from dbo.wb_report r  "
		sql = sql & " 	inner join dbo.wb_menu_mst m on r.midx = m.midx   "
		sql = sql & " 	where r.midx="&midx&" "
		sql = sql & " " & conhighcategory
		sql = sql & " " & concategory
		sql = sql & " " & constrcustcode
		sql = sql & " " & constryear
		sql = sql & " " & constrmon
		sql = sql & " 	and r.title like '%" & searchstring & "%'  "
		sql = sql & " 	order by r.ridx desc "
		sql = sql & " )  "
		sql = sql & " order by r.ridx desc "
	End if

	call get_recordset(objrs, sql)

	dim totalrecord : totalrecord = objrs(0).value

	set objrs = objrs.nextrecordset

	dim ridx, title, attachfile, c_date, c_user, attachfile2, attachfile3, cnt
	if not objrs.eof then
		set ridx = objrs("ridx") 	' idx
		set title = objrs("title")	 ' title
		set c_user = objrs("cuser")	 ' cuser
		set c_date = objrs("cdate") ' cdate
		set cnt = objrs("cnt") ' cdate
	end if

%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<body oncontextmenu="return false">

<table width="1020" height="31" border="3" cellpadding="0" cellspacing="0" align="center" bordercolor="#8D652B">
	<tr>
	  <td>
		  <table width="1018" border="0" cellspacing="0" cellpadding="0" class="header" align="center">
			  <tr>
				<td width="43" align="center" class="header">No</td>
				<td width="3" align="center" ><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
				<td width="680" align="center" >제 목</td>
				<td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
				<td width="100" align="center">작성자</td>
				<td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
				<td width="100" align="center">작성일</td>
				<td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
				<td width="60" align="center">클릭수</td>
			  </tr>
		  </table>
	  </td>
	</tr>
</table>
<table width="1018" height="31" border="0" cellpadding="0" cellspacing="0" align="center">
	<%
		dim num
		num = totalrecord - ((gotopage-1)*pagesize)
		do until objrs.eof
	%>
  <tr >
    <td width="43" height="31" align="center"><%=num%></td>
    <td width="3" align="center">&nbsp;</td>
    <td width="680" align="left" style="padding-left:20px;">
	<%if datediff("d", c_date, date) < 2 then %><img src="/images/new.gif" width="21" height="10" border="0" alt=""><%end if%> &nbsp;&nbsp;<span onClick="pop_report_view('<%=objrs("ridx")%>');pop_report_cntupdate('<%=objrs("ridx")%>');" class="styleLink"><%=title%></span>
	<%
		sql = "select idx, ridx, attachfile from dbo.wb_Report_pds where ridx = " & ridx

		call get_Recordset(objrs2, sql)
		dim count : count = 1
		if not objrs2.eof then
		do until objrs2.eof
		if count < 4 then
	%>
	<span onClick="checkForDownload('<%=objrs2("attachfile")%>');pop_report_downcntupdate('<%=objrs("ridx")%>');" class="styleLink" style="padding-left:10px;"> <img src="/images/ico_attach.gif" width="7" height="12" hspace="3"> <%=objrs2("attachfile")%></span>
	<%
		else
		response.write " ..."
		exit do
		end if
		count = count + 1
		objrs2.movenext
		Loop

		end if
		objrs2.close
		set objrs2 = nothing
	%>
	</td>
    <td width="3">&nbsp;</td>
    <td width="100" align="center"><%=c_user%></td>
    <td width="3">&nbsp;</td>
    <td width="100" align="center"><%=formatdatetime(c_date,2)%></td>
	<td width="60" align="center"><%=cnt%></td>
  </tr>
  <tr>
    <td height="1" bgcolor="#E7E9E3" colspan="11"></td>
  </tr>
<%
	sql = "select cidx, ridx, comment, attachfile, cuser, cdate from dbo.wb_report_comment where ridx="&ridx & " order by cidx desc"
		call get_recordset(objrs2, sql)
		dim comment, c_attachfile, c_user2, c_date2
		if not objrs2.eof then
			set comment = objrs2("comment")
			set c_attachfile = objrs2("attachfile")
			set c_user2 = objrs2("cuser")
			set c_date2= objrs2("cdate")
		end if
	do until objrs2.eof
%>
			  <tr ><!-- 댓글목록 -->
				<td width="43" height="31" align="center"></td>
				<td width="3" align="center">&nbsp;</td>
				<td width="680" align="left" style="padding-left:20px;"><img src="/images/icon_reply.gif" width="28" height="16" hspace="3" align="bottom"  ><span onClick="pop_report_view('<%=objrs("ridx")%>');" class="styleLink"><%=comment%></span> <% if not isnull(c_attachfile) then %><span onClick="checkForDownload('<%=c_attachfile%>');" class="styleLink" style="padding-left:50px;font-size:11px;font-fmaily:돋움"> <img src="/images/ico_attach.gif" width="7" height="12" hspace="3"><%if len(c_attachfile) > 40 then response.write left(c_attachfile, 37)&"..." else response.write c_attachfile%></span><% end if%> </td>
				<td width="3">&nbsp;</td>
				<td width="100" align="center"><%=c_user2%></td>
				<td width="3">&nbsp;</td>
				<td width="100" align="center"><%=formatdatetime(c_date2,2)%></td>
				<td width="60" align="center"><%=formatdatetime(c_date2,2)%></td>
			  </tr>
		  	  <tr>
				<td height="1" bgcolor="#E7E9E3" colspan="11"></td>
			  </tr>
<%
		objrs2.movenext
		loop
		objrs2.close

		num = num -1
		objrs.movenext
		loop
		objrs.close

		set objrs = nothing
		set objrs2 = nothing
	%>
  <tr>
    <td height="40"  colspan="11" class="pagesplit" align="center" valign="bottom"><%call boardpagesplit(totalrecord, gotopage, pagesize, searchstring,midx, strcustcode, strtitle)%></td>
  </tr>
</table>
</body>
