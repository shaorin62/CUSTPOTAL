<%@ language="vbscript" codepage="65001" %>

<!--#include virtual="/inc/func.asp" -->
<%

	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1000


	dim custcode : custcode = request.querystring("custcode")

	if custcode = "" then custcode = Null
	dim FLAG : FLAG = request.querystring("FLAG")
'	response.write FLAG
'	response.end


	dim objrs1, sql1 , objrs2, sql2
	dim custname , highcustname



	If FLAG = "HIGHCUSTCODE" THEN
		sql1 = "select custname from dbo.sc_cust_HDR where highcustcode = '" &custcode & "'"

		call get_recordset(objrs1, sql1)

		if not objrs1.eof then
			set highcustname = objrs1("custname")
		end If

	Else

		sql2 = "select dbo.sc_get_highcustname_fun(highcustcode) highcustname,  custname from dbo.sc_cust_DTL where custcode = '" &custcode & "'"

		call get_recordset(objrs2, sql2)

		if not objrs2.eof Then
			set highcustname = objrs2("highcustname")
			set custname = objrs2("custname")
		end If

	End If

%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">


<div style='margin-top:10px;'>
<TABLE  width="100%">
	<TR>
		<TD ><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle" id="subtitle">
			<%If highcustname = ""  then  %>
			공통메뉴
			<%Else
					If FLAG = "CUSTCODE" Then
						RESPONSE.WRITE highcustname & " > "  & custname
					Else
						RESPONSE.WRITE highcustname
					End If
			End If %>
		</span>
		</TD>
		<TD  align="right" valign="top">  <span class="navigator"  id="navigate">관리모드 &gt; 메뉴관리</span></TD>
	</TR>
</TABLE>
</div>


<!-- 검색 영역 -->
	<Div id='searchtag' style='margin-top:10px;'>
		<table width="1030" height="35" border="0" cellpadding="0" cellspacing="0">
			<tr>
				  <td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
				  <td width="50%" align="left" background="/images/bg_search.gif"></td>
				  <td width="50%" align="right" background="/images/bg_search.gif"><img src="/images/btn_menu_reg.gif" width="78" height="18" alt="" border="0" class="account" onclick="pop_menu_reg();" id="btnReg" style="cursor:hand;"></td>
				  <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
			</tr>
		</table>
	</div>

<!-- 컨텐츠 영역 -->
<%

	'sql = "select m.midx, m.title, c.custname, m.lvl, m.isfile, iscomment, isemail from dbo.wb_menu_mst m left outer join dbo.sc_cust_temp c on m.custcode = c.custcode where m.custcode is null order by  m.ref , m.lvl"


	If FLAG = "HIGHCUSTCODE" Then
		sql = "select m.midx, m.title, c.custname, m.lvl, m.isfile, iscomment, isemail, mp from dbo.wb_menu_mst m inner join dbo.sc_cust_hdr c on m.custcode = c.highcustcode where isnull(m.attr01,'') =''  and  m.custcode = '"&custcode&"' order by  m.ref  , m.lvl"
	ElseIf FLAG ="CUSTCODE" THEN
		sql = "select m.midx, m.title, c.custname, m.lvl, m.isfile, iscomment, isemail, mp from dbo.wb_menu_mst m inner join dbo.sc_cust_dtl c on m.custcode = c.custcode where isnull(m.attr01,'') ='1' and  m.custcode = '"&custcode&"' order by  m.ref  , m.lvl"
	Else
		sql = " select midx, title, '' custname, lvl, isfile, iscomment, isemail, mp from dbo.wb_menu_mst where custcode is null  "
	End IF


	dim objrs, sql
	call get_recordset(objrs, sql)

	dim midx, title, isfile, isemail, iscomment, lvl
	if not objrs.eof then
		set midx = objrs("midx")
		set title = objrs("title")
		set isfile = objrs("isfile")
		set isemail = objrs("isemail")
		set iscomment = objrs("iscomment")
		set ismp = objrs("mp")
		set lvl = objrs("lvl")
	end if
%>

<div id='#contents' style='margin-top:10px;width:1040px;overflow-x:scroll;'>
<link href="/style.css" rel="stylesheet" type="text/css">

<table width="1030" height="31" border="3" cellpadding="0" cellspacing="0" bordercolor="#8D652B" align="center">
	<tr>
	  <td>
		<table width="1024" border="0" cellspacing="0" cellpadding="0" class="header" align="center">
		  <tr>
			<td width="524" align="center" >메뉴명</td>
			<td width="3" align="center" ><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
			<td width="100" align="center" >첨부파일</td>
			<td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
			<td width="100" align="center" >메일발송</td>
			<td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
			<td width="100" align="center" >댓글기능</td>
			<td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
			<td width="100" align="center" >내부(MP)용</td>
			<td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
			<td width="110" align="center" >하위메뉴</td>
		  </tr>
	  </table>
	  </td>
	</tr>
 </table>


<table width="1024"  border="0" cellpadding="0" cellspacing="0" bordercolor="#8D652B" style="margin-left:3px;">
	<% do until objrs.eof %>
	  <tr class="styleLink" height="31">
		<td width="624" align="left"  class="styleLink" style="padding-left:20px;" onClick="go_menu_view('<%=midx%>')" ><%if lvl = 2 then %><img src="/images/tree-branch.gif" width="19" height="14" border="0" alt="" hspace="5"> <%end if%><%=title%>&nbsp;</td>
		<td width="3" align="center">&nbsp;</td>
		<td width="100" align="center"><%if isfile then response.write "사용"%>&nbsp;</td>
		<td width="3">&nbsp;</td>
		<td width="100" align="center"><%if isemail then response.write "사용"%>&nbsp;</td>
		<td width="3">&nbsp;</td>
		<td width="100" align="center"><%if iscomment then response.write "사용"%>&nbsp;</td>
		<td width="3">&nbsp;</td>
		<td width="100" align="center"><%if ismp then response.write "사용"%>&nbsp;</td>
		<td width="3">&nbsp;</td>
			<% If lvl = 1 Then %>
			<td width="110" align="center" onClick="go_submenu_reg('<%=midx%>')" ><%if lvl = 1 then %><img src="/images/btn_submeun_reg.gif" width="100" height="18" border="0" alt="">
			<%End If %>
		<%end if%></td>
	  </tr>
	  <tr>
		<td height="1" bgcolor="#E7E9E3" colspan="13"></td>
	  </tr>
	<%
			objrs.movenext
		loop
		objrs.close
		set objrs = nothing
	%>
</table>
</div>


