<%@ language="vbscript" codepage="65001" %>
<!--#include virtual="/inc/func.asp" -->
<%

	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1000

	' 아이디 검색
	dim findID : findID = request("findID")
	' 이름 검색
	dim findNAME : findNAME = request("findNAME")


%>


<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<div style='margin-top:10px;'>
<TABLE  width="100%">
	<TR>
		<TD ><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle" id="subtitle"> 계정관리</span></TD>
		<TD  align="right" valign="top">  <span class="navigator"  id="navigate">관리모드 &gt; 계정관리</span></TD>
	</TR>
</TABLE>
</div>


<!-- 검색 영역 -->
	<Div id='searchtag' style='margin-top:10px;'>
		<table width="1030" height="35" border="0" cellpadding="0" cellspacing="0" class="header">
			<tr>
				  <td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
				  <td width="50%" align="left" background="/images/bg_search.gif">
				  아이디 : <input type="text" name="txtfindID" value="<%=findID%>">
				  이름 : <input type="text" name="txtfindNAME" value="<%=findNAME%>">
				  <A HREF="#" >
				  <img src="/images/btn_search.gif" width="39" height="20" align="absmiddle" border="0" onclick="getdata('account.asp'); return false;"></A></td>
				  <td width="50%" align="right" background="/images/bg_search.gif"><img src="/images/btn_acc_reg.gif" width="78" height="18" alt="" border="0" class="account" onclick="pop_reg();" id="btnReg" style="cursor:hand;"></td>
				  <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
			</tr>
		</table>
	</div>


<!-- 컨텐츠 영역 -->
<%
	dim sql , objrs, objrs2 , cnt
	dim highcustname, custname,  userid, username,  c_class, isuse

	sql = "select count(userid) cnt from wb_Account where  userid like '%" & findID & "%' and isnull(username,'') like '%" & findNAME & "%'"
	Call get_recordset(objrs2, sql)
	if not objrs2.eof then
		cnt = objrs2("cnt")
	end If

%>
<div id='#contents' style='margin-top:10px;width:1040px;overflow-x:scroll;'>

<link href="/style.css" rel="stylesheet" type="text/css">

<table width="1030" height="31" border="3" cellpadding="0" cellspacing="0" bordercolor="#8D652B">
                <tr>
                  <td><table width="1024" border="0" cellspacing="0" cellpadding="0" class="header">
                      <tr>
                        <td width="44" align="center" class="header">No</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="240" align="center">광고주</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="240" align="center">운영팀</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="150" align="center" >아이디</td>
                        <td width="3" align="center" ><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
						 <td width="100" align="center" >이름</td>
                        <td width="3" align="center" ><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="150" align="center" >권한</td>
                        <td width="3" align="center" ><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="100" align="center" >사용여부</td>
                      </tr>
                  </table></td>
                </tr>
              </table>
                <table width="1024" border="0" cellpadding="0" cellspacing="0" bordercolor="#8D652B">
				<%

					sql = "select c.custname as highcustname, b.custname, a.userid, a.username, a.class, a.isuse "
					sql = sql &  " from wb_account a  inner join sc_cust_dtl b on a.custcode=b.custcode  "
					sql = sql &  " inner join sc_cust_hdr c on b.highcustcode=c.highcustcode  "
					sql = sql &  " where userid like '%" & findID & "%' and isnull(username,'') like '%" & findNAME & "%'"
					sql = sql &  " union all "
					sql = sql &  " select '옥외모니터링' as highcustname, '옥외모니터링' custname, userid, username, class, isuse "
					sql = sql &  " from wb_account "
					sql = sql &  " where class = 'F'  and userid like '%" & findID & "%' and isnull(username,'') like '%" & findNAME & "%'"
					sql = sql & " order by a.class  "

					Call get_recordset(objrs, sql)

					if not objrs.eof then
						do until objrs.eof
							highcustname = objrs("highcustname")
							custname = objrs("custname")
							userid = objrs("userid")
							username = objrs("username")
							c_class = objrs("class")
							isuse = objrs("isuse")
					%>
                  <tr >
                    <td width="44" height="31" align="center"><%=cnt%></td>
                    <td width="3">&nbsp;</td>
                    <td width="240" align=""onClick="checkForView('<%=userid%>','<%=c_class%>')" class="styleLink" style="padding-left:10px;"><%=highcustname%></td>
                    <td width="3">&nbsp;</td>
                    <td width="240" align=""onClick="checkForView('<%=userid%>','<%=c_class%>')" class="styleLink" style="padding-left:10px;"><%=custname%>&nbsp;</td>
                    <td width="3">&nbsp;</td>
                    <td width="150" align="left" onClick="checkForView('<%=userid%>','<%=c_class%>')" class="styleLink header" style="padding-left:10px;"><%=userid%>&nbsp;</td>
                    <td width="3" align="center">&nbsp;</td>
					<td width="100" align="left" onClick="checkForView('<%=userid%>','<%=c_class%>')" class="styleLink header" style="padding-left:10px;"><%=username%>&nbsp;</td>
                    <td width="3" align="center">&nbsp;</td>
                    <td width="150" align="left" onClick="checkForView('<%=userid%>','<%=c_class%>')" class="styleLink" style="padding-left:10px;">
					<%
					select case c_class
						case  "A"	response.write "Administrator"
						case  "N"	response.write "Admin(Non-SKT)"
						case  "C"	response.write "광고주"
						case  "G"	response.write "광고주 관리자"
						case  "D"	response.write "운영팀"
						case  "H"	response.write "운영팀 관리자"
						case  "O"	 response.write "옥외 관리자"
						case  "F"	 response.write "옥외 모니터링"
						case  "M"	 response.write "매체사"
					end select
				%></td>
                    <td width="3" align="center">&nbsp;</td>
                    <td width="100" align="center"><%if ucase(isuse) = "Y" then response.write "사용중" Else response.write "사용중지"%>&nbsp;</td>
                  </tr>
                  <tr>
                    <td height="1" bgcolor="#E7E9E3" colspan="13"></td>
                  </tr>
				<%
							cnt = cnt - 1
							objrs.movenext
						loop
					end If

					objrs.close
					set objrs = nothing
				%>
            </table>
</div>


