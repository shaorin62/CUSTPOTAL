<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim custcode : custcode = request("selcustcode2")
	if custcode = "" then custcode = null
	dim gotopage : gotopage = request.querystring("gotopage")
	if gotopage = "" then gotopage = 1
	dim searchstring : searchstring = request("txtsearchstring")


	dim objrs, sql, menunum
	sql = "select c2.custname ,c.custname as custname2,  userid, class, isuse from dbo.wb_account a inner join dbo.sc_cust_temp c on a.custcode = c.custcode inner join dbo.sc_cust_temp c2 on c.highcustcode = c2.custcode where c2.custcode like '" & custcode &"%' and a.custcode like '"&custcode&"%' and userid like '%" & searchstring &"%' order by c.custcode, c2.custcode, userid "

	call get_recordset(objrs, sql)

	dim custname, custname2,  userid,  c_class, isuse, cnt
	if not objrs.eof then
		set custname = objrs("custname")
		set custname2 = objrs("custname2")
		set userid = objrs("userid")
		set c_class = objrs("class")
		set isuse = objrs("ISUSE")
		cnt = objrs.recordcount
	end if

%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<table width="1030" height="31" border="3" cellpadding="0" cellspacing="0" bordercolor="#8D652B">
                <tr>
                  <td><table width="1024" border="0" cellspacing="0" cellpadding="0" class="header">
                      <tr>
                        <td width="44" align="center" class="header">No</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="240" align="center">광고주</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="240" align="center">사업부</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="200" align="center" >아이디</td>
                        <td width="3" align="center" ><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="200" align="center" >권한</td>
                        <td width="3" align="center" ><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="100" align="center" >사용여부</td>
                      </tr>
                  </table></td>
                </tr>
              </table>
                <table width="1024" height="31" border="0" cellpadding="0" cellspacing="0" bordercolor="#8D652B">
				<% do until objrs.eof %>
                  <tr >
                    <td width="44" height="31" align="center"><%=cnt%></td>
                    <td width="3">&nbsp;</td>
                    <td width="240" align=""onClick="checkForView('<%=userid%>')" class="styleLink" style="padding-left:10px;"><%=custname%></td>
                    <td width="3">&nbsp;</td>
                    <td width="240" align=""onClick="checkForView('<%=userid%>')" class="styleLink" style="padding-left:10px;"><%=custname2%>&nbsp;</td>
                    <td width="3">&nbsp;</td>
                    <td width="200" align="left" onClick="checkForView('<%=userid%>')" class="styleLink header" style="padding-left:10px;"><%=userid%>&nbsp;</td>
                    <td width="3" align="center">&nbsp;</td>
                    <td width="200" align="left" onClick="checkForView('<%=userid%>')" class="styleLink" style="padding-left:10px;"><%if c_class = "A" then response.write "관리자" else response.write "일반사용자"%>&nbsp;</td>
                    <td width="3" align="center">&nbsp;</td>
                    <td width="100" align="center" "><%if isuse = "Y" then response.write "사용" Else response.write "중지"%>&nbsp;</td>
                  </tr>
                  <tr>
                    <td height="1" bgcolor="#E7E9E3" colspan="11"></td>
                  </tr>
				<%
						cnt = cnt - 1
						objrs.movenext
					loop
					objrs.close
					set objrs = nothing
				%>
            </table>