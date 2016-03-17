<table width="210" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="/images/menu_trans_sub_title.gif" width="210" height="75"></td>
  </tr>
  <%
	Dim objmenu
	sql = "select custcode, custname from dbo.sc_cust_temp where custcode = highcustcode and medflag = 'A' "
	Call get_recordset(objmenu, sql)

	Dim menu_custcode, menu_custname
	If Not objmenu.eof Then 
		Set menu_custcode = objmenu("custcode")
		Set menu_custname = objmenu("custname")
	End If
	If menunum = "" Then menunum = "1"
	
	Do Until objmenu.eof 
  %>
  <tr>
    <td  width="210" height="25" class='headermenu' <%if menunum= "1" And custcode = menu_custcode then Response.write "class='menu'" Else response.write "class='menuover'" End If%>><span onclick="location.href='/hq/trans/public_01_list.asp?menuNum=1&custcode=<%=menu_custcode%>'"><%=menu_custname%></td>
  </tr>
  <%  
		If CStr(menu_custcode) = CStr(request("custcode")) Then 
  %>
  <tr>
    <td  width="210" height="24" class='subheadermenu' >광고비 집행 종합</td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "1" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/public_01_list.asp?menuNum=1&custcode=<%=menu_custcode%>'">월별 매체별 광고비</span></td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "2" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/public_02_list.asp?menuNum=2&custcode=<%=menu_custcode%>'">월별 세부 매체별 광고비</span></td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "3" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/public_03_list.asp?menuNum=3&custcode=<%=menu_custcode%>'">CATV/ New Media 내역</span></td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "4" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/public_04_list.asp?menuNum=4&custcode=<%=menu_custcode%>'">ATL 브랜드/소재별 광고비</span></td>
  </tr>
  <tr>
    <td  width="210" height="24" class='subheadermenu' >매체별 집행내역</td>
  </tr>
  <tr>
    <td  width="210" height="22" class='menulist' ><B style="color:#FFCC33;">| </B>공중파</td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "5" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/medium_01_list.asp?menuNum=5&custcode=<%=menu_custcode%>'">AOR 공중파 광고정산</span></td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "6" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/medium_02_list.asp?menuNum=6&custcode=<%=menu_custcode%>'">부문별 실집행 광고비</span></td>
  </tr>
  <tr>
    <td  width="210" height="22" class='menulist' ><B style="color:#FFCC33;">| </B> 케이블</td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "7" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/medium_03_list.asp?menuNum=7&custcode=<%=menu_custcode%>'">부문별 실집행 광고비</span></td>
  </tr>
  <tr>
    <td  width="210" height="22" class='menulist' ><B style="color:#FFCC33;">| </B> 인쇄</td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "8" And custcode = menu_custcode then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/medium_04_list.asp?menuNum=8&custcode=<%=menu_custcode%>'">부문별 실집행 광고비</span></td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "9" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/medium_05_list.asp?menuNum=9&custcode=<%=menu_custcode%>'">매체별 집행내역</span></td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "10" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/medium_06_list.asp?menuNum=10&custcode=<%=menu_custcode%>'">세부 집행내역</span></td>
  </tr>
  <tr>
    <td  width="210" height="22" class='menulist' ><B style="color:#FFCC33;">| </B> 인터넷</td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "11" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/medium_07_list.asp?menuNum=11&custcode=<%=menu_custcode%>'">광고주/부문별 큐시트</span></td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "12" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/medium_08_list.asp?menuNum=12&custcode=<%=menu_custcode%>'">인터넷 매체별 광고비</span></td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "13" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/medium_09_list.asp?menuNum=13&custcode=<%=menu_custcode%>'">광고주/CIC별 매체비</span></td>
  </tr>
  <tr>
    <td  width="210" height="22" class='menulist' ><B style="color:#FFCC33;">| </B> 옥외</td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "14" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/medium_10_list.asp?menuNum=14&custcode=<%=menu_custcode%>'">옥외광고 현황</span></td>
  </tr>
  <%
		End if
	objmenu.movenext
	Loop
		
	objmenu.close
	Set objmenu = Nothing
	
  %>
  <tr>
    <td><img src="/images/menu_sub_bottom.gif" width="210" height="30"></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
