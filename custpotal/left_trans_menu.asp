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
    <td  width="210" height="24" class='subheadermenu' >����� ���� ����</td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "1" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/public_01_list.asp?menuNum=1&custcode=<%=menu_custcode%>'">���� ��ü�� �����</span></td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "2" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/public_02_list.asp?menuNum=2&custcode=<%=menu_custcode%>'">���� ���� ��ü�� �����</span></td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "3" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/public_03_list.asp?menuNum=3&custcode=<%=menu_custcode%>'">CATV/ New Media ����</span></td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "4" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/public_04_list.asp?menuNum=4&custcode=<%=menu_custcode%>'">ATL �귣��/���纰 �����</span></td>
  </tr>
  <tr>
    <td  width="210" height="24" class='subheadermenu' >��ü�� ���೻��</td>
  </tr>
  <tr>
    <td  width="210" height="22" class='menulist' ><B style="color:#FFCC33;">| </B>������</td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "5" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/medium_01_list.asp?menuNum=5&custcode=<%=menu_custcode%>'">AOR ������ ��������</span></td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "6" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/medium_02_list.asp?menuNum=6&custcode=<%=menu_custcode%>'">�ι��� ������ �����</span></td>
  </tr>
  <tr>
    <td  width="210" height="22" class='menulist' ><B style="color:#FFCC33;">| </B> ���̺�</td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "7" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/medium_03_list.asp?menuNum=7&custcode=<%=menu_custcode%>'">�ι��� ������ �����</span></td>
  </tr>
  <tr>
    <td  width="210" height="22" class='menulist' ><B style="color:#FFCC33;">| </B> �μ�</td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "8" And custcode = menu_custcode then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/medium_04_list.asp?menuNum=8&custcode=<%=menu_custcode%>'">�ι��� ������ �����</span></td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "9" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/medium_05_list.asp?menuNum=9&custcode=<%=menu_custcode%>'">��ü�� ���೻��</span></td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "10" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/medium_06_list.asp?menuNum=10&custcode=<%=menu_custcode%>'">���� ���೻��</span></td>
  </tr>
  <tr>
    <td  width="210" height="22" class='menulist' ><B style="color:#FFCC33;">| </B> ���ͳ�</td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "11" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/medium_07_list.asp?menuNum=11&custcode=<%=menu_custcode%>'">������/�ι��� ť��Ʈ</span></td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "12" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/medium_08_list.asp?menuNum=12&custcode=<%=menu_custcode%>'">���ͳ� ��ü�� �����</span></td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "13" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/medium_09_list.asp?menuNum=13&custcode=<%=menu_custcode%>'">������/CIC�� ��ü��</span></td>
  </tr>
  <tr>
    <td  width="210" height="22" class='menulist' ><B style="color:#FFCC33;">| </B> ����</td>
  </tr>
  <tr>
	<td  width="210" height="19" <%if menunum= "14" And custcode = menu_custcode  then Response.write "class='menulistover2'" Else response.write "class='menulist2'" End If%>><span onclick="location.href='/hq/trans/medium_10_list.asp?menuNum=14&custcode=<%=menu_custcode%>'">���ܱ��� ��Ȳ</span></td>
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
