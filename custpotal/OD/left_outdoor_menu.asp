<%
	dim menunum : menunum = request("menunum")
	response.cookies("menunum") = menunum

%>
<table width="210" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="/images/menu_outdoor_sub_title.gif" width="210" height="75"></td>
  </tr>
  <tr>
    <td  width="210" height="25" class='menu' onclick="menuDisplay('outdoormenu1');">���ܱ�����Ȳ</td>
  </tr>
  <tr style="display:block;"  id="outdoormenu1">
    <td  width="210" >
		<table width="210" border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td  width="210" height="25" <%if request.cookies("menunum") = "1" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/od/outdoor/contact_list.asp?menuNum=1'">���ܱ��� ������Ȳ</span></td>
		  </tr>
		  <tr>
			<td  width="210" height="24"<%if request.cookies("menunum") = "2" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/od/outdoor/brand_list.asp?menuNum=2'">�귣�庰 ������Ȳ</span></td>
		  </tr>
		  <tr>
			<td  width="210" height="24" <%if request.cookies("menunum") = "4" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/od/outdoor/enddate_list.asp?menuNum=4'">�����Ϻ� ������Ȳ</span></td>
		  </tr>
		  <tr>
			<td  width="210" height="24" <%if request.cookies("menunum") = "9" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/od/outdoor/execution_list.asp?menuNum=9'">������ ������Ȳ</span></td>
		  </tr>
		 </table>
	</td>
  </tr>
  <tr>
    <td  width="210" height="25" class='menu' onclick="menuDisplay('outdoormenu2');">��ü����</td>
  </tr>
  <tr style="display:block;"  id="outdoormenu2">
    <td  width="210" >
		<table width="210" border="0" cellspacing="0" cellpadding="0">
<!-- 		  <tr>
			<td  width="210" height="24" <%'if request.cookies("menunum") = "5" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/hq/outdoor/medium_list.asp?menuNum=5'">��ü��Ȳ</span></td>
		  </tr> -->
		  <tr>
			<td  width="210" height="24" <%if request.cookies("menunum") = "6" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/od/outdoor/job_list.asp?menuNum=6'">������Ȳ</span></td>
		  </tr>
		  <tr>
			<td  width="210" height="24" <%if request.cookies("menunum") = "11" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/od/outdoor/category_list.asp?menuNum=11'">��ü�з�</span></td>
		  </tr>
		 </table>
	</td>
  </tr>
<!--   <tr>
    <td  width="210" height="24" <%'if request.cookies("menunum") = "7" then Response.write "class='menuover'" Else response.write "class='menu'" End If%>><span onclick="location.href='/hq/outdoor/validation_tool_list.asp?menuNum=7'">ȿ�뼺Tool</span></td>
  </tr> -->
  <tr>
    <td  width="210" height="24" <%if request.cookies("menunum") = "8" then Response.write "class='menuover'" Else response.write "class='menu'" End If%>><span onclick="location.href='/od/outdoor/monitor_list.asp?menuNum=8'">���ܸ���͸�</span></td>
  </tr>
  <tr>
    <td  width="210" height="24" <%if request.cookies("menunum") = "10" then Response.write "class='menuover'" Else response.write "class='menu'" End If%>><span onclick="window.open('/med/','pop_medium','')">��������</span></td>
  </tr>
  <tr>
    <td><img src="/images/menu_sub_bottom.gif" width="210" height="30"></td>
  </tr>
</table>
<script language="JavaScript">
<!--
	var menuFlag = true ;
	function menuDisplay(id) {
//		if (menuFlag) value = "block";
//		else value="none";
//		document.getElementById(id).style.display = value;
//		menuFlag =  !menuFlag;
	}
//-->
</script>
