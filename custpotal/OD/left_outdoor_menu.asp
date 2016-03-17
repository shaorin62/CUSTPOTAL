<%
	dim menunum : menunum = request("menunum")
	response.cookies("menunum") = menunum

%>
<table width="210" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="/images/menu_outdoor_sub_title.gif" width="210" height="75"></td>
  </tr>
  <tr>
    <td  width="210" height="25" class='menu' onclick="menuDisplay('outdoormenu1');">옥외광고현황</td>
  </tr>
  <tr style="display:block;"  id="outdoormenu1">
    <td  width="210" >
		<table width="210" border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td  width="210" height="25" <%if request.cookies("menunum") = "1" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/od/outdoor/contact_list.asp?menuNum=1'">옥외광고 집행현황</span></td>
		  </tr>
		  <tr>
			<td  width="210" height="24"<%if request.cookies("menunum") = "2" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/od/outdoor/brand_list.asp?menuNum=2'">브랜드별 집행현황</span></td>
		  </tr>
		  <tr>
			<td  width="210" height="24" <%if request.cookies("menunum") = "4" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/od/outdoor/enddate_list.asp?menuNum=4'">종료일별 광고현황</span></td>
		  </tr>
		  <tr>
			<td  width="210" height="24" <%if request.cookies("menunum") = "9" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/od/outdoor/execution_list.asp?menuNum=9'">광고비용 집행현황</span></td>
		  </tr>
		 </table>
	</td>
  </tr>
  <tr>
    <td  width="210" height="25" class='menu' onclick="menuDisplay('outdoormenu2');">매체관리</td>
  </tr>
  <tr style="display:block;"  id="outdoormenu2">
    <td  width="210" >
		<table width="210" border="0" cellspacing="0" cellpadding="0">
<!-- 		  <tr>
			<td  width="210" height="24" <%'if request.cookies("menunum") = "5" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/hq/outdoor/medium_list.asp?menuNum=5'">매체현황</span></td>
		  </tr> -->
		  <tr>
			<td  width="210" height="24" <%if request.cookies("menunum") = "6" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/od/outdoor/job_list.asp?menuNum=6'">소재현황</span></td>
		  </tr>
		  <tr>
			<td  width="210" height="24" <%if request.cookies("menunum") = "11" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/od/outdoor/category_list.asp?menuNum=11'">매체분류</span></td>
		  </tr>
		 </table>
	</td>
  </tr>
<!--   <tr>
    <td  width="210" height="24" <%'if request.cookies("menunum") = "7" then Response.write "class='menuover'" Else response.write "class='menu'" End If%>><span onclick="location.href='/hq/outdoor/validation_tool_list.asp?menuNum=7'">효용성Tool</span></td>
  </tr> -->
  <tr>
    <td  width="210" height="24" <%if request.cookies("menunum") = "8" then Response.write "class='menuover'" Else response.write "class='menu'" End If%>><span onclick="location.href='/od/outdoor/monitor_list.asp?menuNum=8'">옥외모니터링</span></td>
  </tr>
  <tr>
    <td  width="210" height="24" <%if request.cookies("menunum") = "10" then Response.write "class='menuover'" Else response.write "class='menu'" End If%>><span onclick="window.open('/med/','pop_medium','')">사진관리</span></td>
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
