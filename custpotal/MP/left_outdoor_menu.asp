<%
	session("menunum") = request("menunum")
	If session("menunum") = "" Then session("menunum") =1
%>
<table width="210" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="/images/menu_outdoor_sub_title.gif" width="210" height="75"></td>
  </tr>
  <tr>
    <td  width="210" height="25" class='menu' onclick="menuDisplay('outdoormenu1');">옥외광고현황</td>
  </tr>
  <tr>
	<td  width="210" height="25" <%if session("menunum") = "1" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/mp/outdoor/list_contact.asp?menuNum=1';insertLogmst_src('옥외광고 집행현황')">옥외광고 집행현황</span></td>
 </tr>
  <tr>
	<td  width="210" height="24"<%if session("menunum") = "2" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/mp/outdoor/list_brand.asp?menuNum=2';insertLogmst_src('브랜드별 집행현황')">브랜드별 집행현황</span></td>
  </tr>
  <tr>
	<td  width="210" height="24" <%if session("menunum") = "4" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/mp/outdoor/list_finishdate.asp?menuNum=4';insertLogmst_src('종료일별 집행현황')">종료일별 집행현황</span></td>
  </tr>
<!--  <tr>
	<td  width="210" height="24" <%if session("menunum") = "9" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/mp/outdoor/list_transaction.asp?menuNum=9'">광고비용 집행현황</span></td>
  </tr>-->
  <tr>
	<td  width="210" height="24" <%if session("menunum") = "15" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/mp/outdoor/list_classsum.asp?menuNum=15';insertLogmst_src('매체별 집행현황')">매체별 집행현황</span></td>
  </tr>
  <tr>
	<td  width="210" height="24" <%if session("menunum") = "16" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/mp/outdoor/list_custsum.asp?menuNum=16';insertLogmst_src('월별 집행현황')">월별 집행현황</span></td>
  </tr>
  <tr>
	<td  width="210" height="24" <%if session("menunum") = "19" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/mp/outdoor/list_validation.asp?menuNum=19';insertLogmst_src('효용성평가 현황')">효용성평가 현황</span></td>
  </tr>

  <tr>
    <td  width="210" height="25" class='menu' onclick="menuDisplay('outdoormenu2');">옥외광고 모니터링</td>
  </tr>
  <tr>
	<td  width="210" height="24" <%if session("menunum") = "8" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/mp/outdoor/list_monitoring.asp?menuNum=8';insertLogmst_src('모니터링 보고현황')">모니터링 보고현황</span></td>
  </tr>
  <tr>
    <td  width="210" height="25" class='menu' onclick="menuDisplay('outdoormenu2');">옥외광고 관리보고</td>
  </tr>
  <tr>
	<td  width="210" height="24" <%if session("menunum") = "10" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/mp/outdoor/list_report.asp?menuNum=10';insertLogmst_src('관리보고')">관리보고</span></td>
  </tr>
  <tr>
    <td><img src="/images/menu_sub_bottom.gif" width="210" height="30"></td>
  </tr>
</table>
<div id="logpage"></div>
<script language="JavaScript">
<!--
	var menuFlag = true ;
	function menuDisplay(id) {

	}


	function insertLogmst_src(progname) {
		var params = "progname="+progname;
		sendRequest("/inc/setlog.asp", params, _insertLogmst_src, "GET");
	}

	function _insertLogmst_src() {
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				var custcode = document.getElementById("logpage");

				custcode.innerHTML = xmlreq.responseText ;
			}
		}
	}
//-->
</script>
