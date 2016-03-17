<!--#include virtual="/inc/func.asp" -->
<%
	session("menunum") = request("menunum")
	If session("menunum") = "" Then session("menunum") =1
	Dim pagename : pagename = request.cookies("pagename")
	response.write pagename

%>
<input type="hidden" name="txtpagename" value=<%=strpagename%>>
<table width="210" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="/images/menu_outdoor_sub_title.gif" width="210" height="75"></td>
  </tr>
  <tr>
    <td  width="210" height="25" class='menu' onclick="menuDisplay('outdoormenu1');">옥외광고현황</td>
  </tr>
  <tr>
	<td  width="210" height="25" <%if session("menunum") = "1" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/hq/outdoor/list_contact.asp?menuNum=1';insertLogmst_src('list_contact.asp')">옥외광고 집행현황</span></td>
 </tr>
 <!--tr>
	<td  width="210" height="25" <%if session("menunum") = "17" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/hq/outdoor/list_contact_month.asp?menuNum=17'">옥외광고 월별조회</span></td>
 </tr-->
  <tr>
	<td  width="210" height="24"<%if session("menunum") = "2" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/hq/outdoor/list_brand.asp?menuNum=2';insertLogmst_src('list_brand.asp')">브랜드별 집행현황</span></td>
  </tr>
  <tr>
	<td  width="210" height="24" <%if session("menunum") = "4" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/hq/outdoor/list_finishdate.asp?menuNum=4';insertLogmst_src('list_finishdate.asp')">종료일별 집행현황</span></td>
  </tr>
  <tr>
	<td  width="210" height="24" <%if session("menunum") = "9" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/hq/outdoor/list_transaction.asp?menuNum=9';insertLogmst_src('list_transaction.asp')">광고비용 집행현황</span></td>
  </tr>
  <tr>
	<td  width="210" height="24" <%if session("menunum") = "15" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/hq/outdoor/list_classsum.asp?menuNum=15';insertLogmst_src('list_classsum.asp')">매체별 집행현황</span></td>
  </tr>
  <tr>
	<td  width="210" height="24" <%if session("menunum") = "16" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/hq/outdoor/list_custsum.asp?menuNum=16';insertLogmst_src('list_custsum.asp')">월별 집행현황</span></td>
  </tr>
  <tr>
	<td  width="210" height="24" <%if session("menunum") = "19" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/hq/outdoor/list_validation.asp?menuNum=19';insertLogmst_src('list_validation.asp')">효용성평가 현황</span></td>
  </tr>
  <tr>
    <td  width="210" height="25" class='menu' onclick="menuDisplay('outdoormenu2');">매체관리</td>
  </tr>
  <tr>
	<td  width="210" height="24" <%if session("menunum") = "6" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/hq/outdoor/list_subseq.asp?menuNum=6';insertLogmst_src('list_subseq.asp')">소재관리</span></td>
  </tr>
  <tr>
	<td  width="210" height="24" <%if session("menunum") = "11" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/hq/outdoor/list_mediumclass.asp?menuNum=11';insertLogmst_src('list_mediumclass.asp')">분류관리 </span></td>
  </tr>
  <tr>
	<td  width="210" height="24" <%if session("menunum") = "13" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/hq/outdoor/list_quality.asp?menuNum=13';insertLogmst_src('list_quality.asp')">재질관리</span></td>
  </tr>
  <tr>
    <td  width="210" height="25" class='menu' onclick="menuDisplay('outdoormenu2');">옥외광고 모니터링</td>
  </tr>
  <tr>
	<td  width="210" height="24" <%if session("menunum") = "8" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/hq/outdoor/list_monitoring.asp?menuNum=8';insertLogmst_src('list_monitoring.asp')">모니터링 보고현황</span></td>
  </tr>
  <tr>
    <td  width="210" height="25" class='menu' onclick="menuDisplay('outdoormenu2');">옥외광고 관리보고</td>
  </tr>
  <tr>
	<td  width="210" height="24" <%if session("menunum") = "10" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/hq/outdoor/list_report.asp?menuNum=10';insertLogmst_src('list_report.asp')">관리보고</span></td>
  </tr>
  <tr>
	<td  width="210" height="24" <%if session("menunum") = "14" then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%>><span onclick="location.href='/hq/outdoor/list_employee.asp?menuNum=14';insertLogmst_src('list_employee.asp')">매체사별 계정관리</span></td>
  </tr>
  <tr>
    <td><img src="/images/menu_sub_bottom.gif" width="210" height="30"></td>
  </tr>
</table>
<div id="logpage" ></div>
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

