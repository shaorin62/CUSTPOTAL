<table width="210" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="/images/menu_trans_sub_title.gif" width="210" height="75"></td>
  </tr>
  <%
	Dim objrs, sql, menu
	sql = "select custcode, custname from dbo.sc_cust_temp where  custcode= '" & request.cookies("custcode") &"' "
	Call get_recordset(objrs, sql)

	Dim menu_custcode, menu_custname
	If Not objrs.eof Then
		Set menu_custcode = objrs("custcode")
		Set menu_custname = objrs("custname")
	End If

	do until objrs.eof
%>
  <tr>
    <td  width="210" height="25"  class="deps" background="/images/menu_bg_over.gif" style=" cursor:hand;" onclick="checkDisplay('<%=menu_custcode%>');"><%=menu_custname%></td>
  </tr>
  <tr id="<%=menu_custcode%>" style="display:block" >
		<td>
			<table border="0" cellspacing="0" cellpadding="0">
				<tr>
				<td  width="210" height="22" class="deps_1" id="<%=menu_custcode%>01" background="/images/menu_bg.gif" onclick="get_data('trans_01.asp','<%=menu_custcode%>','월별 매체 집행내역', '<%=menu_custcode%>01')" style="cursor:hand">월별 매체 집행내역</td>
				</tr>
				<tr>
				<td  width="210" height="22" class="deps_1"  id="<%=menu_custcode%>02" background="/images/menu_bg.gif" onclick="get_data('trans_02.asp','<%=menu_custcode%>','ATL 브랜드/소재별 광고비','<%=menu_custcode%>02')" style="cursor:hand">ATL 브랜드/소재별 광고비</td>
				</tr>
				<tr>
				<td  width="210" height="22" class="deps_1"  id="<%=menu_custcode%>03" background="/images/menu_bg.gif" onclick="get_data('trans_03.asp','<%=menu_custcode%>','AOR 공중파 광고정산','<%=menu_custcode%>03')"  style="cursor:hand">AOR 공중파 광고정산</td>
				</tr>
				<tr>
				<td  width="210" height="22"  class="deps_1"  background="/images/menu_bg_over.gif">공중파</td>
				</tr>
				<tr>
				<td  width="210" height="19"  class="deps_3"   id="<%=menu_custcode%>04" background="/images/menu_bg.gif"  onclick="get_data('trans_04.asp','<%=menu_custcode%>','부문별 실집행 광고비','<%=menu_custcode%>04')"  style="cursor:hand">실집행 광고비</td>
				</tr>
				<tr>
				<td  width="210" height="22"  class="deps_1"  background="/images/menu_bg_over.gif">케이블</td>
				</tr>
				<tr>
				<td  width="210" height="19"  class="deps_3"   id="<%=menu_custcode%>05" background="/images/menu_bg.gif" onclick="get_data('trans_05.asp','<%=menu_custcode%>','월별/매체별 실집행 광고비','<%=menu_custcode%>05')"  style="cursor:hand">월별/매체별 실집행 광고비</td>
				</tr>
				<tr>
				<td  width="210" height="22"  class="deps_1"  background="/images/menu_bg_over.gif">인쇄</td>
				</tr>
				<tr>
				<td  width="210" height="19"  class="deps_3"  id="<%=menu_custcode%>06"  background="/images/menu_bg.gif" onclick="get_data('trans_06.asp','<%=menu_custcode%>','월별 광고비 요약','<%=menu_custcode%>06')"  style="cursor:hand">월별 광고비 요약</td>
				</tr>
				<tr>
				<td  width="210" height="19"  class="deps_3"   id="<%=menu_custcode%>07" background="/images/menu_bg.gif" onclick="get_data('trans_07.asp','<%=menu_custcode%>','매체별 집행내역','<%=menu_custcode%>07')"   style="cursor:hand">매체별 집행내역</td>
				</tr>
				<tr>
				<td  width="210" height="19"  class="deps_3"  id="<%=menu_custcode%>08"  background="/images/menu_bg.gif" onclick="get_data('trans_08.asp','<%=menu_custcode%>','세부 집행내역','<%=menu_custcode%>08')" style="cursor:hand">세부 집행내역</td>
				</tr>
				<tr>
				<td  width="210" height="22"  class="deps_1"  background="/images/menu_bg_over.gif">인터넷</td>
				</tr>
				<tr>
				<td  width="210" height="19"  class="deps_3"  id="<%=menu_custcode%>09"  background="/images/menu_bg.gif"  onclick="get_data('trans_09.asp','<%=menu_custcode%>','월별 큐시트','<%=menu_custcode%>09')" style="cursor:hand">월별 큐시트</td>
				</tr>
				<tr>
				<td  width="210" height="19"  class="deps_3"  id="<%=menu_custcode%>10"  background="/images/menu_bg.gif"  onclick="get_data('trans_10.asp','<%=menu_custcode%>','매체별 광고비','<%=menu_custcode%>10')" style="cursor:hand">매체별 광고비</td>
				</tr>
				<tr>
				<td  width="210" height="19"  class="deps_3"   id="<%=menu_custcode%>11"  background="/images/menu_bg.gif" onclick="get_data('trans_11.asp','<%=menu_custcode%>','광고주/CIC별 매체비','<%=menu_custcode%>11')" style="cursor:hand">광고주/CIC별 매체비</td>
				</tr>
<!-- 				<tr>
				<td  width="210" height="22"  class="deps_1"  background="/images/menu_bg_over.gif">옥외</td>
				</tr>
				<tr>
				<td  width="210" height="19"  class="deps_3"   id="<%'=menu_custcode%>12"  background="/images/menu_bg.gif" onclick="get_data('trans_12.asp','<%'=menu_custcode%>','옥외광고 현황','<%'=menu_custcode%>12')" style="cursor:hand">옥외광고 현황</td>
				</tr> -->
			</table>
		</td>
  </tr>
  <%
	objrs.movenext
	Loop
	objrs.close
	set objrs = nothing
  %>
  <tr>
	<td  width="210" height="5" class="deps"  background="/images/menu_bg.gif"> </td>
  </tr>
  <tr>
    <td><img src="/images/menu_sub_bottom.gif" width="210" height="30"></td>
  </tr>
</table>
<script language="JavaScript">
<!--

	function checkDisplay(code) {
		return true ;
	}
	function get_data(url, code, str, p) {
		var subtitle = document.getElementById("subtitle") ;
		var navigate = document.getElementById("navigator");

		var cyear = document.getElementById("cyear");
		var cyear2 = document.getElementById("cyear2");
		var cmonth = document.getElementById("cmonth");
		var cmonth2 = document.getElementById("cmonth2");
		var custcode2 = document.getElementById("tcustcode2");
		var e7 = document.getElementById("e7");
		var subname = document.getElementById("subname");
		var actionurl = document.getElementById("actionurl");

		var td = document.getElementsByTagName("td")
		for (var i = 0 ; i < td.length ; i++) {
			if (!td[i].getAttribute("id")) continue;
			if (td[i].className == "depsover_3") td[i].className = "deps_3" ;
			if (td[i].className == "depsover_1") td[i].className = "deps_1" ;
			if (td[i].className == "deps_3" || td[i].className == "deps_1" ) {
				if (td[i].getAttribute("id") == p) {
					if (td[i].className == "deps_3") td[i].className = "depsover_3" ;
					if (td[i].className == "deps_1") td[i].className = "depsover_1" ;
				}
			}
		}
		switch (url)
		{
		case "trans_01.asp":
			subtitle.innerText = str;
			navigate.innerText = str;
			cyear.style.display = "";
			cyear2.style.display = "";
			e7.style.display = "";
			cmonth.style.display = "";
			cmonth2.style.display = "";
			actionurl.value = url ;
			break ;
		case "trans_02.asp":
			subtitle.innerText = str;
			navigate.innerText = str;
			cyear.style.display = "";
			cyear2.style.display = "";
			e7.style.display = "";
			cmonth.style.display = "";
			cmonth2.style.display = "";
			actionurl.value = url ;
			break ;
		case "trans_03.asp":
			subtitle.innerText = str;
			navigate.innerText = str;
			cyear.style.display = "";
			cyear2.style.display = "none";
			e7.style.display = "none";
			cmonth.style.display = "";
			cmonth2.style.display = "none";
			actionurl.value = url ;
			break ;
		case "trans_04.asp":
			subtitle.innerText = str;
			navigate.innerText = "공중파 > " + str;
			cyear.style.display = "";
			cyear2.style.display = "none";
			e7.style.display = "none";
			cmonth.style.display = "none";
			cmonth2.style.display = "none";
			actionurl.value = url ;
			break ;
		case "trans_05.asp":
			subtitle.innerText = str;
			navigate.innerText = "케이블 > " + str;
			cyear.style.display = "";
			cyear2.style.display = "none";
			e7.style.display = "none";
			cmonth.style.display = "none";
			cmonth2.style.display = "none";
			actionurl.value = url ;
			break ;
		case "trans_06.asp":
			subtitle.innerText = str;
			navigate.innerText = "인쇄 > " + str;
			cyear.style.display = "";
			e7.style.display = "none";
			cyear2.style.display = "none";
			cmonth.style.display = "none";
			cmonth2.style.display = "none";
			actionurl.value = url ;
			break ;
		case "trans_07.asp":
			subtitle.innerText = str;
			navigate.innerText = "인쇄 > " + str;
			cyear.style.display = "";
			e7.style.display = "";
			cyear2.style.display = "";
			cmonth.style.display = "";
			cmonth2.style.display = "";
			actionurl.value = url ;
			break ;
		case "trans_08.asp":
			subtitle.innerText = str;
			navigate.innerText = "인쇄 > " + str;
			cyear.style.display = "";
			cyear2.style.display = "none";
			e7.style.display = "none";
			cmonth.style.display = "";
			cmonth2.style.display = "none";
			actionurl.value = url ;
			break ;
		case "trans_09.asp":
			subtitle.innerText = str;
			navigate.innerText = "인터넷 > " + str;
			cyear.style.display = "";
			cyear2.style.display = "none";
			e7.style.display = "none";
			cmonth.style.display = "";
			cmonth2.style.display = "none";
			actionurl.value = url ;
			break ;
		case "trans_10.asp":
			subtitle.innerText = str;
			navigate.innerText = "인터넷 > " + str;
			cyear.style.display = "";
			cyear2.style.display = "none";
			e7.style.display = "none";
			cmonth.style.display = "";
			cmonth2.style.display = "none";
			actionurl.value = url ;
			break ;
		case "trans_11.asp":
			subtitle.innerText = str;
			navigate.innerText = "인터넷 > " + str;
			cyear.style.display = "";
			cyear2.style.display = "none";
			e7.style.display = "none";
			cmonth.style.display = "none";
			cmonth2.style.display = "none";
			actionurl.value = url ;
			break ;
		case "trans_12.asp":
			subtitle.innerText = str;
			navigate.innerText = "옥외 > " + str;
			cyear.style.display = "";
			cyear2.style.display = "none";
			e7.style.display = "none";
			cmonth.style.display = "";
			cmonth2.style.display = "none";
			actionurl.value = url ;
			break ;
		}
		subname.innerText = ""
		var frm = document.forms[0];
		frm.tcustcode.value = code ;
		scriptFrame.location.href ="/cust/trans/"+url+"?tcustcode="+code+"&tcustcode2=&cyear=2005&cmonth=1&cyear2=2000&cmonth=1";
	}
		function changeNavigator(url, code) {
		scriptFrame.location.href ="/cust/trans/"+url+"?tcustscode="+code;
	}
	window.onload = function () {
		var td = document.getElementsByTagName("td")
		for (var i = 0 ; i < td.length ; i++) {
			if (!td[i].getAttribute("id")) continue;
			if (td[i].className == "deps_1") {
				td[i].className = "depsover_1" ;
				document.getElementById("cyear").style.display = "";
				document.getElementById("cyear2").style.display = "";
				document.getElementById("e7").style.display = "";
				document.getElementById("cmonth").style.display = "";
				document.getElementById("cmonth2").style.display = "";
				document.getElementById("actionurl").value = "trans_01.asp";
				return false;
			}
		}
	}
//-->
</script>