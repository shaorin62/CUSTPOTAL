<table width="210" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="/images/menu_trans_sub_title.gif" width="210" height="75"></td>
  </tr>
  <%
	Dim objrs, sql, menu
	sql = "select highcustcode, custname from dbo.sc_cust_hdr where  medflag = 'A' order by custname"
	Call get_recordset(objrs, sql)

	Dim menu_custcode, menu_custname
	If Not objrs.eof Then
		Set menu_custcode = objrs("highcustcode")
		Set menu_custname = objrs("custname")
	End If
  %>
  <tr>
	<td  width="210" height="19" class="deps" id="pub_01" background="/images/menu_bg.gif" onclick="getData();" style="cursor:hand">���� ��ü�� �����</td>
  </tr>
  <tr>
	<td  width="210" height="19" class="deps"  id="pub_02"  background="/images/menu_bg.gif" onclick="changeNavigator('public_02.asp',''); "  style="cursor:hand">���� ���� ��ü�� �����</td>
  </tr>
  <tr>
	<td  width="210" height="19" class="deps"  id="pub_03" background="/images/menu_bg.gif" onclick="changeNavigator('public_03.asp','')" style="cursor:hand;">CATV/ New Media ����</td>
  </tr>
<%
	do until objrs.eof
%>
  <tr>
    <td  width="210" height="25"  class="deps" background="/images/menu_dot_bg_over.gif" style="padding-top:5px; cursor:hand;" onclick="checkDisplay('<%=menu_custcode%>');"><%=menu_custname%></td>
  </tr>
  <tr id="<%=menu_custcode%>" style="display:none" >
		<td>
			<table border="0" cellspacing="0" cellpadding="0">
				<tr>
				<td  width="210" height="22" class="deps_1" id="<%=menu_custcode%>01" background="/images/menu_bg.gif" onclick="get_data('trans_01.asp','<%=menu_custcode%>','���� ��ü ���೻��', '<%=menu_custcode%>01')" style="cursor:hand">���� ��ü ���೻��</td>
				</tr>
				<tr>
				<td  width="210" height="22" class="deps_1"  id="<%=menu_custcode%>02" background="/images/menu_bg.gif" onclick="get_data('trans_02.asp','<%=menu_custcode%>','ATL �귣��/���纰 �����','<%=menu_custcode%>02')" style="cursor:hand">ATL �귣��/���纰 �����</td>
				</tr>
				<tr>
				<td  width="210" height="22" class="deps_1"  id="<%=menu_custcode%>03" background="/images/menu_bg.gif" onclick="get_data('trans_03.asp','<%=menu_custcode%>','AOR ������ ��������','<%=menu_custcode%>03')"  style="cursor:hand">AOR ������ ��������</td>
				</tr>
				<tr>
				<td  width="210" height="22"  class="deps_1"  background="/images/menu_bg_over.gif">������</td>
				</tr>
				<tr>
				<td  width="210" height="19"  class="deps_3"   id="<%=menu_custcode%>04" background="/images/menu_bg.gif"  onclick="get_data('trans_04.asp','<%=menu_custcode%>','�ι��� ������ �����','<%=menu_custcode%>04')"  style="cursor:hand">������ �����</td>
				</tr>
				<tr>
				<td  width="210" height="22"  class="deps_1"  background="/images/menu_bg_over.gif">���̺�</td>
				</tr>
				<tr>
				<td  width="210" height="19"  class="deps_3"   id="<%=menu_custcode%>05" background="/images/menu_bg.gif" onclick="get_data('trans_05.asp','<%=menu_custcode%>','����/��ü�� ������ �����','<%=menu_custcode%>05')"  style="cursor:hand">����/��ü�� ������ �����</td>
				</tr>
				<% if menu_custcode <> "A00005"  or  request.cookies("class") <> "N"  then  %>
				<tr>
				<td  width="210" height="22"  class="deps_1"  background="/images/menu_bg_over.gif">�μ�</td>
				</tr>
				<tr>
				<td  width="210" height="19"  class="deps_3"  id="<%=menu_custcode%>06"  background="/images/menu_bg.gif" onclick="get_data('trans_06.asp','<%=menu_custcode%>','���� ����� ���','<%=menu_custcode%>06')"  style="cursor:hand">���� ����� ���</td>
				</tr>
				<tr>
				<td  width="210" height="19"  class="deps_3"   id="<%=menu_custcode%>07" background="/images/menu_bg.gif" onclick="get_data('trans_07.asp','<%=menu_custcode%>','��ü�� ���೻��','<%=menu_custcode%>07')"   style="cursor:hand">��ü�� ���೻��</td>
				</tr>
				<tr>
				<td  width="210" height="19"  class="deps_3"  id="<%=menu_custcode%>08"  background="/images/menu_bg.gif" onclick="get_data('trans_08.asp','<%=menu_custcode%>','���� ���೻��','<%=menu_custcode%>08')" style="cursor:hand">���� ���೻��</td>
				</tr>
				<% end if%>
				<tr>
				<td  width="210" height="22"  class="deps_1"  background="/images/menu_bg_over.gif">���ͳ�</td>
				</tr>
				<tr>
				<td  width="210" height="19"  class="deps_3"  id="<%=menu_custcode%>09"  background="/images/menu_bg.gif"  onclick="get_data('trans_09.asp','<%=menu_custcode%>','���� ť��Ʈ','<%=menu_custcode%>09')" style="cursor:hand">���� ť��Ʈ</td>
				</tr>
				<tr>
				<td  width="210" height="19"  class="deps_3"  id="<%=menu_custcode%>10"  background="/images/menu_bg.gif"  onclick="get_data('trans_10.asp','<%=menu_custcode%>','��ü�� �����','<%=menu_custcode%>10')"  style="cursor:hand">��ü�� �����</td>
				</tr>
				<tr>
				<td  width="210" height="19"  class="deps_3"   id="<%=menu_custcode%>11"  background="/images/menu_bg.gif" onclick="get_data('trans_11.asp','<%=menu_custcode%>','������/CIC�� ��ü��','<%=menu_custcode%>11')"   style="cursor:hand">������/CIC�� ��ü��</td>
				</tr>
<!-- 				<tr>
				<td  width="210" height="22"  class="deps_1"  background="/images/menu_bg_over.gif">����</td>
				</tr>
				<tr>
				<td  width="210" height="19"  class="deps_3"   id="<%'=menu_custcode%>12"  background="/images/menu_bg.gif" onclick="get_data('trans_12.asp','<%'=menu_custcode%>','���ܱ��� ��Ȳ','<%'=menu_custcode%>12')" style="cursor:hand">���ܱ��� ��Ȳ</td>
				</tr> -->
			</table>
		</td>
  </tr>
  <%
	objrs.movenext
	Loop
	objrs.movefirst
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
		<% do until objrs.eof %>
		document.getElementById("<%=menu_custcode%>").style.display = "none";
		<%
				objrs.movenext
				loop
		%>

		document.getElementById(code).style.display = "block";
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
		var subname = document.getElementById("subname")
		var actionurl = document.getElementById("actionurl");


		var pub_01 = document.getElementById("Pub_01");
		var pub_02 = document.getElementById("pub_02");
		var pub_03 = document.getElementById("pub_03");
		pub_01.className = "deps" ;
		pub_02.className = "deps" ;
		pub_03.className = "deps" ;

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
			navigate.innerText = "������ > " + str;
			cyear.style.display = "";
			cyear2.style.display = "none";
			e7.style.display = "none";
			cmonth.style.display = "none";
			cmonth2.style.display = "none";
			actionurl.value = url ;
			break ;
		case "trans_05.asp":
			subtitle.innerText = str;
			navigate.innerText = "���̺� > " + str;
			cyear.style.display = "";
			cyear2.style.display = "none";
			e7.style.display = "none";
			cmonth.style.display = "none";
			cmonth2.style.display = "none";
			actionurl.value = url ;
			break ;
		case "trans_06.asp":
			subtitle.innerText = str;
			navigate.innerText = "�μ� > " + str;
			cyear.style.display = "";
			e7.style.display = "none";
			cyear2.style.display = "none";
			cmonth.style.display = "none";
			cmonth2.style.display = "none";
			actionurl.value = url ;
			break ;
		case "trans_07.asp":
			subtitle.innerText = str;
			navigate.innerText = "�μ� > " + str;
			cyear.style.display = "";
			e7.style.display = "";
			cyear2.style.display = "";
			cmonth.style.display = "";
			cmonth2.style.display = "";
			actionurl.value = url ;
			break ;
		case "trans_08.asp":
			subtitle.innerText = str;
			navigate.innerText = "�μ� > " + str;
			cyear.style.display = "";
			cyear2.style.display = "none";
			e7.style.display = "none";
			cmonth.style.display = "";
			cmonth2.style.display = "none";
			actionurl.value = url ;
			break ;
		case "trans_09.asp":
			subtitle.innerText = str;
			navigate.innerText = "���ͳ� > " + str;
			cyear.style.display = "";
			cyear2.style.display = "none";
			e7.style.display = "none";
			cmonth.style.display = "";
			cmonth2.style.display = "none";
			actionurl.value = url ;
			break ;
		case "trans_10.asp":
			subtitle.innerText = str;
			navigate.innerText = "���ͳ� > " + str;
			cyear.style.display = "";
			cyear2.style.display = "none";
			e7.style.display = "none";
			cmonth.style.display = "";
			cmonth2.style.display = "none";
			actionurl.value = url ;
			break ;
		case "trans_11.asp":
			subtitle.innerText = str;
			navigate.innerText = "���ͳ� > " + str;
			cyear.style.display = "";
			cyear2.style.display = "none";
			e7.style.display = "none";
			cmonth.style.display = "none";
			cmonth2.style.display = "none";
			actionurl.value = url ;
			break ;
		case "trans_12.asp":
			subtitle.innerText = str;
			navigate.innerText = "���� > " + str;
			cyear.style.display = "";
			cyear2.style.display = "none";
			e7.style.display = "none";
			cmonth.style.display = "";
			cmonth2.style.display = "none";
			actionurl.value = url ;
			break ;
		}
		subname.innerText = ""
//		var frm = document.forms[0];
//		frm.tcustcode.value = code ;
//		scriptFrame.location.href =url+"?tcustcode="+code+"&tcustcode2=&cyear=2005&cmonth=01&cyear2=2005&cmonth2=01";
	}
		function changeNavigator(url, code) {
		var frm = document.forms[0];
		var pub_01 = document.getElementById("Pub_01");
		var pub_02 = document.getElementById("pub_02");
		var pub_03 = document.getElementById("pub_03");

		var subtitle = document.getElementById("subtitle") ;
		var navi = document.getElementById("navigator");
		var cyear = document.getElementById("cyear");
		var cyear2 = document.getElementById("cyear2");
		var cmonth = document.getElementById("cmonth");
		var cmonth2 = document.getElementById("cmonth2");
		var actionurl = document.getElementById("actionurl");
		var subname = document.getElementById("subname") ;
		subname.innerText = "";
		pub_01.className = "deps" ;
		pub_02.className = "deps" ;
		pub_03.className = "deps" ;
		var td = document.getElementsByTagName("td")
		for (var i = 0 ; i < td.length ; i++) {
			if (!td[i].getAttribute("id")) continue;
			if (td[i].className == "depsover_3") td[i].className = "deps_3" ;
			if (td[i].className == "depsover_1") td[i].className = "deps_1" ;
		}
		switch (url) {
			case "public_01.asp":
				subtitle.innerText = "���� ��ü�� �����";
				navi.innerText = "���� ��ü�� �����";
				cyear.style.display = "";
				cyear2.style.display = "";
				cmonth.style.display = "";
				cmonth2.style.display = "";
				actionurl.value = url ;
				if (pub_01.className == "deps")	pub_01.className = "depsover";
			break;
			case "public_02.asp":
				subtitle.innerText = "���� ���� ��ü�� �����";
				navi.innerText = "���� ���� ��ü�� �����";
				cyear.style.display = "";
				cyear2.style.display = "none";
				cmonth.style.display = "";
				cmonth2.style.display = "";
				actionurl.value = url ;
				if (pub_02.className == "deps")	pub_02.className = "depsover";
			break;
			case "public_03.asp":
				subtitle.innerText = "CATV/ New Media ����";
				navi.innerText = "CATV/ New Media ����";
				cyear.style.display = "";
				cyear2.style.display = "";
				cmonth.style.display = "";
				cmonth2.style.display = "";
				actionurl.value = url ;
				if (pub_03.className == "deps")	pub_03.className = "depsover";
			break;
		}
		frm.tcustcode.value = code ;
		// Iframe ���� ������ ���� ��Ų��. �ش� �������ڵ� �ѱ��
		scriptFrame.location.href =url+"?tcustscode="+code+"&cyear=2005&cmonth=01&cyear2=2005&cmonth2=01";
	}




	// -------------- New Javascript

	function getData() {
		alert(srcElement.id);
	}
//-->
</script>