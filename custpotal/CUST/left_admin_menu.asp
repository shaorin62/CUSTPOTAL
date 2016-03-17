<table width="210" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="210" height="75"><img src="/images/menu_admin_sub_title.gif" width="210" height="75"></td>
  </tr>

  <tr>
    <td width="210" height="20" class="deps"   background="/images/menu_bg.gif">皋春包府</td>
  </tr>
		<%
			Dim objrs2
			sql = "select highcustcode, custname from dbo.sc_cust_hdr where highcustcode = '" & request.cookies("highcustcode") & "'   and use_flag = 1 and medflag ='A' order by custname "

			Call get_recordset(objrs2, sql)
			Do Until objrs2.eof
		%>
		  <tr>
			<td  width="210" height="22" class='deps_1' background="/images/menu_bg_over.gif"  id="<%=objrs2("highcustcode")%>_cust" style="width:150px;cursor:hand" onclick="checkDisplay('<%=objrs2("highcustcode")%>'); go_page('menu.asp?custcode=<%=objrs2("highcustcode")%>','<%=objrs2("highcustcode")%>_cust','皋春包府 > <%=objrs2("custname")%>','<%=objrs2("custname")%>','highcustcode'); left_menu_getdata('menu.asp','<%=objrs2("highcustcode")%>','highcustcode'); "><%=objrs2("custname")%></span></td>
		  </tr>
		  <tr id="<%=objrs2("highcustcode")%>" style="display:block;">
			<td height="0" width="210" background="/images/menu_bg.gif">
			<table  border="0" cellspacing="0" cellpadding="0">
		  <%
					dim objrs3
					sql = "select custcode, custname from dbo.sc_cust_dtl where highcustcode ='" & objrs2("highcustcode") & "' and use_flag = 1 and gbnflag = '0'  order by custname"

					call get_recordset(objrs3, sql)
					do until objrs3.eof
			%>
		  <tr>
			<td  width="210" height="22" class='deps_2' background="/images/menu_bg.gif" id="<%=objrs3("custcode")%>" onclick="go_page('menu.asp?custcode=<%=objrs3("custcode")%>','<%=objrs3("custcode")%>','皋春包府 > <%=objrs3("custname")%>','<%=objrs3("custname")%>','timcode'); left_menu_getdata('menu.asp','<%=objrs3("custcode")%>','timcode');" style="width:150px;cursor:hand"><%=objrs3("custname")%></span></td>
		  </tr>
		  <%
					objrs3.movenext
					Loop
					objrs3.close
		%>
				</table>
			</td>
		  </tr>
		<%
				objrs2.movenext
				Loop
				objrs2.movefirst
				set objrs3 = nothing
		  %>
  </tr>
	<tr>
		<td  width="210" height="24" class="submenu">&nbsp;</td>
	</tr>
  <tr>
    <td width="210" height="30"><img src="/images/menu_sub_bottom.gif" width="210" height="30"></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<script language="JavaScript">
<!--
	function go_page(url, id, str, str2,flag) {

		var td = document.getElementsByTagName("td")
		for (var i = 0 ; i < td.length ; i++) {
			if (!td[i].getAttribute("id")) continue;
			if (td[i].className == "depsover") td[i].className = "deps" ;
			if (td[i].className == "depsover_1") td[i].className = "deps_1" ;
			if (td[i].className == "depsover_2") td[i].className = "deps_2" ;
			if (td[i].className == "deps" || td[i].className == "deps_1" || td[i].className == "deps_2") {
				if (td[i].getAttribute("id") == id) {
					if (td[i].className == "deps") td[i].className = "depsover" ;
					if (td[i].className == "deps_1") td[i].className = "depsover_1" ;
					if (td[i].className == "deps_2") td[i].className = "depsover_2" ;
				}
			}
		}


		var subtitle = document.getElementById("subtitle");
		var navi = document.getElementById("navi");
		subtitle.firstChild.nodeValue = str2 ;
		navi.firstChild.nodeValue = "包府葛靛 > " + str ;
		var img = document.getElementById("btnReg");


		img.setAttribute("src","/images/btn_menu_reg.gif");
		img.setAttribute("class", "menu");
		document.getElementById("searchsection").style.display = "none";

		document.forms[0].tcustcode.value = id.replace("_cust", ""); ;

	}

	function checkDisplay(code) {

		<% do until objrs2.eof %>
		document.getElementById("<%=objrs2("highcustcode")%>").style.display = "none";
		<%
				objrs2.movenext
				loop
		%>

		document.getElementById(code).style.display = "block";

	}
//	window.onload = function () {
//		var td = document.getElementsByTagName("td")
//		for (var i = 0 ; i < td.length ; i++) {
//			if (!td[i].getAttribute("id")) continue;
//			if (td[i].className == "deps_1") {
//				td[i].className = "depsover_1" ;
//				document.getElementById("scriptFrame").setAttribute("src", "menu.asp");
//				return false;
//			}
//		}
//	}
//-->
</script>