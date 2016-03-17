
<%
	dim objrs, sql
%>

<table width="210" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="210" height="75"><img src="/images/menu_admin_sub_title.gif" width="210" height="75"></td>
  </tr>
  <tr>
	<td  width="210" height="19"   class="menuheader" style="cursor:hand" onclick="left_getdata('account.asp'); return false;">계정관리</td>
  </tr>
  <tr>
    <td width="210" height="20" class="deps"   background="/images/menu_bg.gif">메뉴관리</td>
  </tr>
		  <tr>
				<td  width="210" height="19"   class="deps_1" background="/images/menu_bg_over.gif" style="cursor:hand" onclick="left_getdata('sk_admin_menu.asp'); return false;">공통메뉴</td>
		  </tr>
		<%
			Dim objrs2
			sql = " select highcustcode custcode, custname from MD_CLIENTCODE_LIST_V order by custname  "

			Call get_recordset(objrs2, sql)

			Do Until objrs2.eof

		%>
		  <tr>
			<td  width="210" height="22" class='deps_1' background="/images/menu_bg_over.gif"  id="<%=objrs2("custcode")%>_cust" style="width:150px;cursor:hand" onclick="checkDisplay('<%=objrs2("custcode")%>');   return false;" ><%=objrs2("custname")%></td>
			<!--left_menu_getdata('sk_admin_menu.asp','<%=objrs2("custcode")%>' ,'HIGHCUSTCODE'); -->
		  </tr>
		  <tr id="<%=objrs2("custcode")%>" style="display:none;">
			<td height="0" width="210" background="/images/menu_bg.gif">
			<table  border="0" cellspacing="0" cellpadding="0">
		  <%

					dim objrs3
					sql = "select custcode, custname from dbo.sc_cust_dtl where isnull(gbnflag,'0') = '0'  and highcustcode ='" & objrs2("custcode") & "' and isnull(use_flag,0) = 1   and medflag = 'A' order by custname"

					call get_recordset(objrs3, sql)
					do until objrs3.eof
			%>
		  <tr>
			<td  width="210" height="22" class='deps_2' background="/images/menu_bg.gif" id="- " & <%=objrs3("custcode")%>" onclick="left_menu_getdata('sk_admin_menu.asp','<%=objrs3("custcode")%>','CUSTCODE');  return false; " style="width:150px;cursor:hand"><%=objrs3("custname")%></td>
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

	function checkDisplay(code) {
		<% do until objrs2.eof %>
		document.getElementById("<%=objrs2("custcode")%>").style.display = "none";
		<%
				objrs2.movenext
				loop
		%>

		document.getElementById(code).style.display = "block";
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

//	window.onload = function () {
//		var td = document.getElementsByTagName("td")
//		for (var i = 0 ; i < td.length ; i++) {
//			if (!td[i].getAttribute("id")) continue;
//			if (td[i].className == "deps") {
//				td[i].className = "depsover" ;
//				document.getElementById("scriptFrame").setAttribute("src", "account.asp");
//				return false;
//			}
//		}
//	}
//-->
</script>