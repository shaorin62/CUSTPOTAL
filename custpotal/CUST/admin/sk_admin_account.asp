<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<form target="scriptFrame">
<!--#include virtual="/cust/top.asp" -->
<input type="hidden" name="actionurl" value="account.asp">
<input type="hidden" name="tcustcode">
  <table id="Table_01" width="1240" height="600" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/cust/left_admin_menu.asp"--></td>
      <td height="65"><img src="/images/middle_navigater_admin.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td height="600" align="left" valign="top"><table width="1030" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="19">&nbsp;</td>
          </tr>
          <tr>
            <td height="17"><TABLE  width="100%">
            <TR>
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle" id="subtitle"> 계정관리 </span></TD>
				<TD width="50%" align="right"><span class="navigator" id="navi">관리모드 &gt; 계정관리 </span></TD>
            </TR>
            </TABLE>
			</td>
          </tr>
          <tr>
            <td height="15">&nbsp;</td>
          </tr>
          <tr>
            <td ><table width="1030" height="35" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
                  <td width="50%" align="left" background="/images/bg_search.gif"><span id="searchsection"><input type="text" name="txtsearchstring"> <img src="/images/btn_search.gif" width="39" height="20" align="top" class="styleLink" onClick="checkForSearch(document.forms[0].txtsearchstring.value)"></span></td>
                  <td width="50%" align="right" background="/images/bg_search.gif"><img src="/images/btn_acc_reg.gif" width="78" height="18" alt="" border="0" class="account" onclick="pop_reg();" id="btnReg" style="cursor:hand;"></td>
                  <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="15" >&nbsp;</td>
          </tr>
          <tr>
            <td >
			<!--  -->

<%
	' 아이디 검색
	dim findID : findID = request("txtuserid")

	dim sql : sql = "select count(*) from wb_Account ; select c.custname as highcustname, b.custname, a.userid, a.class, a.isuse from wb_account a left outer join sc_cust_dtl b on a.custcode=b.custcode left outer join sc_cust_hdr c on b.highcustcode=c.highcustcode where userid like ? order by a.class; "

	dim cmd : set cmd = server.createobject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCMdTExt
	cmd.parameters.append cmd.createparameter("userid", adVarchar, adParamInput, 12)
	cmd.parameters("userid").value = "%"&findID&"%"

	dim objrs : set objrs = cmd.execute
	dim cnt : cnt = objrs(0)
	set objrs = objrs.nextrecordset

	dim highcustname, custname,  c_user_id, c_class, isuse
	if not objrs.eof then
		set highcustname = objrs("custname")
		set custname = objrs("custname")
		set c_user_id = objrs("userid")
		set c_class = objrs("class")
		set isuse = objrs("isuse")
	end if

%>

<table width="1030" height="31" border="3" cellpadding="0" cellspacing="0" bordercolor="#8D652B">
                <tr>
                  <td><table width="1024" border="0" cellspacing="0" cellpadding="0" class="header">
                      <tr>
                        <td width="44" align="center" class="header">No</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="240" align="center">광고주</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="240" align="center">운영팀</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="200" align="center" >아이디</td>
                        <td width="3" align="center" ><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="200" align="center" >권한</td>
                        <td width="3" align="center" ><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="100" align="center" >사용여부</td>
                      </tr>
                  </table></td>
                </tr>
              </table>
                <table width="1024" height="31" border="0" cellpadding="0" cellspacing="0" bordercolor="#8D652B">
				<% do until objrs.eof %>
                  <tr >
                    <td width="44" height="31" align="center"><%=cnt%></td>
                    <td width="3">&nbsp;</td>
                    <td width="240" align=""onClick="checkForView('<%=c_user_id%>')" class="styleLink" style="padding-left:10px;"><%=highcustname%></td>
                    <td width="3">&nbsp;</td>
                    <td width="240" align=""onClick="checkForView('<%=c_user_id%>')" class="styleLink" style="padding-left:10px;"><%=custname%>&nbsp;</td>
                    <td width="3">&nbsp;</td>
                    <td width="200" align="left" onClick="checkForView('<%=c_user_id%>')" class="styleLink header" style="padding-left:10px;"><%=c_user_id%>&nbsp;</td>
                    <td width="3" align="center">&nbsp;</td>
                    <td width="200" align="left" onClick="checkForView('<%=c_user_id%>')" class="styleLink" style="padding-left:10px;">
					<%
					select case c_class
						case  "A"	response.write "Administrator"
						case  "N"	response.write "Admin(Non-SKT)"
						case  "C"	response.write "광고주"
						case  "G"	response.write "광고주 관리자"
						case  "D"	response.write "운영팀"
						case  "H"	response.write "운영팀 관리자"
						case  "O"	 response.write "옥외 관리자"
						case  "F"	 response.write "옥외 모니터링"
						case  "M"	 response.write "매체사"
					end select
				%></td>
                    <td width="3" align="center">&nbsp;</td>
                    <td width="100" align="center"><%if ucase(isuse) = "Y" then response.write "사용중" Else response.write "사용중지"%>&nbsp;</td>
                  </tr>
                  <tr>
                    <td height="1" bgcolor="#E7E9E3" colspan="11"></td>
                  </tr>
				<%
						cnt = cnt - 1
						objrs.movenext
					loop
					objrs.close
					set objrs = nothing
				%>
            </table>
			<!--  -->
			</td>
          </tr>
      </table></td>
    </tr>
  </table>
<!--#include virtual="/bottom.asp" -->
</body>
</html>
<script language="JavaScript">
<!--
	function checkForView(uid) {
		var url = "/cust/admin/popup/account_view.asp?userid=" + uid;
		var name = "pop_account_view";
		var opt = "width=540, height=296, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		window.open(url, name, opt);
	}
	function pop_reg() {
		var p = document.getElementById("btnReg") ;
		var custcode = document.forms[0].tcustcode.value.replace("null","") ;
		if (p.getAttribute("class") == "account" || p.getAttribute("class") == null) {
			var url = "pop_account_reg.asp?tcustcode="+custcode;
			var name = "pop_reg";
			var opt = "width=540, height=366, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		} else {
			var url = "pop_menu_reg.asp?tcustcode="+custcode;
			var name ="pop_menu_reg" ;
			var opt = "width=540, height=205, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		}
		window.open(url, name, opt);
	}

	function checkForSearch(str) {
		var frm = document.forms[0];
		if (str !="") {
			if (str.indexOf("--") != -1) {
				alert("사용할 수 없는 문자를 입력하셨습니다.");
				frm.txtsearchstring.value = "";
				frm.txtsearchstring.focus();
				return false;
			}
		}
		frm.action = frm.actionurl.value;
		frm.method = "post";
		frm.submit();
	}

	document.onkeyup = function() {
		if (event.keyCode == "13") return false;
	}
//-->
</script>