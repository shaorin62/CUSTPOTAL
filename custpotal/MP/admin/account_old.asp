<%@ language="vbscript" codepage="65001" %>
<!--#include virtual="/inc/func.asp" -->
<%

	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1000


	' 아이디 검색
	'dim findID : findID = request("txtsearchstring")
	dim findID : findID = ""

	dim sql : sql = "select count(userid) from wb_Account where isuse='Y'; select c.custname as highcustname, b.custname, a.userid, a.class, a.isuse from wb_account a inner join sc_cust_dtl b on a.custcode=b.custcode inner join sc_cust_hdr c on b.highcustcode=c.highcustcode where userid like ? order by a.class; "
	
	dim cmd : set cmd = server.createobject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCMdTExt
	cmd.parameters.append cmd.createparameter("userid", adVarchar, adParamInput, 12)
	cmd.parameters("userid").value = "%"&findID&"%"

	dim objrs : set objrs = cmd.execute
	dim cnt : cnt = objrs(0)
	set objrs = objrs.nextrecordset

	dim highcustname, custname,  userid,  c_class, isuse
	if not objrs.eof then
		set highcustname = objrs("custname")
		set custname = objrs("custname")
		set userid = objrs("userid")
		set c_class = objrs("class")
		set isuse = objrs("isuse")
	end if


%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/style.css" rel="stylesheet" type="text/css">
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
			  <div id='#contents' style='margin-top:10px;width:1030px;overflow-x:scroll;'>
                <table width="1024" border="0" cellpadding="0" cellspacing="0" bordercolor="#8D652B">
				<% do until objrs.eof %>
                  <tr >
                    <td width="44" height="31" align="center"><%=cnt%></td>
                    <td width="3">&nbsp;</td>
                    <td width="240" align=""onClick="checkForView('<%=userid%>')" class="styleLink" style="padding-left:10px;"><%=highcustname%></td>
                    <td width="3">&nbsp;</td>
                    <td width="240" align=""onClick="checkForView('<%=userid%>')" class="styleLink" style="padding-left:10px;"><%=custname%>&nbsp;</td>
                    <td width="3">&nbsp;</td>
                    <td width="200" align="left" onClick="checkForView('<%=userid%>')" class="styleLink header" style="padding-left:10px;"><%=userid%>&nbsp;</td>
                    <td width="3" align="center">&nbsp;</td>
                    <td width="200" align="left" onClick="checkForView('<%=userid%>')" class="styleLink" style="padding-left:10px;">
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
</div>
<script language="JavaScript">
<!--
	function checkForView(uid) {
		var url = "pop_account_view.asp?userid=" + uid;
		var name = "pop_account_view";
		var opt = "width=540, height=306, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		window.open(url, name, opt);
	}
//	function checkForSearch(str) {
//		var frm = document.forms[0];
//		if (str !="") {
//			if (str.indexOf("--") != -1) {
//				alert("사용할 수 없는 문자를 입력하셨습니다.");
//				frm.txtsearchstring.value = "";
//				frm.txtsearchstring.focus();
//				return false;
//			}
//		}
//		frm.txtsearchstring.value = str;
//		frm.action = "acc_list.asp";
//		frm.method = "post";
//		frm.submit();
//	}
//-->
</script>