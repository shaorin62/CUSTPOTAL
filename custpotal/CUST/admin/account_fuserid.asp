<%@ language="vbscript" codepage="65001" %>
<!--#include virtual="/inc/func.asp" -->
<%

	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1000

	' 아이디 검색
	dim findID : findID = request("findID")
	' 이름 검색
	dim findNAME : findNAME = request("findNAME")

%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- 컨텐츠 영역 -->
<%
	dim sql , objrs, objrs2 , cnt
	dim highcustname, custname,  userid, username,  c_class, isuse

	sql = "select count(userid) cnt from wb_Account where userid like '%" & findID & "%'"
	Call get_recordset(objrs2, sql)
	if not objrs2.eof then
		cnt = objrs2("cnt")
	end If
%>

<link href="/style.css" rel="stylesheet" type="text/css">
<form name="framefrm" >
		<input type="hidden" name="txtuserid">
		<input type="hidden" name="txtclass">
		<table width="360" border="0" cellpadding="0" cellspacing="0" bordercolor="#8D652B">
		<%

			sql = "select userid, username, class, isuse "
			sql = sql &  " from wb_account "
			sql = sql &  " where userid like '%" & findID & "%' order by class"

			Call get_recordset(objrs, sql)

			if not objrs.eof then
				do until objrs.eof
					userid = objrs("userid")
					username = objrs("username")
					c_class = objrs("class")
					isuse = objrs("isuse")
			%>
		  <tr onClick="checkStyle('<%=cnt%>','<%=userid%>','<%=c_class%>');" id="view<%=cnt%>" >
			<td width="30" height="31" align="center"><%=cnt%></td>
			<td width="3">&nbsp;</td>
			<td width="132" align="left" onDblClick="checkForView('<%=userid%>','<%=c_class%>')" class="styleLink header" >
			<%=userid%>&nbsp;</td>
			<td width="3" align="center">&nbsp;</td>
			<td width="122" align="left" onDblClick="checkForView('<%=userid%>','<%=c_class%>')" class="styleLink header" style="padding-left:10px;">
			<%=username%>&nbsp;</td>
			<td width="3" align="center">&nbsp;</td>
			<td width="87" align="left" onDblClick="checkForView('<%=userid%>','<%=c_class%>')" class="styleLink" style="padding-left:10px;">
			<%
			select case c_class
				case  "A"	response.write "Administrator"
				case  "G"	response.write "MP"
				case  "C"	response.write "광고주"
				case  "F"	response.write "모니터링"
			end select
		%></td>
		  </tr>
		   <tr>
			<td height="1" bgcolor="#E7E9E3" colspan="13"></td>
		  </tr>
		<%
					cnt = cnt - 1
					objrs.movenext
				loop
			end If

			objrs.close
			set objrs = nothing
		%>
	</table>
</form>

<SCRIPT LANGUAGE="JavaScript">
<!--	
	var oldnum

	function checkForView(uid,c_class) {
		var url = "pop_account_view.asp?userid=" + uid + "&c_class=" + c_class;
		var name = "pop_account_view";
		var opt = "width=540, height=320, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		window.open(url, name, opt);
	}
	
	function checkStyle(num,userid,c_class) {
		if (oldnum > 0 )
		{
			var div1 = document.getElementById("view"+oldnum);
			div1.style.backgroundColor = "white";
		}

		var div2 = document.getElementById("view"+num);
		div2.style.backgroundColor = "#8D652B";
		oldnum = num

		var txtuserid = document.getElementById("txtuserid");
		txtuserid.innerText = userid ;

		var txtclass = document.getElementById("txtclass");
		txtclass.innerText = c_class ;


		parent.document.getElementById("mstruserid").innerText = userid;
		parent.user_Class_src(userid);
	}





-->
</SCRIPT>



