<%@ language="vbscript" codepage="65001" %>
<!--#include virtual="/inc/func.asp" -->
<%

	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1000

	' 아이디 검색
	dim strUserid : strUserid = request("strUserid")
	dim findID : findID = request("findID")
	' 이름 검색
	dim findNAME : findNAME = request("findNAME")
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- 컨텐츠 영역 -->
<%
	dim sql , objrs,  cnt
	dim clientcode, highcustname,  userid
	

	sql = "select count(userid) cnt from wb_account_cust where userid = '" & strUserid & "'"

	Call get_recordset(objrs2, sql)
	if not objrs2.eof then
		cnt = objrs2("cnt")
	end If


%>

<link href="/style.css" rel="stylesheet" type="text/css">
<form name="framefrm_cust" >
		<input type="hidden" name="txtuserid">
		<input type="hidden" name="txtcustcode">
		<table width="203" border="0" cellpadding="0" cellspacing="0" bordercolor="#8D652B">
		<%

			sql = "select userid, clientcode, dbo.sc_get_highcustname_fun(clientcode) highcustname"
			sql = sql &  " from wb_account_cust "
			sql = sql &  " where userid = '" & strUserid & "'"

			Call get_recordset(objrs, sql)

			if not objrs.eof then
				do until objrs.eof
					userid = objrs("userid")
					clientcode = objrs("clientcode")
					highcustname = objrs("highcustname")
			%>
		  <tr onClick="checkStyle('<%=cnt%>','<%=userid%>','<%=clientcode%>');" id="view<%=cnt%>" >
			<td width="30" height="31" align="center"><%=cnt%></td>
			<td width="3">&nbsp;</td>
			<td width="180" align="left"  class="styleLink header" >
			<%=highcustname%>&nbsp;</td>
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

	
	function checkStyle(num,userid,custcode) {
		if (oldnum > 0 )
		{
			var div1 = document.getElementById("view"+oldnum);
			div1.style.backgroundColor = "white";
		}

		var div2 = document.getElementById("view"+num);
		div2.style.backgroundColor = "#8D652B";
		oldnum = num

		
		//parent.document.getElementById("mstruserid").innerText = userid;
		var txtcustcode = document.getElementById("txtcustcode");
		txtcustcode.innerText = custcode


		//광고주에해당하는팀 검색
		parent.cust_Class_src(userid,custcode);

		
	}
	
	
	window.onload = function () {
		
		var txtuserid = document.getElementById("txtuserid");
		txtuserid.innerText = parent.document.getElementById("mstruserid").value ;
		
	}




-->
</SCRIPT>



