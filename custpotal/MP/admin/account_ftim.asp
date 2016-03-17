<%@ language="vbscript" codepage="65001" %>
<!--#include virtual="/inc/func.asp" -->
<%

	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1000

	' 아이디 검색
	dim strUserid : strUserid = request("strUserid")
	dim strCustcode : strCustcode = request("strCustcode")

	dim findID : findID = request("findID")
	' 이름 검색
	dim findNAME : findNAME = request("findNAME")
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- 컨텐츠 영역 -->
<%
	dim sql , objrs,  cnt
	dim clientcode, timcode, userid
	

	sql = "select count(userid) cnt from wb_account_tim where userid = '" & strUserid & "' and clientcode = '" & strCustcode & "'"

	
	Call get_recordset(objrs2, sql)
	if not objrs2.eof then
		cnt = objrs2("cnt")
	end If


%>

<link href="/style.css" rel="stylesheet" type="text/css">
<form name="framefrm_tim" >
		<input type="hidden" name="txtuserid">
		<input type="hidden" name="txttimcode">
		<table width="203" border="0" cellpadding="0" cellspacing="0" bordercolor="#8D652B">
		<%

			sql = "select userid, clientcode, timcode, dbo.sc_get_custname_fun(timcode) timname "
			sql = sql &  " from wb_account_tim "
			sql = sql &  " where userid = '" & strUserid & "' and clientcode = '" & strCustcode & "'"
			
			Call get_recordset(objrs, sql)

			if not objrs.eof then
				do until objrs.eof
					userid = objrs("userid")
					clientcode = objrs("clientcode")
					timcode = objrs("timcode")
					timname = objrs("timname")
			%>
		  <tr onClick="checkStyle('<%=cnt%>','<%=userid%>','<%=timcode%>');" id="view<%=cnt%>" >
			<td width="30" height="31" align="center"><%=cnt%></td>
			<td width="3">&nbsp;</td>
			<td width="180" align="left"  class="styleLink header" >
			<%=timname%>&nbsp;</td>
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

	
	function checkStyle(num,userid,timcode) {
		if (oldnum > 0 )
		{
			var div1 = document.getElementById("view"+oldnum);
			div1.style.backgroundColor = "white";
		}

		var div2 = document.getElementById("view"+num);
		div2.style.backgroundColor = "#8D652B";
		oldnum = num

		
		var txttimcode = document.getElementById("txttimcode");
		txttimcode.innerText = timcode
	}
	
	
	window.onload = function () {
		
		var txtuserid = document.getElementById("txtuserid");
		txtuserid.innerText = parent.document.getElementById("mstruserid").value ;
		
	}




-->
</SCRIPT>



