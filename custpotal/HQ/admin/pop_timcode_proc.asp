<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%

	Dim intLoop 
	Dim userid : userid = request.form("strUserid")
	Dim Custcode : Custcode = request.form("strCustcode")
	
	Dim sql , objrs
	Dim pk


	sql = "select userid, clientcode, timcode, cuser, cdate from dbo.wb_account_tim where userid = '" & userid & "'"
	
	Call set_recordset(objrs, sql)

	For intLoop = 1 To Request.Form("timidx").count	
		pk = Split(Request.Form("timidx")(intLoop), ",")
		
		objrs.addnew
		objrs.fields("USERID").value = userid
		objrs.fields("CLIENTCODE").value = Custcode
		objrs.fields("TIMCODE").value = pk(0)
		objrs.fields("CUSER").value = Request.Cookies("userid")
		objrs.fields("CDATE").value = date
		objrs.update

		'Request.Cookies("userid")   / session("userid")
	Next

	objrs.close
	Set objrs = Nothing



%>
<script type="text/javascript">
<!--
	//	parent.opener.location.href="account_fcust.asp?strUserid="+userid;

		
	parent.opener.frmtim.location.href="account_ftim.asp?strUserid=<%=userid%>&strCustcode=<%=Custcode%>";

	//parent.opener.frmcust.reload();

	this.close();
-->
</script>