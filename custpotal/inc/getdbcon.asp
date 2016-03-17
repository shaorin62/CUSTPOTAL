<%@LANGUAGE="VBSCRIPT"%>
<%
	Option Explicit
	dim objconn : set objconn = server.createobject("adodb.connection")

	 Sub dbcon
		objconn.connectionstring = application("connectionstring")
		objconn.open
	 End Sub

	Sub dbclose
		objconn.close
		Set objconn = nothing
	End sub

%>

<%
	dim def_userid : def_userid = request.cookies("userid")
	If def_userid = "" Then
		response.write "<script type='text/javascript'> parent.location.href = '/'; </script>"
		response.end
	End if
%>
<SCRIPT LANGUAGE="JavaScript" src="/js/secure.js"></SCRIPT>