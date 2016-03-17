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
<SCRIPT LANGUAGE="JavaScript" src="/js/secure.js"></SCRIPT>