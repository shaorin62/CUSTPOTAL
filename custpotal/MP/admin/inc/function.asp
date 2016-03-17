<%
	function getCommand(cmd, sql)
		set cmd = server.createobject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		getCommand = cmd
	end function
%>