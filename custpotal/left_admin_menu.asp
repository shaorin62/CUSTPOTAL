<table width="210" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="210" height="75"><img src="/images/menu_admin_sub_title.gif" width="210" height="75"></td>
  </tr>
  <tr>
    <td width="210" height="24"  <%if request.cookies("menunum") = "1" then Response.write "class='topmenuover'" Else response.write "class='topmenu'" End If%>><span onclick="location.href='/hq/admin/acc_list.asp?menuNum=1'">拌沥包府</span></td>
  </tr>
  <tr>
    <td width="210" height="25" class='menu'>皋春包府</td>
  </tr>  
  <tr style="display:block;"  id="outdoormenu1">
    <td  width="210" >	
		<table width="210" border="0" cellspacing="0" cellpadding="0">
		<% 
			Dim objrs2
			sql = "select custcode, custname from dbo.sc_cust_temp where custcode = highcustcode and medflag ='A' "
			Call get_recordset(objrs2, sql)
			Do Until objrs2.eof 
			if mnu_menu = "1" then 
		%>
			<tr>
				<td  width="210" height="25" class='submenuover' ><span onclick="location.href='acc_list.asp?menuNum=<%=objrs2("custcode")%>'"><%=objrs2("custname")%></span></td>
		  </tr>
		<%else%>
		  <tr>
			<td  width="210" height="25" <%if mnu_menu = objrs2("custcode")  then Response.write "class='submenuover'" Else response.write "class='submenu'" End If%> ><span onclick="location.href='menu_list.asp?menuNum=<%=objrs2("custcode")%>'"><%=objrs2("custname")%></span></td>
		  </tr>
		  <%
			end if
				objrs2.movenext
				Loop
				objrs2.close
				Set objrs2 = nothing
		  %>
		 </table>
	</td>
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
