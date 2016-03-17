<table width="210" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="/images/menu_report_sub_title.gif" width="210" height="75"></td>
  </tr>
  <tr>
    <td width="210" height="25"   <%if cstr(mtitle) = "공지사항" then Response.write "class='headermenuover'" Else response.write "class='headermenu'" End If%>><span onclick="location.href='list.asp?midx=<%=midx%>'">공지사항</span></td>
  </tr>
  <%
	dim objcustcode, objcustcode2, objmenu, cust_code, cust_name, cust_code2, cust_name2, menu_idx, menu_title, menu_lvl, iscustmenu, iscustcode, menu_issub
	iscustmenu = false
	iscustcode = false
	sql = "select custcode, custname from dbo.sc_cust_temp where medflag = 'A' and highcustcode = custcode"
	call get_recordset(objcustcode, sql)
	if not objcustcode.eof then 
	set cust_code = objcustcode("custcode")
	set cust_name = objcustcode("custname")
	do until objcustcode.eof 
	response.write "<tr><td  width='210'  height='25'  class='headermenu' style='stylelink' onclick=""location.href='list.asp?selcustcode="&cust_code&"'"">"& cust_name&"</td></tr>"
	if cust_code = custcode then 
		
		sql = "select midx, title, lvl, ref, (select max(lvl) from dbo.wb_menu_mst where ref=m.midx) as issub from dbo.wb_menu_mst m where custcode = '" & custcode &"' "
		call get_recordset(objmenu, sql)
		if not objmenu.eof then 
			set menu_idx = objmenu("midx")
			set menu_title = objmenu("title")
			set menu_lvl = objmenu("lvl")
			set menu_issub= objmenu("issub")
			response.write "<tr><td>"
			iscustmenu = true
			response.write "<table width='210' border='0' cellspacing='0' cellpadding='0'>"
				do until objmenu.eof 
					response.write "<tr><td  width='210'  height='25'  style='stylelink' "
						If menu_lvl = 1 Then 
							if cint(menu_idx) = cint(midx) then 	response.write "class = 'subheadermenuover' " 	else 	response.write " class='subheadermenu' " 
						Else
							if cint(menu_idx) = cint(midx) then 	response.write "class = 'menulistover' " 	else 	response.write " class='menulist' " 
						End if
						if (menu_lvl = 1 And menu_issub = 2)  then 	
							response.write "> <B style='color:#FFCC33;'>| </B>"& menu_title&"</td></tr>"
						Else 
							response.write " onclick=""get_report("&menu_idx&",'"&custcode&"','','"&menu_title&"')""  >"& menu_title&"</td></tr>"
						end if
				objmenu.movenext
				Loop
				objmenu.close
			response.write "</table>"
		end if

		
		sql = "select custcode, custname from dbo.sc_cust_temp where medflag='A' and highcustcode <> custcode and highcustcode <> 'A00000' and highcustcode = '" & cust_code &"' "
		call get_recordset(objcustcode2, sql)
		if not objcustcode2.eof then
			set cust_code2 = objcustcode2("custcode")
			set cust_name2 = objcustcode2("custname")
			response.write "<tr><td>"
			iscustcode = true
				response.write "<table width='210' border='0' cellspacing='0' cellpadding='0'>"
				do until objcustcode2.eof 
					response.write "<tr><td  width='210'  height='25'  class='subheadermenu' style='stylelink' > <B style='color:#FFCC33;'>| </B>"& cust_name2&"</td></tr>"
					' 사업부 메뉴가
		sql = "select midx, title, lvl, ref, (select max(lvl) from dbo.wb_menu_mst where ref=m.midx) as issub from dbo.wb_menu_mst m where custcode = '" & cust_code2 &"' "
						call get_recordset(objmenu, sql)
						if not objmenu.eof then 
						set menu_idx = objmenu("midx")
						set menu_title = objmenu("title")
						set menu_lvl = objmenu("lvl")
						set menu_issub = objmenu("issub")
							response.write "<tr><td><table width='210' border='0' cellspacing='0' cellpadding='0'>"
							do until objmenu.eof 
								If menu_lvl = 1 And menu_issub <> 1 Then 
								response.write "<tr><td  width='210'  height='25'   class='menulist'> "& menu_title&"</td></tr>"
								Else
								response.write "<tr><td  width='210'  height='25'  "
									If menu_lvl = 1 Then 
										if cint(menu_idx)  = cint(midx) then response.write " class = 'menulistover' " else response.write " class='menulist' "
									Else
										if cint(menu_idx)  = cint(midx) then response.write " class = 'menulistover2' " else response.write " class='menulist2' "
									End if
								response.write " onclick=""get_report("&menu_idx&",'"&cust_code&"', '" &cust_code2&"','"&menu_title&"')""> "& menu_title&"</td></tr>"
								End if
							objmenu.movenext
							loop
							response.write "</table></td></tr>"
						objmenu.close
						end if
					' 사업부 메뉴
				objcustcode2.movenext
				Loop
				objcustcode2.close
				response.write "</table>"
			if iscustmenu and iscustcode then response.write "</td></tr>"
		end if
	iscustmenu = false
	iscustcode = false
	end if
	objcustcode.movenext
	loop
	objcustcode.close
	end if
  %>
  <tr>
    <td><img src="/images/menu_sub_bottom.gif" width="210" height="30"></td>
  </tr>
</table>
<SCRIPT LANGUAGE="JavaScript">
<!--
	function get_report(idx, code, code2, title) {
		location.href="list.asp?midx="+idx+"&selcustcode="+code+"&selcustcode2="+code2+"&mtitle="+title;
	}
//-->
</SCRIPT>

