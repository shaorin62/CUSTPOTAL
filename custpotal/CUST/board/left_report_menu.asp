<table width="210" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><input type="text" name="cookiemidx" id="cookiemidx" value ="<%=request.cookies("cookiemidx")%>">
			<input type="text" name="cookietitlename" id="cookietitlename"  value ="<%=request.cookies("cookietitlename")%>">
			<input type="text" name="cookiecustcode" id="cookiecustcode"  value ="<%=request.cookies("cookiecustcode")%>">
	</td>
  </tr>
  <tr>
    <td><img src="/images/menu_report_sub_title.gif" width="210" height="75"></td>
  </tr>
  <%
	sql = "select midx, title, lvl, ref from dbo.wb_menu_mst where custcode is null order by ref, lvl"
	call get_recordset(objrs, sql)

	if not objrs.eof then
		do until objrs.eof
			if objrs("lvl") = 1 then
%>
    <tr  id="<%=objrs("midx")%>"  style="display:block;">
		<td  width='210'  height='20' class="deps" background="/images/menu_bg.gif" id="<%=objrs("midx")%>"  onclick=" get_page(<%=objrs("midx")%>,'<%=objrs("title")%>','<%=objrs("midx")%>'); left_menu_getdata('list.asp','<%=objrs("midx")%>','','midx','<%=objrs("title")%>'); " style="cursor:hand"><%=objrs("title")%></td>
	</tr>
<% else %>
	<tr>
		<td  width='210'  height='20' class="deps_1" background="/images/menu_bg.gif" id="<%=objrs("midx")%>"  onclick="get_page(<%=objrs("midx")%>,'<%=objrs("title")%>','<%=objrs("midx")%>'); left_menu_getdata('list.asp','<%=objrs("midx")%>','','midx','<%=objrs("title")%>') " style="cursor:hand">&nbsp;&nbsp;&nbsp;&nbsp;<%=objrs("title")%> </td>
	</tr>
<%
		end if
		objrs.movenext
		loop
	end if

	dim objrs2, objrs3, objrs4
	sql = "select c.highcustcode custcode, c.custname from dbo.wb_menu_mst m inner join dbo.sc_cust_hdr c on m.custcode = c.highcustcode where c.use_flag = 1 and isnull(m.attr01,'') =''  group by c.highcustcode, c.custname order by c.custname"
	call get_recordset(objrs, sql)

	do until objrs.Eof
%>
	<tr>
		<td  width='210'  height='25' class="deps" background="/images/menu_dot_bg_over.gif"  onclick="checkDisplay('<%=objrs("custcode")%>');" style="padding-top:5px; cursor:hand;" ><%=objrs("custname")%></td>
	</tr>
	<tr id="<%=objrs("custcode")%>" style="display:none" >
		<td  width='210'  height='0'  background="/images/menu_bg.gif" >
			<%
				sql = "select midx, title, lvl from dbo.wb_menu_mst where isnull(attr01,'') ='' and  custcode = '" & objrs("custcode") &"' order by ref , lvl"

				call get_recordset(objrs2, sql)
			%>
			<table  border="0" cellspacing="0" cellpadding="0">
			<%
				do until objrs2.eof
					If objrs2("lvl") = 1 Then
			%>
				<tr>
					<td   height='19' class="deps_1" background="/images/menu_bg.gif" id="<%=objrs2("midx")%>" onclick="get_page(<%=objrs2("midx")%>, '<%=objrs2("title")%>','<%=objrs2("midx")%>'); left_menu_getdata('list.asp','<%=objrs2("midx")%>','<%=objrs("custcode")%>','highcustcode','<%=objrs2("title")%>')" style="cursor:hand;"><%=objrs2("title")%></td>
				</tr>
				<% Else %>
				<tr>
					<td   height='19' class="deps_1" background="/images/menu_bg.gif" id="<%=objrs2("midx")%>" onclick="get_page(<%=objrs2("midx")%>, '<%=objrs2("title")%>','<%=objrs2("midx")%>');  left_menu_getdata('list.asp','<%=objrs2("midx")%>','<%=objrs("custcode")%>','highcustcode','<%=objrs2("title")%>')" style="cursor:hand;">&nbsp;&nbsp;&nbsp;&nbsp;<%=objrs2("title")%></td>
				</tr>
				<%
					End If
					objrs2.movenext
					Loop
					objrs2.close
			%>
			</table>
			<%
				sql = "select c.custcode, c.custname from dbo.sc_cust_dtl c inner join dbo.wb_menu_mst m on m.custcode = c.custcode  where c.highcustcode = '" & objrs("custcode") & "' and c.use_flag = 1  and isnull(m.attr01,'') = 1 and c.gbnflag =0 group by c.custcode, c.custname order by c.custname"
				call get_recordset(objrs3, sql)
				if not objrs3.eof then
			%>
			<table  border="0" cellspacing="0" cellpadding="0">
			<%
				do until objrs3.eof
			%>
				<tr>
					<td   height='25' class="deps_1" background="/images/menu_bg_over.gif" ><%=objrs3("custname")%></td>
				</tr>
				<%
						sql = "Select  midx, title, lvl,  highmenu from dbo.wb_menu_mst where isnull(attr01,'') = 1 and  custcode = '" & objrs3("custcode") &"' order by ref , lvl"
						call get_recordset(objrs4, sql)

						do until objrs4.eof
						if objrs4("highmenu") = "0" then
				%>
				<tr>
					<td   height='19' class="deps_3" background="/images/menu_bg_over.gif" ><%=objrs4("title")%></td>
				</tr>
				<% else %>
						<% if objrs4("lvl") = 1 then %>
						<tr>
							<td   height='19' class="deps_3" background="/images/menu_bg.gif"  id="<%=objrs4("midx")%>"  onclick="get_page(<%=objrs4("midx")%>,'<%=objrs4("title")%>','<%=objrs4("midx")%>');  left_menu_getdata('list.asp','<%=objrs4("midx")%>','<%=objrs3("custcode")%>','timcode','<%=objrs4("title")%>')" style="cursor:hand"><%=objrs4("title")%></td>
						</tr>
							<% else %>
						<tr>
							<td   height='19' class="deps_3" background="/images/menu_bg.gif"  id="<%=objrs4("midx")%>"  onclick="get_page(<%=objrs4("midx")%>,'<%=objrs4("title")%>','<%=objrs4("midx")%>');  left_menu_getdata('list.asp','<%=objrs4("midx")%>','<%=objrs3("custcode")%>','timcode','<%=objrs4("title")%>')" style="cursor:hand">&nbsp;&nbsp;&nbsp;&nbsp;<%=objrs4("title")%></td>
						</tr>
						<% end if%>
				<%
					end if
				objrs4.movenext
				Loop
				objrs4.close

			objrs3.movenext
			Loop
			objrs3.close
		%>
		</table>
		<%	end if %>
		</td>
	</tr>
<%
	' 광고주 목록 돌리기
	objrs.moveNext
	loop
	objrs.movefirst
%>

  <tr>
    <td><img src="/images/menu_sub_bottom.gif" width="210" height="30"></td>
  </tr>
</table>
<SCRIPT LANGUAGE="JavaScript">
<!--

	function get_page(idx, str, p) {
		var td = document.getElementsByTagName("td")

		for (var i = 0 ; i < td.length ; i++) {
			if (!td[i].getAttribute("id")) continue;
			if (td[i].className == "depsover") td[i].className = "deps" ;
			if (td[i].className == "depsover_3") td[i].className = "deps_3" ;
			if (td[i].className == "depsover_1") td[i].className = "deps_1" ;
			if (td[i].className == "deps_3" || td[i].className == "deps_1" || td[i].className == "deps") {
				if (td[i].getAttribute("id") == p) {
					if (td[i].className == "deps") td[i].className = "depsover" ;
					if (td[i].className == "deps_3") td[i].className = "depsover_3" ;
					if (td[i].className == "deps_1") td[i].className = "depsover_1" ;
				}
			}
		}

		var subtitle = document.getElementById("subtitle") ;
		subtitle.innerText = str ;
		var navigate = document.getElementById("navi");
		navigate.innerText = "매체별 리포트 > " + str ;
		document.getElementById("midx").value = idx ;

	}

	function checkDisplay(code) {
		<% do until objrs.eof %>
		document.getElementById("<%=objrs("custcode")%>").style.display = "none";
		<%
				objrs.movenext
				loop
		%>
		document.getElementById(code).style.display = "block";
	}

	function checkDisplay_on(code) {
		document.getElementById(code).style.display = "block";
	}


	window.onload = function () {
		var cookiemidx = document.getElementById("cookiemidx").value ;
		var cookietitlename = document.getElementById("cookietitlename").value ;
		var cookiecustcode = document.getElementById("cookiecustcode").value ;

		if ( cookietitlename !="" ){
			alert(cookietitlename);
		}
		if ( cookiemidx == "" ) {
			cookieittlename = "전체 공지사항"
		}

		get_page(cookiemidx,cookietitlename,cookiemidx);
		left_menu_getdata('list.asp',cookiemidx,cookiecustcode,'',cookietitlename);

		if ( cookiecustcode != "") {
			checkDisplay_on(cookiecustcode);
		}


	}

//	window.onload = function () {
//		var td = document.getElementsByTagName("td")
//		for (var i = 0 ; i < td.length ; i++) {
//			if (!td[i].getAttribute("id")) continue;
//			if (td[i].className == "deps") {
//				td[i].className = "depsover" ;
//				document.getElementById("midx").value = td[i].getAttribute("id") ;
//				document.getElementById("scriptFrame").setAttribute("src", "list.asp");
//				return false;
//			}
//		}
//	}
//-->


</SCRIPT>

