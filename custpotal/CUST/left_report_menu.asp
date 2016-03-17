<table width="210" border="0" cellspacing="0" cellpadding="0" background="/images/menu_bg.gif">
  <tr>
    <td><input type="hidden" name="cookiemidx" id="cookiemidx" value ="<%=request.cookies("cookiemidx")%>">
			<input type="hidden" name="cookietitlename" id="cookietitlename"  value ="<%=unescape(request.cookies("cookietitlename"))%>">
			<input type="hidden" name="cookiecustcode" id="cookiecustcode"  value ="<%=request.cookies("cookiecustcode")%>">
			<input type="hidden" name="cookiesearchstring" id="cookiesearchstring"  value ="<%=request.cookies("cookiesearchstring")%>">
			<input type="hidden" name="cookiehighcategory" id="cookiehighcategory"  value ="<%=request.cookies("cookiehighcategory")%>">
			<input type="hidden" name="cookiecategory" id="cookiecategory"  value ="<%=request.cookies("cookiecategory")%>">
			<input type="hidden" name="cookieattr02" id="cookieattr02"  value ="<%=request.cookies("cookieattr02")%>">
	</td>
  </tr>
  <tr>
    <td><img src="/images/menu_report_sub_title.gif" width="210" height="75"></td>
  </tr>
  <%
	'sql = "select midx, title, lvl, ref, CASE ISNULL(dbo.WB_GET_REPORT_NEWIMAGE(midx),'') WHEN '' THEN '999999' ELSE dbo.WB_GET_REPORT_NEWIMAGE(midx) END NewImageCnt from dbo.wb_menu_mst where custcode is null order by ref, lvl"

	sql =" select midx, title, lvl, ref, "
	sql = sql & " CASE ISNULL(dbo.WB_GET_REPORT_NEWIMAGE(midx),'')  "
	sql = sql & " WHEN '' THEN '999999'  "
	sql = sql & " ELSE dbo.WB_GET_REPORT_NEWIMAGE(midx)  "
	sql = sql & " END NewImageCnt , midx1, "
	sql = sql & " dbo.WB_GET_REPORT_PLUSIMAGE(midx) PlusImageCnt, attr02"
	sql = sql & " from ( "
	sql = sql & " 	select midx, midx midx1, title, lvl, ref, attr02 "
	sql = sql & " 	from dbo.wb_menu_mst  "
	sql = sql & " 	where custcode is null "
	sql = sql & " 	and lvl = 1 and isnull(mp,0) = 1 "
	sql = sql & " 	union all "
	sql = sql & " 	select midx, ref midx1, title, lvl, midx ref, attr02 "
	sql = sql & " 	from dbo.wb_menu_mst  "
	sql = sql & " 	where custcode is null "
	sql = sql & " 	and lvl = 2 and isnull(mp,0) = 1 "
	sql = sql & " 	union all "
	sql = sql & " 	select a.midx, b.ref midx1, a.title, a.lvl, a.ref, attr02 "
	sql = sql & " 	from dbo.wb_menu_mst a "
	sql = sql & " 	inner join ( "
	sql = sql & " 		select midx, ref  "
	sql = sql & " 		from wb_menu_mst "
	sql = sql & " 		where custcode is null "
	sql = sql & " 		and lvl  = 2 "
	sql = sql & " 	)b on a.ref = b.midx "
	sql = sql & " 	where a.custcode is null "
	sql = sql & " 	and a.lvl = 3 and isnull(mp,0) = 1 "
	sql = sql & " ) a "
	sql = sql & " order by midx1, ref, lvl "

	call get_recordset(objrs, sql)

	if not objrs.eof then
		do until objrs.eof
			if objrs("lvl") = 1 Then
%>
    <tr  id="<%=objrs("midx")%>"  style="display:block;">
		<td  width='210'  height='20' class="deps" background="/images/menu_bg.gif" id="<%=objrs("midx")%>"  onclick="insertLogmst_src('<%=objrs("midx")%>'); get_page(<%=objrs("midx")%>,'<%=objrs("title")%>','<%=objrs("midx")%>'); get_pageSRC('list.asp','<%=objrs("midx")%>','','midx','<%=objrs("title")%>', 1, '<%=objrs("attr02")%>');" style="cursor:hand"><%If objrs("PlusImageCnt") <= 1 Then %>&nbsp;&nbsp;&nbsp;<%End if%><img src="/images/mms_plus.gif" id="A<%=objrs("midx")%>_B<%=objrs("ref")%>_C<%=objrs("midx1")%>" value="0" border="0" width="10px" onclick="menuclick1(this.id, <%=objrs("midx")%>);" alt="" <%If objrs("PlusImageCnt") <= 1 Then %> style="cursor:hand;display:none;" <%End if%>><%If objrs("PlusImageCnt") > 1 Then %>&nbsp;<%End if%><%=objrs("title")%>&nbsp;<%if objrs("NewImageCnt") <= 7 then %><img src="/images/new.gif" width="21" height="10" border="0" alt=""><%end if%></td>
	</tr>
<% ElseIf objrs("lvl") = 2 then %>
	<tr id="A<%=objrs("midx")%>-B<%=objrs("ref")%>-trC<%=objrs("midx1")%>-"  style="display:none;">
		<td  width='210'  height='20' class="deps_1" background="/images/menu_bg.gif" id="A<%=objrs("midx")%>-B<%=objrs("ref")%>-C<%=objrs("midx1")%>-"  onclick="insertLogmst_src('<%=objrs("midx")%>');get_page(<%=objrs("midx")%>,'<%=objrs("title")%>',this.id);  get_pageSRC('list.asp','<%=objrs("midx")%>','','midx','<%=objrs("title")%>', 1, '<%=objrs("attr02")%>') " style="cursor:hand;display:none;"><%If objrs("PlusImageCnt") <= 0 Then %>&nbsp;&nbsp;&nbsp;&nbsp;<%End if%><img src="/images/mms_plus.gif" id="A<%=objrs("midx")%>_B<%=objrs("ref")%>_C<%=objrs("midx1")%>_"  value="0" border="0" width="8px" onclick="menuclick2(this.id, <%=objrs("midx")%>);"  alt="" <%If objrs("PlusImageCnt") <= 0 Then %> style="cursor:hand;display:none;" <%End if%>><%If objrs("PlusImageCnt") > 0 Then %>&nbsp; <%End if%><%=objrs("title")%>&nbsp;<%if objrs("NewImageCnt") <= 7 then %><img src="/images/new.gif" width="21" height="10" border="0" alt=""><%end if%></td>
	</tr>
<% Else %>
	<tr id="A<%=objrs("midx")%>-trB<%=objrs("ref")%>-C<%=objrs("midx1")%>-" style="display:none;">
		<td  width='210'  height='20' class="deps_2" background="/images/menu_bg.gif" id="A<%=objrs("midx")%>-B<%=objrs("ref")%>-C<%=objrs("midx1")%>-"  onclick="insertLogmst_src('<%=objrs("midx")%>');get_page(<%=objrs("midx")%>,'<%=objrs("title")%>',this.id);  get_pageSRC('list.asp','<%=objrs("midx")%>','','midx','<%=objrs("title")%>', 1, '<%=objrs("attr02")%>') " style="cursor:hand; display:none;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=objrs("title")%>&nbsp;<%if objrs("NewImageCnt") <= 7 then %><img src="/images/new.gif" width="21" height="10" border="0" alt=""><%end if%></td>
	</tr>
<%
		end if
		objrs.movenext
		loop
	end if

	dim objrs2, objrs3, objrs4

	sql = " select dbo.SC_GET_HIGHCUSTCODE_FUN(custcode) custcode, "
	sql = sql & " dbo.sc_get_highcustname_fun(dbo.SC_GET_HIGHCUSTCODE_FUN(custcode)) custname "
	sql = sql & " from wb_menu_mst "
	sql = sql & " where isnull(custcode,'') <> '' "
	sql = sql & " and dbo.SC_GET_HIGHCUSTCODE_FUN(custcode) in( "
	sql = sql & " 	select clientcode  "
	sql = sql & " 	from wb_account_cust  "
	sql = sql & " 	where userid ='"& request.cookies("userid") &"'  "
	sql = sql & " 	group by clientcode "
	sql = sql & " ) "
	sql = sql & " group by dbo.SC_GET_HIGHCUSTCODE_FUN(custcode) order by custname "

	call get_recordset(objrs, sql)
	if not objrs.eof then
	do until objrs.Eof
%>
	<tr>
		<td  width='210'  height='25' class="deps" background="/images/menu_dot_bg_over.gif"  onclick="checkDisplay('<%=objrs("custcode")%>');" style="padding-top:5px; cursor:hand;" ><%=objrs("custname")%></td>
	</tr>

	<tr id="<%=objrs("custcode")%>" style="display:none" >
		<td  width='210'  height='0'  background="/images/menu_bg.gif" >
			<%
				sql = " select count(*) cnt from wb_account_tim where userid = '"& request.cookies("userid") &"'  and clientcode = '" & objrs("custcode") & "' "
				call get_recordset(objrs2, sql)

				If objrs2("cnt") = 0 Then
					sql = " select custcode, dbo.sc_get_custname_fun(custcode) custname "
					sql = sql & " from dbo.wb_menu_mst "
					sql = sql & " where isnull(attr01,'') = 1  "
					sql = sql & " and dbo.sc_get_highcustcode_fun(custcode) =  '" & objrs("custcode") & "' "
					sql = sql & " group by custcode "
					sql = sql & " order by custname "
				Else
					sql = " select custcode, dbo.sc_get_custname_fun(custcode) custname "
					sql = sql & " from dbo.wb_menu_mst "
					sql = sql & " where isnull(attr01,'') = 1  "
					sql = sql & " and custcode in( "
					sql = sql & " 	select timcode  "
					sql = sql & " 	from wb_account_tim  "
					sql = sql & " 	where userid ='"& request.cookies("userid") &"'  "
					sql = sql & " 	and clientcode = '" & objrs("custcode") & "' "
					sql = sql & " 	group by timcode "
					sql = sql & " ) "
					sql = sql & " group by custcode "
					sql = sql & " order by custname "
				End if

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
						sql =" select midx, title, lvl, ref, highmenu, "
						sql = sql & " CASE ISNULL(dbo.WB_GET_REPORT_NEWIMAGE(midx),'')  "
						sql = sql & " WHEN '' THEN '999999'  "
						sql = sql & " ELSE dbo.WB_GET_REPORT_NEWIMAGE(midx)  "
						sql = sql & " END NewImageCnt, attr02"
						sql = sql & " from dbo.wb_menu_mst  "
						sql = sql & " where isnull(attr01,'') = 1 and isnull(custcode,'') = '" & objrs3("custcode") &"' "
						sql = sql & " order by ref, lvl "

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
							<td   height='19' class="deps_3" background="/images/menu_bg.gif"  id="<%=objrs4("midx")%>"  onclick="insertLogmst_src('<%=objrs4("midx")%>');get_page(<%=objrs4("midx")%>,'<%=objrs4("title")%>','<%=objrs4("midx")%>');  get_pageSRC('list.asp','<%=objrs4("midx")%>','','midx','<%=objrs4("title")%>', 1, '<%=objrs4("attr02")%>');" style="cursor:hand"><%=objrs4("title")%>&nbsp;<%if objrs4("NewImageCnt") <= 7 then %><img src="/images/new.gif" width="21" height="10" border="0" alt=""><%end if%></td>
						</tr>
							<% else %>
						<tr>
							<td   height='19' class="deps_3" background="/images/menu_bg.gif"  id="<%=objrs4("midx")%>"  onclick="insertLogmst_src('<%=objrs4("midx")%>');get_page(<%=objrs4("midx")%>,'<%=objrs4("title")%>','<%=objrs4("midx")%>');  get_pageSRC('list.asp','<%=objrs4("midx")%>','','midx','<%=objrs4("title")%>', 1, '<%=objrs4("attr02")%>');" style="cursor:hand">&nbsp;&nbsp;&nbsp;&nbsp;<%=objrs4("title")%>&nbsp;<%if objrs4("NewImageCnt") <= 7 then %><img src="/images/new.gif" width="21" height="10" border="0" alt=""><%end if%></td>
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
	end if
%>

  <tr>
    <td><img src="/images/menu_sub_bottom.gif" width="210" height="30"></td>
  </tr>
</table>
<div id="logpage"></div>
<SCRIPT LANGUAGE="JavaScript">
<!--

	function get_page(idx, str, p) {
		var td = document.getElementsByTagName("td")

		for (var i = 0 ; i < td.length ; i++) {
			if (!td[i].getAttribute("id")) continue;
			if (td[i].className == "depsover") td[i].className = "deps" ;
			if (td[i].className == "depsover_3") td[i].className = "deps_3" ;
			if (td[i].className == "depsover_2") td[i].className = "deps_2" ;
			if (td[i].className == "depsover_1") td[i].className = "deps_1" ;
			if (td[i].className == "deps_3" || td[i].className == "deps_2" || td[i].className == "deps_1" || td[i].className == "deps") {
				if (td[i].getAttribute("id") == p) {
					if (td[i].className == "deps") td[i].className = "depsover" ;
					if (td[i].className == "deps_3") td[i].className = "depsover_3" ;
					if (td[i].className == "deps_2") td[i].className = "depsover_2" ;
					if (td[i].className == "deps_1") td[i].className = "depsover_1" ;
				}
			}
		}

		var subtitle = document.getElementById("subtitle") ;
		subtitle.innerText = str ;
		var navigate = document.getElementById("navi");
		navigate.innerText = "매체별 리포트 > " + str ;
		//document.getElementById("midx").value = idx ;
		//document.getElementById("ttitle").value = idx ;
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
		var cookiesearchstring = document.getElementById("cookiesearchstring").value ;
		var cookiehighcategory = document.getElementById("cookiehighcategory").value ;
		var cookiecategory = document.getElementById("cookiecategory").value ;
		var attr02 = document.getElementById("cookieattr02").value ;



		if ( cookiemidx == "" ) {
			cookietitlename = "공지사항"
		}

		get_page(cookiemidx,cookietitlename,cookiemidx);
		get_pageSRC('list.asp',cookiemidx,cookiecustcode,'',cookietitlename, 1, attr02);
		//left_menu_getdata('list.asp',cookiemidx,cookiecustcode,'',cookietitlename);
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

	function insertLogmst_src(progname) {
		var params = "progname="+progname;
		sendRequest("/inc/setlog.asp", params, _insertLogmst_src, "GET");
	}

	function _insertLogmst_src() {
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				var custcode = document.getElementById("logpage");

				custcode.innerHTML = xmlreq.responseText ;

			}
		}
	}

	function menuclick1(imgid, A){
		var td = document.getElementsByTagName("td")
		var tr = document.getElementsByTagName("tr")
		var strmidx1
		var strtr
		var strimgvalue
		strmidx1 = 'C' + A + '-';
		strtr = 'trC' + A + '-';

		strimgvalue = document.getElementById(imgid).value;

		if (strimgvalue == "0")	{
			document.getElementById(imgid).src = "/images/mms_minus.gif";
			document.getElementById(imgid).value = "1";
		}else{
			document.getElementById(imgid).src = "/images/mms_plus.gif";
			document.getElementById(imgid).value = "0";
		}

		for (var i = 0 ; i < tr.length ; i++) {
			if (!tr[i].getAttribute("id")) continue;
			if (tr[i].getAttribute("id").indexOf(strtr) != -1){
				if (strimgvalue =="0"){
					document.getElementById(tr[i].getAttribute("id")).style.display ="";
				}else{
					document.getElementById(tr[i].getAttribute("id")).style.display ="none";
				}
			}
		}

		for (var i = 0 ; i < td.length ; i++) {
			if (!td[i].getAttribute("id")) continue;
			if (td[i].getAttribute("id").indexOf(strmidx1) != -1){
				if (td[i].className == "deps_1"){
					if (strimgvalue =="0"){
						document.getElementById(td[i].getAttribute("id")).style.display ="";
						document.getElementById(td[i].getAttribute("id").replace(/-/g,"_")).src = "/images/mms_plus.gif";
						document.getElementById(td[i].getAttribute("id").replace(/-/g,"_")).value = "0";
					}else{
						document.getElementById(td[i].getAttribute("id")).style.display ="none";
						document.getElementById(td[i].getAttribute("id").replace(/-/g,"_")).src = "/images/mms_minus.gif";
						document.getElementById(td[i].getAttribute("id").replace(/-/g,"_")).value = "1";
					}
				}else if (td[i].className == "deps_2"){
					document.getElementById(td[i].getAttribute("id")).style.display ="none";
				}
			}
		}
	}


	function menuclick2(imgid, A){
		var td = document.getElementsByTagName("td")
		var tr = document.getElementsByTagName("tr")
		var strmidx1
		var strtr
		var strimgvalue
		strmidx1 = 'B' + A + '-';
		strtr = 'trB' + A + '-';

		strimgvalue = document.getElementById(imgid).value;

		if (strimgvalue == "0")	{
			document.getElementById(imgid).src = "/images/mms_minus.gif";
			document.getElementById(imgid).value = "1";
		}else{
			document.getElementById(imgid).src = "/images/mms_plus.gif";
			document.getElementById(imgid).value = "0";
		}

		for (var i = 0 ; i < tr.length ; i++) {
			if (!tr[i].getAttribute("id")) continue;
			if (tr[i].getAttribute("id").indexOf(strtr) != -1){
				if (strimgvalue =="0"){
					document.getElementById(tr[i].getAttribute("id")).style.display ="";
				}else{
					document.getElementById(tr[i].getAttribute("id")).style.display ="none";
				}
			}
		}

		for (var i = 0 ; i < td.length ; i++) {
			if (!td[i].getAttribute("id")) continue;
			if (td[i].getAttribute("id").indexOf(strmidx1) != -1){
				if (td[i].className == "deps_2"){
					if (strimgvalue =="0"){
						document.getElementById(td[i].getAttribute("id")).style.display ="";
					}else{
						document.getElementById(td[i].getAttribute("id")).style.display ="none";
					}
				}
			}
		}
	}


	function get_page1(idx, str, p) {
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
		//document.getElementById("midx").value = idx ;
		//document.getElementById("ttitle").value = idx ;
	}




//-->


</SCRIPT>

