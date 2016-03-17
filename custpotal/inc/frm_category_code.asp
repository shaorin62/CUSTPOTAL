<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim ggroupidx : ggroupidx = request("ggroupidx")
	if ggroupidx = "" then ggroupidx = null
	dim mgroupidx : mgroupidx = request("mgroupidx")
	if mgroupidx = "" then mgroupidx = null
	dim sgroupidx : sgroupidx = request("sgroupidx")
	if sgroupidx = "" then sgroupidx = null
	dim dgroupidx : dgroupidx = request("dgroupidx")
	if dgroupidx = "" then dgroupidx = null

	dim rs, sql, gstr, mstr, sstr, dstr
	if not isnull(ggroupidx) then
		sql = "select categoryidx, categoryname from dbo.wb_category where categorylvl  = 1 and highcategoryidx = "&ggroupidx&" order by categoryidx "
		response.write sql
		call get_recordset(rs, sql)

		mstr = "<select name='selmcategory' style='width:320px;' onchange = 'get_category_middle()' > <option value = ''> 중분류를 선택하세요. </option>"
		do until rs.eof
			mstr = mstr & "<option value ='" & rs("categoryidx") & "' "
				if not isnull(mgroupidx) then
					if cint(mgroupidx) = cint(rs("categoryidx")) then mstr = mstr & " selected "
				end if
			mstr = mstr & ">" & rs("categoryname") & " </option>"
		rs.movenext
		loop
		mstr = mstr & "</select>"
		rs.close
	end if

	if not isnull(mgroupidx) then
		sql = "select categoryidx, categoryname from dbo.wb_category where categorylvl  = 2 and highcategoryidx = "&mgroupidx&" order by categoryidx "
		response.write sql
		call get_recordset(rs, sql)

		sstr = "<select name='selscategory' style='width:320px;' onchange = 'get_category_small()' > <option value = ''> 소분류를 선택하세요. </option>"
		do until rs.eof
			sstr = sstr & "<option value ='" & rs("categoryidx") & "' "
				if not isnull(sgroupidx) then
					if cint(sgroupidx) = cint(rs("categoryidx")) then sstr = sstr & " selected "
				end if
			sstr = sstr & ">" & rs("categoryname") & " </option>"
		rs.movenext
		loop
		sstr = sstr & "</select>"
		rs.close
	end if

	if not isnull(sgroupidx) then
		sql = "select categoryidx, categoryname from dbo.wb_category where categorylvl  = 3 and highcategoryidx = "&sgroupidx&" order by categoryidx "
		response.write sql
		call get_recordset(rs, sql)

		dstr = "<select name='seldcategory' style='width:320px;' onchange = 'get_category_detail()' > <option value = ''> 세분류를 선택하세요. </option>"
		do until rs.eof
			dstr = dstr & "<option value ='" & rs("categoryidx") & "' "
				if not isnull(dgroupidx) then
					if cint(dgroupidx) = cint(rs("categoryidx")) then dstr = dstr & " selected "
				end if
			dstr = dstr & ">" & rs("categoryname") & " </option>"
		rs.movenext
		loop
		dstr = dstr & "</select>"
		rs.close
	end if

	if not isnull(ggroupidx) then response.write mstr
	if not isnull(mgroupidx) then response.write sstr
	if not isnull(sgroupidx) then response.write dstr


%>
<SCRIPT LANGUAGE="JavaScript">
<!--
<% if not isnull(ggroupidx) then %> parent.document.getElementById("mgroup").innerHTML = "<%=mstr%>"  ;<% end if %>
<% if not isnull(mgroupidx) then %> parent.document.getElementById("sgroup").innerHTML = "<%=sstr%>"  ;<% end if %>
<% if not isnull(sgroupidx) then %>  parent.document.getElementById("dgroup").innerHTML = "<%=dstr%>"  ;<% end if %>


//-->
</SCRIPT>