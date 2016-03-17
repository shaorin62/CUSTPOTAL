<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim gotopage : gotopage = request("gotopage")
	dim searchstring : searchstring = request("searchstring")

	dim mdidx : mdidx = request("mdidx")
	'response.write mdidx
	dim sql : sql = "select  m.mdidx, m.title, m.custcode, m.categoryidx, m.unit, m.region, m.locate, m.map, m.mediummemo, m.regionmemo,  m.cuser, m.cdate, m.uuser, m.udate , v.mgroupidx, c.tidx from dbo.WB_MEDIUM_MST M inner join dbo.vw_medium_category v on m.categoryidx = v.mdidx  left outer  join dbo.wb_validation_class c on m.mdidx = c.mdidx and c.isuse = 1 where  m.mdidx = " & mdidx

	dim objrs
	call get_recordset(objrs, sql)

	if objrs.eof then response.write "<script> alert('삭제 또는 잘못된 매체정보 입니다.'); location.href='medium_list.asp?gotopage=" & gotopage & "&serarchstring=" & searchstring

	dim title : title = objrs.fields("title")
	dim custcode : custcode = objrs.fields("custcode")
	dim categoryidx : categoryidx = objrs.fields("categoryidx")
	dim unit : unit = objrs.fields("unit")
	dim region : region = objrs.fields("region")
	dim locate : locate = objrs.fields("locate")
	dim map : map = objrs.fields("map")
	dim mediummemo : mediummemo = objrs.fields("mediummemo")
	dim regionmemo : regionmemo = objrs.fields("regionmemo")
	dim mgroupidx : mgroupidx = objrs.fields("mgroupidx")
	dim tidx : tidx = objrs.fields("tidx")

	objrs.close

	sql = " select sidx, mdidx, side, standard, quality, unitprice from dbo.WB_MEDIUM_DTL where mdidx = " & mdidx
	call get_recordset(objrs, sql)

	dim sidx,  side, standard, quality, unitprice, flag
	flag = true

	if not objrs.eof then
		set sidx = objrs("sidx")
		set side = objrs("side")
		set standard = objrs("standard")
		set quality = objrs("quality")
		set unitprice = objrs("unitprice")
		flag = false
	end if

%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  oncontextmenu="return false">
<form>
<input type="hidden" name="txtmap" value="<%=map%>">
<!--#include virtual="/od/top.asp" -->
  <table id="Table_01" width="1240"  border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/od/left_outdoor_menu.asp"--></td>
      <td height="65"><img src="/images/middle_navigater_outdoor.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td height="400" align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="19">&nbsp;</td>
          </tr>
          <tr>
            <td height="17"><TABLE  width="100%">
            <TR>
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle"> <%=title%> </span></TD>
				<TD width="50%" align="right"><span class="navigator" > 옥외관리 &gt; 매체관리 &gt; <%=title%> </span></TD>
            </TR>
            </TABLE>
			</td>
          </tr>
          <tr>
            <td height="15">&nbsp;</td>
          </tr>
          <tr>
            <td valign="top" class="bdpdd">
			<table border="0" cellpadding="0" cellspacing="0" >
			<tr>
				<td colspan="2" bgcolor="#cacaca" height="1"></td>
			</tr>
			<tr>
				<td class="hw">매체명</td>
				<td class="bw bd"><%=title%></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr >
				<td class="hw">매체사</td>
				<td class="bw bd"> <% call get_medium_custcode(custcode, "r")%>	</td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr >
				<td class="hw">매체분류</td>
				<td class="bw bd"><span id="category"><% call get_medium_catetory(categoryidx)%></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hw">수량단위</td>
				<td class="bw bd"><%=unit%></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr >
				<td class="hw">설치지역</td>
				<td class="bw bd"><%call get_region_code(region, "r") %> </td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr >
				<td class="hw">위치정보</td>
				<td class="bw bd"><%=locate%></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<%	do until objrs.eof	%>
			<tr >
				<td class="hw">면별정보</td>
				<td class="bw bd stylelink" onclick="pop_side_view('<%=sidx%>');"><span class="side" alt="매체 면 위치"><%=trim(side)%></span> <span class="standard" alt="면별 규격정보"><%=trim(standard)%></span> <span class="quality" alt="면별 재질정보"><%=trim(quality)%></span> <span class="unitprice" alt="면별 단위가격"><%=formatnumber(unitprice,0)%> (단위:원)</span></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<%
					objrs.movenext
				loop
			%>
			<tr >
				<td class="hw">매체약도</td>
				<td class="bw bd"><% if not isnull(map) then %><img src="/pds/media/<%=map%>" width="500" height="350" border="1" alt="" vspace="11"><%end if%></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			</table>
			<table width="1002" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="50%" height="50" valign="bottom"><!-- <a href="/hq/outdoor/medium_list.asp?gotopage=<%=gotopage%>&txtsearchstring=<%=searchstring%>&selcategory=<%=categoryidx%>&selcustcode=<%=custcode%>"><img src="/images/btn_list.gif" width="59" height="18" border="0"></a> --><img src="/images/btn_list.gif" width="59" height="18" border="0" class="stylelink" onclick="history.back();"></td>
                  <td width="50%" align="right" valign="bottom"><% if mgroupidx = 9 or mgroupidx = 10 or mgroupidx = 11 then %><img src="/images/btn_md_validation.gif" width="78" height="18" vspace="5" hspace="10" border="0" class="stylelink" onClick="pop_medium_validation(<%=mgroupidx%>);"><% end if %><img src="/images/btn_side_reg.gif" width="86" height="18" vspace="5" border="0" class="stylelink" onClick="pop_side_add();"><img src="/images/btn_edit.gif" width="59" height="18" hspace="10" vspace="5" border="0" class="stylelink" onClick="pop_medium_edit();"><img src="/images/btn_delete.gif" width="59" height="18" vspace="5" border="0" class="stylelink" onClick="go_medium_delete();"></td>
                </tr>
              </table></td>
          </tr>
      </table>	  </td>
    </tr>
  </table>
<!--#include virtual="bottom.asp" -->
</form>
</body>
</html>
<script language="javascript">
<!--
	function pop_medium_edit() {
		if (confirm("매체정보를 수정하시겠습니까?")) {
			var url = "pop_medium_edit.asp?mdidx=<%=mdidx%>"
			var name = "pop_medium_edit";
			var opt = "width=540, height=359, resizable=no, scrollbars=no, status=yes, left=100, top=100";
			window.open(url, name, opt);
		}
	}

	function go_medium_delete() {
		if (confirm("시스템에서 매체정보가 삭제됩니다.\n\n매체정보를 삭제하시겠습니까?")) {
			location.href = "medium_delete_proc.asp?mdidx=<%=mdidx%>&gotopage=<%=gotopage%>&searchstring=<%=searchstring%>";
		}
	}

	function pop_side_add() {
		var url = "pop_side_add.asp?mdidx=<%=mdidx%>&title=<%=title%>";
		var name = "pop_side_add";
		var opt = "width=540, height=266, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		window.open(url, name, opt);
	}

	function pop_side_view(sidx) {
		var url = "pop_side_view.asp?sidx="+sidx+"&title=<%=title%>";
		var name = "pop_side_add";
		var opt = "width=540, height=266, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		window.open(url, name, opt);
	}

	function pop_medium_validation(idx) {
		var url ;
		if (idx == 9)			url = "validation_led.asp?tidx=<%=tidx%>&mdidx=<%=mdidx%>";						//LED
		else if (idx == 10)	url = "validation_neon.asp?tidx=<%=tidx%>&mdidx=<%=mdidx%>";			//옥탑
		else if (idx == 11)	url = "validation_board.asp?tidx=<%=tidx%>&mdidx=<%=mdidx%>";		//야립
		else url = "validation_etc.asp?tidx=<%=tidx%>&mdidx=<%=mdidx%>";
		var name = "pop_medium_validation";
		var opt = "width=838, height=800, resizable=no, scrollbars=yes, status=yes, top=100, left=100";
		window.open(url, name, opt);
	}

	window.onload = function () {
	<% if flag then %>

		if (confirm("면별 정보가 등록되지 않은 매체정보 입니다.\n\n면별 정보를 등록하시겠습니까?")) {
			pop_side_add();
			return false;
		}
	<% end if %>
	}
//-->
</script>
<%
	objrs.close
	set objrs = nothing
%>