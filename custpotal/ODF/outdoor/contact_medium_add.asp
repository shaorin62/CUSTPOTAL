<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim contidx : contidx = request("contidx")
	dim searchstring : searchstring = request("txtsearchstring")
	dim searchcategory : searchcategory = request("selcategory")
	if searchcategory = "" then searchcategory = 0

	dim objrs, sql
	sql = "select title, highcustcode from dbo.wb_contact_mst m inner join dbo.sc_cust_temp c on m.custcode = c.custcode where contidx = " & contidx &"; select d.sidx, m.title, d.side, d.standard, d.quality, m.locate " &_
			" from dbo.wb_medium_mst m inner join dbo.wb_medium_dtl d on m.mdidx = d.mdidx " &_
			" inner join dbo.vw_medium_category c on m.categoryidx = c.mdidx  " &_
			" where m.title like '%" & searchstring &"%' or c.mgroupidx = " & searchcategory
	call get_recordset(objrs, sql)
	dim contacttitle, highcustcode
	contacttitle = objrs("title")
	highcustcode = objrs("highcustcode")
	set objrs = objrs.NextRecordset
	dim sidx, title, side, standard, quality, locate
	if not objrs.eof then
		set sidx = objrs("sidx")
		set title = objrs("title")
		set side = objrs("side")
		set standard = objrs("standard")
		set quality = objrs("quality")
		set locate = objrs("locate")
	end if
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
</head>

<body leftmargin="0" topmargin="0"  oncontextmenu="return false">
<form>
  <table border="0" cellpadding="1" cellspacing="0" align="center">
	<tr>
		<td colspan="2">
			<table  height="35" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
                  <td width="50%" align="left" background="/images/bg_search.gif">&nbsp;<span class="header"><%=contacttitle%></span></td>
                  <td width="50%" align="right" background="/images/bg_search.gif">&nbsp;</td>
                  <td width="13"><img src="../../images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table>
		</td>
	</tr>
    <tr>
      <td valign="top"><table cellpadding="0" cellspacing="0">
          <tr >
            <td class="thd2" width="200" style="padding-left:5px;">매체명</td>
            <td class="thd2" width="30">면</td>
            <td class="thd2" width="170">규격(M)/재질</td>
            <td class="thd2" width="300">설치위치</td>
          </tr>
          <tr>
            <td colspan="4">
			<table width="700" border="0" border="0" cellpadding="0" cellspacing="0">
			  <% do until objrs.eof %>
              <tr height="30" onclick="get_medium_data(<%=sidx%>)" class="stylelink">
                <td width="200" style="padding-left:5px;"><%=title%>&nbsp;</td>
                <td width="30"><%=side%>&nbsp;</td>
                <td width="170"><%=standard%> <% if not isnull(quality) then response.write " (" &quality &")"%> &nbsp;</td>
                <td width="300"><%=locate%>&nbsp;</td>
              </tr>
			  <tr>
				<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
			  </tr>
			  <%
					objrs.movenext
					loop
					objrs.close

					dim sidx_ : sidx_ = request("sidx_")
					if sidx_ = "" then sidx_ = 0
					sql = "select d.sidx, m.title, c.mdname as categoryname, d.side, m.unit, d.standard, d.quality , m.custcode" &_
							" from dbo.wb_medium_mst m inner join dbo.wb_medium_dtl d on m.mdidx = d.mdidx" &_
							" inner join dbo.vw_medium_category c on m.categoryidx = c.mdidx" &_
							" where d.sidx = " & sidx_
					call get_recordset(objrs, sql)
					dim  title_, categoryname_, side_, unit_, standard_, quality_, custcode_
					if not objrs.eof then
						title_ = objrs("title")
						categoryname_ = objrs("categoryname")
						side_ = objrs("side")
						unit_ = objrs("unit")
						standard_ = objrs("standard")
						quality_ = objrs("quality")
						custcode_ = objrs("custcode")
					end if
					objrs.close
					set objrs = nothing
			  %>
            </table>
			</td>
          </tr>

      </table></td>
      <td valign="top">
	  <table border="0" cellpadding="0" cellspacing="0">
		  <tr>
			<td colspan="4" bgcolor="#cacaca" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">매체명</td>
            <td colspan="3" class="tdbd s7"><%=title_%><input type="hidden" name="sidx" value="<%=sidx_%>">&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">분류</td>
            <td  class="thbd s5"><%=categoryname_%>&nbsp;</td>
            <td  class="tdhd s4">면</td>
            <td  class="tdbd s6"><%=side_%>&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">수량</td>
            <td class="thbd s5"><input name="txtqty" type="text" size="5" value="1" class="number"></td>
            <td class="tdhd s4">단위</td>
            <td class="thbd s6"><%=unit_%>&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">규격</td>
            <td class="thbd s5"><%=standard_%>&nbsp;</td>
            <td class="tdhd s4">재질</td>
            <td class="thbd s6"><%=quality_%>&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">등급</td>
            <td colspan="3" class="thbd s7"><input name="rdotrust" type="radio" value="일반" checked>
              일반
              <input name="rdotrust" type="radio" value="정책">
              정책</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">월청구액</td>
            <td colspan="3"  class="tdbd s7"><input type="text" name="txtmonthprice"  class="number"  onfocus="initNum(this);" onblur="initZero(this);" value="0">&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">월지급구액</td>
            <td colspan="3"  class="tdbd s7"><input type="text" name="txtexpense"  class="number"  onfocus="initNum(this);" onblur="initZero(this);" value="0">&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">매체사</td>
            <td colspan="3"  class="tdbd s7"><%call get_medium_custcode(custcode_, "R")%>&nbsp;</td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">소재명</td>
            <td colspan="3"  class="tdbd s7"><%call get_jobcust_subject(highcustcode, null, null) %><%=highcustcode%></td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">시작일</td>
            <td class="thbd s5"><input type="text" name="txtstartdate" style="width:80;"> <img src="/images/calendar.gif" width="24" height="22" align="absmiddle"  onclick="Calendar_D(document.all.txtstartdate)"></td>
            <td class="tdhd s4">종료일</td>
            <td class="thbd s6"><input type="text" name="txtenddate" style="width:80;"> <img src="/images/calendar.gif" width="24" height="22" align="absmiddle"   onclick="Calendar_D(document.all.txtenddate)"></td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">특이사항</td>
            <td colspan="3" style="padding-top:5px;padding-bottom:5px;"  class="tdbd s7"><textarea name="txtcomment" rows="5" style="width:380px;"></textarea></td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
      </table>
	  </td>
    </tr>
	<tr>
		<td align="right" colspan="2" height="50" valign="bottom"><img src="/images/btn_save.gif" width="59" height="20"  vspace="5" style="cursor:hand"  hspace="10" onClick="check_submit();"><img src="/images/btn_close.gif" width="57" height="18" vspace="6" style="cursor:hand" onClick="set_close();"><img src="/images/btn_close.gif" width="59" height="20" hspace="10" vspace="6" style="cursor:hand" onClick="set_close();"></td>
	</tr>
  </table>
  <input type="hidden" name="contidx" value=<%=contidx%>>
</form>
</body>
</html>
<script type="text/javascript" src="/js/calendar.js"></script>
<script type="text/javascript" src="/js/script.js"></script>
<script language="JavaScript">
<!--
	window.onload = function () {
		self.focus();
	}

	function get_medium_data(idx) {
		location.href = "contact_medium_add.asp?sidx_="+idx+"&searchstring=<%=searchstring%>&searchcategory=<%=searchcategory%>&contidx=<%=contidx%>";
	}

	function check_submit() {
		var frm = document.forms[0];
		if (frm.sidx.value == "") {
			alert("등록할 매체를 선택하세요");
			return false;
		}
		frm.action = "contact_medium_reg_proc.asp";
		frm.method = "post";
		frm.submit();
	}
//-->
</script>
