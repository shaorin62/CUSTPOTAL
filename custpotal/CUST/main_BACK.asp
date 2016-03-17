<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim gotopage : gotopage = request("gotopage")
	if gotopage = "" then gotopage = 1
	dim pagesize : pagesize = 10


	dim userid : userid = request.cookies("userid")

	dim sql : sql = "select custcode, class from wb_account where userid = ?"
	dim cmd : set cmd = server.createobject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adcmdText
	cmd.parameters.append cmd.createparameter("userid", adChar, adparaminput, 6)
	cmd.parameters("userid").value = userid

	dim rs : set rs = cmd.execute

	if rs.eof then
		response.redirect "/"
	else
		Response.Cookies("custcode2") = rs(0)
		response.Cookies("class") = rs(1)
	end if
	rs.close
	set rs = nothing
	set cmd = nothing

	response.write request.cookies("custcode2")
%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<link href="new.css" rel="stylesheet" type="text/css">
<title>▒SK MARKETING EXCELLENT▒</title>
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<body>
<table id="Table_01" width="1240" height="80" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td width="210" height="80" rowspan="2" valign="top"><a href="/cust/main.asp"><img src="/images/main_01.gif" width="210" height="240" alt="go main" border="0"></a></td>
		<td width="1030" height="35" valign="top" background="/images/top_02.gif" align="top">
		<table width="600" height="33" border="0" align="right" cellpadding="0" cellspacing="0">
      <tr>
        <td>&nbsp;</td>
        <td width="244" align="right"><span class="log">&nbsp;<%=request.cookies("custname")%></span> &nbsp;</td>
        <td width="1" valign="top"><img src="/images/top_vline_bg.gif" width="1" height="32"></td>
        <td width="104" align="right"><span class="log">&nbsp;<%=request.cookies("userid")%></span> &nbsp;</td>
        <td width="1" valign="top"><img src="/images/top_vline_bg.gif" width="1" height="32"></td>
        <td width="164" align="right"><span class="log"><%=request.cookies("logtime")%>&nbsp;&nbsp;</span></td>
        <td width="1" valign="top"><img src="/images/top_vline_bg.gif" width="1" height="32"></td>
        <td width="85" align="center"><A HREF="/Log_out.asp"><img src="/images/btn_logout.gif" width="64" height="19" border="0"></A></td>
      </tr>
    </table></td>
	</tr>
	<tr>
		<td valign="top" ><table width="1030" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="right" style="padding-right:300px;"><table width="1" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td><a href="/cust/trans/" target="_self" onClick="MM_nbGroup('down','group1','menu01','/images/top_menu_01_over.gif',1)" onMouseOver="MM_nbGroup('over','menu01','/images/top_menu_01_over.gif','/images/top_menu_01_over.gif',1)" onMouseOut="MM_nbGroup('out')"><img src="/images/top_menu_01.gif" alt="" name="menu01" width="99" height="40" border="0" onload=""></a></td>
        <td><img src="/images/top_dot_03.gif" alt="" name="blank01" width="44" height="40" border="0" onload=""></td>
        <td><a href="/cust/board/" target="_self" onClick="MM_nbGroup('down','group1','menu02','/images/top_menu_02_over.gif',1)" onMouseOver="MM_nbGroup('over','menu02','/images/top_menu_02_over.gif','/images/top_menu_02_over.gif',1)" onMouseOut="MM_nbGroup('out')"><img src="/images/top_menu_02.gif" alt="" name="menu02" width="113" height="40" border="0" onload=""></a></td>
        <td><img src="/images/top_dot_03.gif" alt="" name="blank02" width="44" height="40" border="0" onload=""></td>
        <td><a href="/cust/outdoor/contact_list.asp?menuNum=1" target="_self" onClick="MM_nbGroup('down','group1','menu03','/images/top_menu_03_over.gif',1)" onMouseOver="MM_nbGroup('over','menu03','/images/top_menu_03_over.gif','/images/top_menu_03_over.gif',1)" onMouseOut="MM_nbGroup('out')"><img src="/images/top_menu_03.gif" alt="" name="menu03" width="84" height="40" border="0" onload=""></a></td>
        <td><% if request.cookies("class") = "G" then %><img src="/images/top_dot_03.gif" alt="" name="blank02" width="44" height="40" border="0" onload=""><% end if%></td>
        <td><% if request.cookies("class") = "G" then %><a href="/cust/admin/" target="_self" onClick="MM_nbGroup('down','group1','menu04','/images/top_menu_04_over.gif',1);" onMouseOver="MM_nbGroup('over','menu04','/images/top_menu_04_over.gif','/images/top_menu_04_over.gif',1)" onMouseOut="MM_nbGroup('out')"><img src="/images/top_menu_04.gif" alt="" name="menu04" width="100" height="40" border="0" onload=""></a><% end if%></td>
      </tr>
    </table></td>
  </tr>
</table>
		  <script type="text/javascript">
AC_FL_RunContent('codebase','http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0','width','1036','height','155','src','/images/bn','quality','high','pluginspage','http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash','movie','/images/bn' ); //end AC code
</script><noscript><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" width="1036" height="155">
            <param name="movie" value="/images/bn.swf" />
            <param name="quality" value="high" />
            <embed src="/images/bn.swf" quality="high" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="1036" height="155"></embed>
      </object></noscript></td>
</tr>
	<tr>
	  <td height="80"  valign="top"> </td>
	  <td height="568" valign="top">
<!--  --><table width="98%" border="0" align="left" cellpadding="0" cellspacing="0" >

  <tr align="left">
    <td colspan="5" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <!-- ##################  레코드 출력 시작 #####################-->
        <%
			dim objrs
			dim ridx, midx, title, c_user, c_date, totalcount, totalpage, trgbcolor
			sql = "select midx from dbo.wb_menu_mst where custcode is null and mp = 0 "
			call get_recordset(objrs, sql)

			if not objrs.eof then
				midx = objrs(0)
			else
				midx = 0
			end if

			sql = "select count(ridx) from dbo.wb_report r inner join dbo.wb_menu_mst m on r.midx = m.midx where custcode is null and mp=0"

			call get_recordset(objrs, sql)
			if not objrs.eof then
				totalcount = objrs(0)
			else
				totalcount = 0
			end if
			objrs.close

			sql = "select top " & pagesize & "  r.ridx, r.title, r.cuser, r.cdate, r.midx from dbo.wb_report r inner join dbo.wb_menu_mst m on r.midx = m.midx where ridx not in (select top " & pagesize*(gotopage-1) &" r.ridx from dbo.wb_report r inner join dbo.wb_menu_mst m on r.midx = m.midx where custcode is null and mp = 0 order by ridx desc) and custcode is null and mp = 0 order by ridx desc "

			call get_recordset(objrs, sql)

			if not objrs.eof then
				set ridx = objrs("ridx")
				set title = objrs("title")
				set c_user = objrs("cuser")
				set c_date = objrs("cdate")
				set midx = objrs("midx")
				totalpage = int((totalcount-1)/pagesize)+1
			else
				totalpage = 0
			end if
		%>
        <tr>
          <td width="26%" height="20"><font color="#B9B9B9">Total Count :&nbsp;<font color="#FF6600"><%=totalcount%></font></font></td>
          <td width="74%" height="20" align="right"><font color="#B9B9B9">Total Page :<font color="#FF6600">
            <%=totalpage%></font></font>&nbsp;&nbsp;</td>
        </tr>
      </table></td>
  </tr>
   <tr>
	<td colspan=4 height="2" bgcolor="#B9B9B9"></td>
  </tr>
  <tr height=30>
    <td align="center" width="40">번호</td>
    <td align="center" width = "700">제목</td>
    <td align="center" width="100">작성자</td>
    <td align="center"  width="120">작성일자</td>
  </tr>
  <tr>
	<td colspan=4 height="1" bgcolor="#B9B9B9"></td>
  </tr>
  <%
		dim intLoop : intLoop = totalcount - ((gotopage-1)*pagesize)
		do until objrs.eof
		if intLoop mod 2 = "0" then
		 trgbcolor = "#F6F6F6"
        else
		 trgbcolor = "#FFFFFF"
        end if
	%>
  <!--####################   실제 데이터 출력     #####################-->
  <tr onMouseOver="this.style.backgroundColor='<%=trgbcolor%>'" onMouseOut="this.style.backgroundColor='#FFFFFF'">
    <td  height="25" align="center" ><font color="#46676F"><%=intLoop%></font></td>
    <td  height="25" align="left" style="padding-left:20px;">
	<span onClick='pop_report_view(<%=ridx%>,<%=midx%>);' class='styleLink'><% if len(title) > 50 Then response.write Mid(title,1,50) + "..." Else response.write title End If%></span>
	<%
		dim objrs2, objrs3
		sql = "select idx, ridx, attachfile from dbo.wb_Report_pds where ridx = " & ridx
		call get_Recordset(objrs2, sql)
		dim count : count = 1
		if not objrs2.eof then
		do until objrs2.eof
		if count < 4 then
	%>
	<span onClick="checkForDownload('<%=objrs2("attachfile")%>');" class="styleLink" style="padding-left:10px;"> <img src="/images/ico_attach.gif" width="7" height="12" hspace="3"> <%=objrs2("attachfile")%></span>
	<%
		else
		response.write " ..."
		exit do
		end if
		count = count + 1
		objrs2.movenext
		Loop

		end if
		objrs2.close
		set objrs2 = nothing
	%>
	<!--   '등록일자를 24시간을 비교해서 작으면 뉴이미지를 찍는다(최신글 리스트)  -->
      <% if datediff("d",c_date,now()) = 1 then %> <img src="/images/i_new.gif" width="12" height="12" align="absmiddle" ><% end if %>
    </td>
    <td  height="25" align="center"><%=c_user%></td>
    <td height="25" align="center" ><%=c_date%></td>
  </tr>
  <tr>
    <td height="1" colspan="4" background="/Image/line2.gif"></td>
  </tr>
  <%
	sql = "select cidx, ridx, comment, attachfile, cuser, cdate from dbo.wb_report_comment where ridx="& ridx
	call get_recordset(objrs3, sql)

	if not objrs3.eof then
		do until objrs3.eof
	%>
			  <tr ><!-- 댓글목록 -->
				<td align="center" height="31">&nbsp;</td>
				<td align="left" style="padding-left:20px;"><img src="/images/icon_reply.gif" width="28" height="16" hspace="3" align="bottom"  ><span onClick="pop_report_view(<%=ridx%>,<%=midx%>);" class="styleLink"><%=objrs3("comment")%></span> <% if not isnull(objrs3("attachfile")) then %><span onClick="checkForDownload('<%=objrs3("attachfile")%>');" class="styleLink" style="padding-left:50px;font-size:11px;font-fmaily:돋움"> <img src="/images/ico_attach.gif" width="7" height="12" hspace="3"><%if len(objrs3("attachfile")) > 40 then response.write left(objrs3("attachfile"), 37)&"..." else response.write objrs3("attachfile")%></span><% end if%> </td>
				<td  align="center"><%=objrs3("cuser")%></td>
				<td  align="center"><%=formatdatetime(objrs3("cdate"),2)%></td>
			  </tr>
	<%
		objrs3.movenext
		Loop
	end if
		intLoop = intLoop - 1
	objrs.movenext ''다음 레코드로 이동
	loop        ''루프를 반복한다
	%>
  <tr>
    <td height="2" colspan="4" bgcolor="#B9B9B9"></td>
  </tr>
  <tr align="right" height=30>
    <td height="21" colspan="4" align="right">

  <%  %>
      <!-- 페이징 파일 -->
<%

				   Dim cdivide,blockpage,x
                   cdivide = 10 '페이지 나누는 개수
                   'response.write "cdivide='"&cdivide&"'<br>"

                  blockPage=Int((gotopage - 1) / cdivide ) * cdivide + 1
        '************************ 이전 10 개구문 시작 ***************************
                if blockPage = 1 Then
                   Response.Write ""
                Else
                %>
                <a href="main.asp?page=<%=blockPage-cdivide%>" class="pagesplit"> <img src="/images/i_pp.gif" align="absmiddle"  border="0"></a>
                <%
                End If
        '************************ 이전 10 개 구문 끝***************************

        '---이전으로 가기-------------------------------------------------------
               if gotopage=1 and int(gotopage)<>int(totalpage) then
              %>
                <img src="/images/i_pre.gif"  border="0" align=absmiddle >
                <% elseif gotopage=1 and int(gotopage)=int(totalpage) then %>
                <img src="/images/i_pre.gif"  border="0" align=absmiddle > <!--width="16" height="12"-->
                <% elseif int(gotopage)=int(totalpage) then %>
                <a href="main.asp?page=<%=gotopage - 1%>" class="pagesplit"> <img src="/images/i_pre.gif" align=absmiddle  border="0"></a>
                <% else %>
                <a href="main.asp?page=<%=gotopage - 1%>" class="pagesplit"> <img src="/images/i_pre.gif" align=absmiddle  border="0"></a>
                <% end if
       '---이전으로 가기 끝---------------------------------------------------


             x=1

	         Do Until x > cdivide or blockPage > totalpage
             If blockPage=int(gotopage) Then
             %>
                <font color="#FF9900"><%=blockPage%></font>
                <%Else%>
                <a href="main.asp?page=<%=blockPage%>" class="pagesplit"><%=blockPage%></a>
                <%
    End If

    blockPage=blockPage+1
    x = x + 1
    Loop


'----다음으로 가기---------------------------------------------------
if gotopage=1 and int(gotopage)<>int(totalpage) then
%>
                <a href="main.asp?page=<%=gotopage+1%>"><img src="/images/i_next.gif" align=absmiddle  border="0" class="pagesplit"></a>
                <%elseif gotopage=1 and int(gotopage)=int(totalpage) then%>
                <img src="/images/i_next.gif"  border="0" align=absmiddle >
                <%elseif int(gotopage)=int(totalpage) then%>
                <img src="/images/i_next.gif" border="0" align=absmiddle >
                <%else%>
                <a href="main.asp?page=<%=gotopage+1%>"> <img src="/images/i_next.gif" align=absmiddle  border="0" class="pagesplit"></a>
                <%end if
'-----다음으로 가기 끝-------------------------------------------------

'************************ 다음 10 개 구문 시작***************************
if blockPage > totalpage Then
   Response.Write ""
Else
%>
                <a href="main.asp?page=<%=blockPage%>"> <img src="/images/i_ff.gif" align=absmiddle  border="0" class="pagesplit"></a>
                <%
End If
'************************ 다음 10 개 구문 끝***************************
%>
      <!-- 페이징 파일  -->    </td>
  </tr>
  <tr>
    <td height="1" colspan="4" align="center" bgcolor="#B9B9B9"></td>
  </tr>
</table>
<!--  -->
</td>
</tr>
  <tr>
    <td colspan="2"><!--#include virtual="/bottom.asp" --></td>
  </tr>
</table>
</body>
<script language="JavaScript">
<!--

	function MM_preloadImages() { //v3.0
	  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
		var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
		if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
	}

	function MM_findObj(n, d) { //v4.01
	  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
		d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
	  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
	  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
	  if(!x && d.getElementById) x=d.getElementById(n); return x;
	}

	function MM_nbGroup(event, grpName) { //v6.0
	  var i,img,nbArr,args=MM_nbGroup.arguments;
	  if (event == "init" && args.length > 2) {
		if ((img = MM_findObj(args[2])) != null && !img.MM_init) {
		  img.MM_init = true; img.MM_up = args[3]; img.MM_dn = img.src;
		  if ((nbArr = document[grpName]) == null) nbArr = document[grpName] = new Array();
		  nbArr[nbArr.length] = img;
		  for (i=4; i < args.length-1; i+=2) if ((img = MM_findObj(args[i])) != null) {
			if (!img.MM_up) img.MM_up = img.src;
			img.src = img.MM_dn = args[i+1];
			nbArr[nbArr.length] = img;
		} }
	  } else if (event == "over") {
		document.MM_nbOver = nbArr = new Array();
		for (i=1; i < args.length-1; i+=3) if ((img = MM_findObj(args[i])) != null) {
		  if (!img.MM_up) img.MM_up = img.src;
		  img.src = (img.MM_dn && args[i+2]) ? args[i+2] : ((args[i+1])? args[i+1] : img.MM_up);
		  nbArr[nbArr.length] = img;
		}
	  } else if (event == "out" ) {
		for (i=0; i < document.MM_nbOver.length; i++) {
		  img = document.MM_nbOver[i]; img.src = (img.MM_dn) ? img.MM_dn : img.MM_up; }
	  } else if (event == "down") {
		nbArr = document[grpName];
		if (nbArr)
		  for (i=0; i < nbArr.length; i++) { img=nbArr[i]; img.src = img.MM_up; img.MM_dn = 0; }
		document[grpName] = nbArr = new Array();
		for (i=2; i < args.length-1; i+=2) if ((img = MM_findObj(args[i])) != null) {
		  if (!img.MM_up) img.MM_up = img.src;
		  img.src = img.MM_dn = (args[i+1])? args[i+1] : img.MM_up;
		  nbArr[nbArr.length] = img;
	  } }
	}
	MM_preloadImages('top_menu_01_over.gif','top_dot_03.gif','top_menu_02_over.gif','top_menu_03_over.gif','top_menu_04_over.gif');


	function pop_report_view(ridx, midx) {
		var url = "/cust/board/pop_report_view.asp?ridx="+ridx+"&midx="+midx+"&read=y";
		var name = "pop_report_view" ;
		var opt = "width=658, height=680, resizable=no, scrollbars=yes, top=100, left=100";
		window.open(url, name, opt);
	}

	function checkForDownload(name) {
		location.href="/cust/board/download.asp?filename="+name;
	}


//-->
</script>