<table id="Table_01" width="1240" height="88" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td width="210" height="88" rowspan="2" valign="top">
			<a href="/main.asp"><img src="/images/top_01.gif" width="210" height="88" alt="go main" border="0"></a></td>
		<td width="1030" height="40" valign="top" background="/images/top_02.gif">
		<table width="600" height="33" border="0" align="right" cellpadding="0" cellspacing="0">
      <tr>
        <td>&nbsp;</td>
        <td width="244" align="right"><span class="log">&nbsp;<%=request.cookies("classname")%></span> &nbsp;</td>
        <td width="1" valign="top"><img src="/images/top_vline_bg.gif" width="1" height="32"></td>
        <td width="104" align="right"><span class="log">&nbsp;1212<%=request.cookies("userid")%></span> &nbsp;</td>
        <td width="1" valign="top"><img src="/images/top_vline_bg.gif" width="1" height="32"></td>
        <td width="164" align="right"><span class="log"><%=formatdatetime(request.cookies("logtime"))%>&nbsp;&nbsp;</span></td>
        <td width="1" valign="top"><img src="/images/top_vline_bg.gif" width="1" height="32"></td>
        <td width="85" align="center"><A HREF="/Log_out.asp"><img src="/images/btn_logout.gif" width="64" height="19" border="0"></A></td>
      </tr>
    </table>
	</td>
	</tr>
	<tr>
		<td ><table width="1030" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="right" style="padding-right:280px;"><table width="1" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td><a href="/hq/trans/public_01_list.asp" target="_self" onClick="MM_nbGroup('down','group1','menu01','/images/top_menu_01_over.gif',1)" onMouseOver="MM_nbGroup('over','menu01','/images/top_menu_01_over.gif','/images/top_menu_01_over.gif',1)" onMouseOut="MM_nbGroup('out')"><img src="/images/top_menu_01.gif" alt="" name="menu01" width="99" height="50" border="0" onload=""></a></td>
        <td><a href="javascript:;" target="_self" onClick="MM_nbGroup('down','group1','blank01','/images/top_dot_03.gif',1)" onMouseOver="MM_nbGroup('over','blank01','/images/top_dot_03.gif','/images/top_dot_03.gif',1)" onMouseOut="MM_nbGroup('out')"><img src="/images/top_dot_03.gif" alt="" name="blank01" width="44" height="50" border="0" onload=""></a></td>
        <td><a href="/hq/board/list.asp" target="_self" onClick="MM_nbGroup('down','group1','menu02','/images/top_menu_02_over.gif',1)" onMouseOver="MM_nbGroup('over','menu02','/images/top_menu_02_over.gif','/images/top_menu_02_over.gif',1)" onMouseOut="MM_nbGroup('out')"><img src="/images/top_menu_02.gif" alt="" name="menu02" width="113" height="50" border="0" onload=""></a></td>
        <td><a href="javascript:;" target="_self" onClick="MM_nbGroup('down','group1','blank02','/images/top_dot_03.gif',1)" onMouseOver="MM_nbGroup('over','blank02','/images/top_dot_03.gif','/images/top_dot_03.gif',1)" onMouseOut="MM_nbGroup('out')"><img src="/images/top_dot_03.gif" alt="" name="blank02" width="44" height="50" border="0" onload=""></a></td>
        <td><a href="/hq/outdoor/contact_list.asp?menuNum=1" target="_self" onClick="MM_nbGroup('down','group1','menu03','/images/top_menu_03_over.gif',1)" onMouseOver="MM_nbGroup('over','menu03','/images/top_menu_03_over.gif','/images/top_menu_03_over.gif',1)" onMouseOut="MM_nbGroup('out')"><img src="/images/top_menu_03.gif" alt="" name="menu03" width="84" height="50" border="0" onload=""></a></td>
        <td><a href="javascript:;" target="_self" onClick="MM_nbGroup('down','group1','blank03','/images/top_dot_03.gif',1)" onMouseOver="MM_nbGroup('over','blank03','/images/top_dot_03.gif','/images/top_dot_03.gif',1)" onMouseOut="MM_nbGroup('out')"></a></td>
        <td><% If request.cookies("class") = "A"  Or request.cookies("class") = "N"  Then %><img src="/images/top_dot_03.gif" alt="" name="blank03" width="44" height="50" border="0" onload=""><a href="/hq/admin/acc_list.asp?acc_menu=A00000" target="_self" onClick="MM_nbGroup('down','group1','menu04','/images/top_menu_04_over.gif',1);" onMouseOver="MM_nbGroup('over','menu04','/images/top_menu_04_over.gif','/images/top_menu_04_over.gif',1)" onMouseOut="MM_nbGroup('out')"><img src="/images/top_menu_04.gif" alt="" name="menu04" width="100" height="50" border="0" onload=""></a><% End if%></td>
      </tr>
    </table></td>
  </tr>
</table>
	</td>
</tr>
</table>
<script type="text/javascript">
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
//-->
</script>