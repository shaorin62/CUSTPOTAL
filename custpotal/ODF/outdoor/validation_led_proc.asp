<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim mdidx : mdidx = request("mdidx")
	dim tidx : tidx = request("tidx")
	if tidx = "" then tidx = 0
	dim mclass : mclass = request("txtclass")
	dim avg : avg = request("txtavg")
	dim v_1_1 : v_1_1 = cint(request("sel1_1"))
	dim v_1_2 : v_1_2 = cint(request("sel1_2"))
	dim v_1_3 : v_1_3 = cint(request("sel1_3"))
	dim v_1_4 : v_1_4 = cint(request("sel1_4"))
	dim v_1_5 : v_1_5 = cint(request("sel1_5"))
	dim v_1_6 : v_1_6 = cint(request("sel1_6"))
	dim v_1_7 : v_1_7 = cint(request("sel1_7"))
	dim v_1_8 : v_1_8 = cint(request("sel1_8"))
	dim a_val : a_val = v_1_1 + v_1_2+  v_1_3 + v_1_4 + v_1_5 + v_1_6 + v_1_7 + v_1_8
	dim a_tot : a_tot = 40
	dim a_fee : a_fee = a_val / a_tot 


	
	dim v_2_1 : v_2_1 = cint(request("sel2_1"))
	dim v_2_2 : v_2_2 = cint(request("sel2_2"))
	dim v_2_3 : v_2_3 = cint(request("sel2_3"))
	dim v_2_4 : v_2_4 = cint(request("sel2_4"))
	dim b_val : b_val = v_2_1 + v_2_2 + v_2_3 + v_2_4
	dim b_tot : b_tot = 20
	dim b_fee : b_fee = b_val / b_tot 


	dim v_3_1 : v_3_1 = cint(request("sel3_1"))
	dim v_3_2 : v_3_2 = cint(request("sel3_2"))
	dim v_3_3 : v_3_3 = cint(request("sel3_3"))
	dim v_3_4 : v_3_4 = cint(request("sel3_4"))
	dim v_3_5 : v_3_5 = cint(request("sel3_5"))
	dim c_val : c_val = v_3_1 + v_3_2 + v_3_3 + v_3_4 + v_3_5
	dim c_tot : c_tot = 30
	dim c_fee : c_fee = c_val / c_tot 

	
	dim v_4_1 : v_4_1 = cint(request("sel4_1"))
	dim v_4_2 : v_4_2 = cint(request("sel4_2"))
	dim d_val : d_val = v_4_1 + v_4_2 
	dim d_tot : d_tot = 5
	dim d_fee : d_fee = d_val / d_tot 

	
	dim v_5_1 : v_5_1 = cint(request("sel5_1"))
	dim v_5_2 : v_5_2 = cint(request("sel5_2"))
	dim e_val : e_val = v_5_1 + v_5_2 
	dim e_tot : e_tot = 5
	dim e_fee : e_fee = e_val / e_tot 

	dim objrs, sql
	sql = "select tidx, mdidx, editdate, a_val, a_fee, b_val, b_fee, c_val, c_fee, d_val, d_fee, e_val, e_fee, avg, class, cuser, cdate, uuser, udate, isuse from dbo.wb_validation_class  where tidx = " & tidx
	call set_recordset(objrs, sql)
	
	if objrs.eof then 
		objrs.addnew
		objrs.fields("mdidx").value = mdidx
		objrs.fields("editdate").value = date
		objrs.fields("a_val").value = a_val
		objrs.fields("a_fee").value = a_fee
		objrs.fields("b_val").value = b_val
		objrs.fields("b_fee").value = b_fee
		objrs.fields("c_val").value = c_val
		objrs.fields("c_fee").value = c_fee
		objrs.fields("d_val").value = d_val
		objrs.fields("d_fee").value = d_fee
		objrs.fields("e_val").value = e_val
		objrs.fields("e_fee").value = e_fee
		objrs.fields("avg").value = avg
		objrs.fields("class").value = mclass
		objrs.fields("cuser").value = request.cookies("userid")
		objrs.fields("cdate").value = date
		objrs.fields("uuser").value = request.cookies("userid")
		objrs.fields("udate").value = date
		objrs.fields("isuse").value = 1
		objrs.update
		
		tidx = objrs.fields("tidx").value
	else 
		objrs.fields("editdate").value = date
		objrs.fields("a_val").value = a_val
		objrs.fields("a_fee").value = a_fee
		objrs.fields("b_val").value = b_val
		objrs.fields("b_fee").value = b_fee
		objrs.fields("c_val").value = c_val
		objrs.fields("c_fee").value = c_fee
		objrs.fields("d_val").value = d_val
		objrs.fields("d_fee").value = d_fee
		objrs.fields("e_val").value = e_val
		objrs.fields("e_fee").value = e_fee
		objrs.fields("avg").value = avg
		objrs.fields("class").value = mclass
		objrs.fields("uuser").value = request.cookies("userid")
		objrs.fields("udate").value = date
		objrs.update
	end if
	objrs.close
	
	sql = "select tidx, code, value from dbo.wb_validation_value  where tidx = " & tidx
	call set_recordset(objrs, sql)
	if not objrs.eof then 
		do until objrs.eof 
			objrs.delete
		objrs.movenext
		loop
	end if 
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel1_1"
	objrs.fields("value").value = v_1_1
	objrs.update

	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel1_2"
	objrs.fields("value").value = v_1_2
	objrs.update

	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel1_3"
	objrs.fields("value").value = v_1_3
	objrs.update

	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel1_4"
	objrs.fields("value").value = v_1_4
	objrs.update

	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel1_5"
	objrs.fields("value").value = v_1_5
	objrs.update

	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel1_6"
	objrs.fields("value").value = v_1_6
	objrs.update

	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel1_7"
	objrs.fields("value").value = v_1_7
	objrs.update

	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel1_8"
	objrs.fields("value").value = v_1_8
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel2_1"
	objrs.fields("value").value = v_2_1
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel2_2"
	objrs.fields("value").value = v_2_2
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel2_3"
	objrs.fields("value").value = v_2_3
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel2_4"
	objrs.fields("value").value = v_2_4
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel3_1"
	objrs.fields("value").value = v_3_1
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel3_2"
	objrs.fields("value").value = v_3_2
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel3_3"
	objrs.fields("value").value = v_3_3
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel3_4"
	objrs.fields("value").value = v_3_4
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel3_5"
	objrs.fields("value").value = v_3_5
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel4_1"
	objrs.fields("value").value = v_4_1
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel4_2"
	objrs.fields("value").value = v_4_2
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel5_1"
	objrs.fields("value").value = v_5_1
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel5_2"
	objrs.fields("value").value = v_5_2
	objrs.update

	objrs.close
	set objrs = nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	location.href = "validation_led.asp?mdidx=<%=mdidx%>&tidx=<%=tidx%>";
//-->
</SCRIPT>