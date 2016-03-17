<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim contidx : contidx = request("contidx")
	dim tidx : tidx = request("tidx")
	if tidx = "" then tidx = 0
	dim mclass : mclass = request("txtclass")
	dim avg : avg = request("txtavg")

	dim v_1_1 : v_1_1 = cint(request("1_1"))
	dim v_1_2 : v_1_2 = cint(request("1_2"))
	dim v_1_3 : v_1_3 = cint(request("1_3"))
	dim v_1_4 : v_1_4 = cint(request("1_4"))
	dim v_1_5 : v_1_5 = cint(request("1_5"))
	dim sel1_1 : sel1_1 = cint(request("sel1_1"))
	dim sel1_2 : sel1_2 = cint(request("sel1_2"))
	dim sel1_3 : sel1_3 = cint(request("sel1_3"))
	dim sel1_4 : sel1_4 = cint(request("sel1_4"))
	dim sel1_5 : sel1_5 = cint(request("sel1_5"))
	dim a_val : a_val = (v_1_1 + v_1_2+  v_1_3 + v_1_4 + v_1_5) / 4
	dim a_selval : a_selval = sel1_1 + sel1_2+  sel1_3 + sel1_4 + sel1_5
	dim a_tot : a_tot = 30
	dim a_fee : a_fee = a_selval / a_tot 


	
	dim v_2_1 : v_2_1 = cint(request("2_1"))
	dim v_2_2 : v_2_2 = cint(request("2_2"))
	dim v_2_3 : v_2_3 = cint(request("2_3"))
	dim v_2_4 : v_2_4 = cint(request("2_4"))
	dim v_2_5 : v_2_5 = cint(request("2_5"))
	dim v_2_6 : v_2_6 = cint(request("2_6"))
	dim sel2_1 : sel2_1 = cint(request("sel2_1"))
	dim sel2_2 : sel2_2 = cint(request("sel2_2"))
	dim sel2_3 : sel2_3 = cint(request("sel2_3"))
	dim sel2_4 : sel2_4 = cint(request("sel2_4"))
	dim sel2_5 : sel2_5 = cint(request("sel2_5"))
	dim sel2_6 : sel2_6 = cint(request("sel2_6"))
	dim b_val : b_val = (v_2_1 + v_2_2 + v_2_3 + v_2_4 + v_2_5 + v_2_6) / 4
	dim b_selval : b_selval = sel2_1 + sel2_2 + sel2_3 + sel2_4 + sel2_5 + sel2_6
	dim b_tot : b_tot = 40
	dim b_fee : b_fee = b_selval / b_tot 


	dim v_3_1 : v_3_1 = cint(request("3_1"))
	dim v_3_2 : v_3_2 = cint(request("3_2"))
	dim v_3_3 : v_3_3 = cint(request("3_3"))
	dim sel3_1 : sel3_1 = cint(request("sel3_1"))
	dim sel3_2 : sel3_2 = cint(request("sel3_2"))
	dim sel3_3 : sel3_3 = cint(request("sel3_3"))
	dim c_val : c_val =  (v_3_1 + v_3_2 + v_3_3) / 4
	dim c_selval : c_selval =  sel3_1 + sel3_2 + sel3_3
	dim c_tot : c_tot = 15
	dim c_fee : c_fee = c_selval / c_tot 

	
	dim v_4_1 : v_4_1 = cint(request("4_1"))
	dim v_4_2 : v_4_2 = cint(request("4_2"))
	dim v_4_3 : v_4_3 = cint(request("4_3"))
	dim sel4_1 : sel4_1 = cint(request("sel4_1"))
	dim sel4_2 : sel4_2 = cint(request("sel4_2"))
	dim sel4_3 : sel4_3 = cint(request("sel4_3"))
	dim d_val : d_val = (v_4_1 + v_4_2  + v_4_3) / 4
	dim d_selval : d_selval = sel4_1 + sel4_2 + sel4_3
	dim d_tot : d_tot = 15
	dim d_fee : d_fee = d_selval / d_tot 

	
	dim e_val : e_val = 0
	dim e_fee : e_fee = 0

	sql = "select tidx, contidx,  isuse from dbo.wb_validation_class  where contidx = " & contidx

	call set_recordset(objrs, sql)
	if Not objrs.eof then 
		do until objrs.eof 
			objrs.fields("isuse").value = 0
			objrs.update
			objrs.movenext
		loop
	end if 
	objrs.close

	dim objrs, sql
	sql = "select tidx, contidx, editdate, a_val, a_fee, b_val, b_fee, c_val, c_fee, d_val, d_fee, e_val, e_fee, avg, class, cuser, cdate, uuser, udate, isuse from dbo.wb_validation_class  where tidx = " & tidx
	call set_recordset(objrs, sql)
	
	if objrs.eof then 
		objrs.addnew
		objrs.fields("contidx").value = contidx
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
	objrs.fields("value").value = sel1_1
	objrs.update

	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel1_2"
	objrs.fields("value").value = sel1_2
	objrs.update

	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel1_3"
	objrs.fields("value").value = sel1_3
	objrs.update

	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel1_4"
	objrs.fields("value").value = sel1_4
	objrs.update

	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel1_5"
	objrs.fields("value").value = sel1_5
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel2_1"
	objrs.fields("value").value = sel2_1
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel2_2"
	objrs.fields("value").value = sel2_2
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel2_3"
	objrs.fields("value").value = sel2_3
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel2_4"
	objrs.fields("value").value = sel2_4
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel2_5"
	objrs.fields("value").value = sel2_5
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel2_6"
	objrs.fields("value").value = sel2_6
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel3_1"
	objrs.fields("value").value = sel3_1
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel3_2"
	objrs.fields("value").value = sel3_2
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel3_3"
	objrs.fields("value").value = sel3_3
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel4_1"
	objrs.fields("value").value = sel4_1
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel4_2"
	objrs.fields("value").value = sel4_2
	objrs.update
	
	objrs.addnew
	objrs.fields("tidx").value = tidx
	objrs.fields("code").value = "sel4_3"
	objrs.fields("value").value = sel4_3
	objrs.update

	objrs.close
	set objrs = nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	location.href = "validation_etc.asp?contidx=<%=contidx%>&tidx=<%=tidx%>";
//-->
</SCRIPT>