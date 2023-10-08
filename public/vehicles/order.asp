<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%'response.Buffer=true%>
<!--#include file="conn.asp" -->
<html>
<head>
<title><!--#INCLUDE FILE="inc/title.htm" --></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css/css.css" type=text/css rel=stylesheet>
<style type="text/css">
.text{
 font-size:12px;
 color:#3399cc;
}
</style>
</head>
<%
  set rs=server.CreateObject("adodb.recordset")
%>
<body bgcolor="#FFFFFF" topmargin="0"  leftmargin="0"  text="#000000" background="star/dd_boot.jpg">
<table width="1000" border="0" cellspacing="0" cellpadding="0" height="509">
  <tr>
    <td valign="top" height="572" width="9"> 
      <table width="980" border="0" cellspacing="0" cellpadding="0" height="70">
        <tr> 
          <td><!--#include file="inc/top.htm" --></td>
        </tr>
      </table>
      <table width="1000" border="0" cellspacing="0" cellpadding="0" height="86" align="center">
        <tr> 
          <td background="star/dd_boot_6.jpg" valign="top"> 
            <table width="979" border="0" cellspacing="0" cellpadding="0" height="131">
              <tr> 
                <td height="114">&nbsp;</td>
              </tr>
              <tr> 
                <td height="14"> 
                  <table width="980" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="250"><img src="star/dd_boot_7.jpg" width="251" height="14"></td>
                      <td width="730"><img src="star/dd_boot_8.jpg" width="729" height="14"></td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <table width="996" border="0" cellspacing="0" cellpadding="0" height="540">
        <tr> 
          <td width="256" height="377" valign="top"> 
            <table width="256" border="0" cellspacing="0" cellpadding="0" height="251">
              <tr> 
                <td width="11" height="274">&nbsp;</td>
                <td valign="top" height="274"> 
                  <table width="240" border="0" cellspacing="0" cellpadding="0" height="25">
                    <tr> 
                      <td height="60"> 
                        <div align="center"><img src="star/dd_1.jpg" width="211" height="44"></div>
                      </td>
                    </tr>
                    <tr> 
                      <td height="25"> 
                        <div align="center"> <a href="products.asp?id=1"> </a><a href="products.asp?id=1"><img src="star/dd_2_a_1.jpg" width="196" height="25" border="0"></a> 
                        </div>
                      </td>
                    </tr>
                    <tr> 
                      <td height="25"> 
                        <div align="center"> <a href="products.asp?id=2"> </a><a href="products.asp?id=2"><img src="star/dd_2_a_2.jpg" width="196" height="25" border="0"></a> 
                        </div>
                      </td>
                    </tr>
                    <tr> 
                      <td height="25"> 
                        <div align="center"> <a href="products.asp?id=3"> </a><a href="products.asp?id=3"><img src="star/dd_2_a_3.jpg" width="196" height="25" border="0"></a> 
                        </div>
                      </td>
                    </tr>
                    <tr> 
                      <td height="25"> 
                        <div align="center"> <a href="products.asp?id=4"> </a><a href="products.asp?id=4"><img src="star/dd_2_a_4.jpg" width="196" height="25" border="0"></a> 
                        </div>
                      </td>
                    </tr>
                    <tr> 
                      <td height="25"> 
                        <div align="center"> <a href="products.asp?id=5"> </a><a href="products.asp?id=5"><img src="star/dd_2_a_5.jpg" width="196" height="25" border="0"></a> 
                        </div>
                      </td>
                    </tr>
                    <tr> 
                      <td height="25"> 
                        <div align="center"> <a href="products.asp?id=6"> </a><a href="products.asp?id=6"><img src="star/dd_2_a_6.jpg" width="196" height="25" border="0"></a> 
                        </div>
                      </td>
                    </tr>
                    <tr> 
                      <td height="25"> 
                        <div align="center"> <a href="products.asp?id=7"> </a><a href="products.asp?id=7"><img src="star/dd_2_a_7.jpg" width="196" height="25" border="0"></a> 
                        </div>
                      </td>
                    </tr>
                    <tr> 
                      <td height="25"> 
                        <div align="center"> <a href="new_products.asp"> </a><a href="new_products.asp"><img src="star/dd_2_a_8.jpg" width="196" height="25" border="0"></a> 
                        </div>
                      </td>
                    </tr>
                    <tr> 
                      <td height="25"> 
                        <div align="center"><img src="star/dd_2_a_9.jpg" width="196" height="31" border="0"></div>
                      </td>
                    </tr>
                  </table>
                </td>
                <td width="5" height="274">&nbsp;</td>
              </tr>
            </table>
          </td>
          <td height="377" valign="top">
            <table width="732" border="0" cellspacing="0" cellpadding="0" height="563">
              <tr> 
                <td width="723" height="563" valign="top">
<table width="721" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td><table width="710" border="0" cellspacing="1" cellpadding="0" align="center" height="50" bgcolor="#CCCCCC">
                          <tr> 
                            <td height="50" bgcolor="#F6F6F6"> 
							<form name="form1" method="post" action="search.asp" onsubmit="return check1(form1)">
                                <table width="707" height="36" border="0" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td width="125" valign="bottom"> <div align="center"><strong><font color="#6699CC" size="4">Search</font></strong></div></td>
                                    <td width="165" valign="bottom"> 
                                      <%
								dim id(),v_type(),type_code(),types,i
								i=1
							    sql="select * from type"
								set rs=conn.execute(sql)
								do while not rs.eof
								 redim preserve id(i),v_type(i),type_code(i)
								 id(i)=rs("id")
								 v_type(i)=rs("type")
								 type_code(i)=rs("type_code")
								 if not rs.eof then rs.movenext
								 i=i+1
								loop
								rs.close
								types=ubound(id)
							  %>
                                      <select name="type" id="type">
									    <option value="">===All===</option>
                                        <%for i=1 to types%>
                                        <option value="<%=type_code(i)%>"><%=v_type(i)%></option>
                                        <%
								 next
								%>
                                      </select> </td>
                                    <td width="269" valign="bottom"> <div align="center"><strong><font color="#6699CC">Keywords:</font></strong> 
                                        <input name="keywords" type="text" id="keywords" size="15">
                                      </div></td>
                                    <td width="148" valign="bottom"> 
									<input type="submit" name="Submit" value="Search"> 
                                    </td>
                                  </tr>
                                </table>
                              </form></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr>
                      <td height=11></td>
                    </tr>
                    <tr>
                      <td height="355" valign="top">
					  <form name="form2" method="post" action="script/addorder.asp"  onsubmit="return check2(this)">
                          <table width="459" height="403" border="0" align="center" cellpadding="0" cellspacing="0" class="text">
                            <tr> 
                              <td width="148" height="27"> Type: </td>
                              <td width="311">
							  <select name="type" id="type">
                                        <%for i=1 to types%>
                                        <option value="<%=v_type(i)%>"><%=v_type(i)%></option>
                                        <%
								 next
								%>
                            </select></td>
                            </tr>
                            <tr> 
                              <td height="27"> Model No. : </td>
                              <td><input name="model" type="text" id="model">
                                *</td>
                            </tr>
                            <tr> 
                              <td height="27"> Order quantity: </td>
                              <td><input name="quantity" type="text" id="quantity"></td>
                            </tr>
                            <tr> 
                              <td height="28"> Expected delivery time:</td>
                              <td><input name="delivery_time" type="text" id="delivery_time"></td>
                            </tr>
                            <tr> 
                              <td height="27" colspan="2"><font color="#999999">Your 
                                Contact Information</font></td>
                            </tr>
                            <tr> 
                              <td height="27">Contact person:</td>
                              <td><input name="contact_person" type="text" id="contact_person"></td>
                            </tr>
                            <tr> 
                              <td height="27">Tel:</td>
                              <td><input name="tel" type="text" id="tel"></td>
                            </tr>
                            <tr> 
                              <td height="27">Fax:</td>
                              <td><input name="fax" type="text" id="fax"></td>
                            </tr>
                            <tr> 
                              <td height="27">Address:</td>
                              <td><input name="address" type="text" id="address"></td>
                            </tr>
                            <tr> 
                              <td height="27">Email: </td>
                              <td><input name="email" type="text" id="email"></td>
                            </tr>
                            <tr> 
                              <td height="90">Notes:</td>
                              <td><textarea name="notes" cols="30" rows="5" id="notes"></textarea></td>
                            </tr>
                            <tr> 
                              <td height="12">&nbsp;</td>
                              <td>&nbsp;</td>
                            </tr>
                            <tr> 
                              <td height="30"><div align="right"> </div></td>
                              <td> <input type="submit" name="Submit2" value="Submit"> 
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="reset" name="Submit3" value="Reset"> 
                              </td>
                            </tr>
                          </table>
                        </form></td>
                    </tr>
                    <tr>
                      <td>

                      </td>
                    </tr>
                    <tr>
                      <td><div align="center"><img src="star/dd_boot_9.jpg" width="710" height="27"></div></td>
                    </tr>
                    <tr>
                      <td>
<%
dim p_id(),modelno(),photo(),j,introduction(),sql,commended_pros,c_reds
sql="select id from products where ifcommended=true"
rs.open sql,conn,1,1
c_reds=rs.recordcount
if c_reds>=4 then
    sql="select top 4 * from products where ifcommended=true"
  elseif c_reds<4 and c_reds>0 then
    sql="select * from products where ifcommended=true"
end if
rs.close
j=1
set rs=conn.execute(sql)
do while not rs.eof
  redim preserve p_id(j),modelno(j),photo(j),introduction(j)
  p_id(j)=rs("id")
  modelno(j)=rs("modelno")
  photo(j)=rs("photo")
  introduction(j)=rs("introduction")
  if not rs.eof then rs.movenext
  j=j+1
loop
rs.close
commended_pros=ubound(p_id)
%>
<table width="710" border="0" cellpadding="0" align="center" height="130" cellspacing="0">
                          <tr> 
                            <td width="169" height="130"> 
                              <%if commended_pros>=1 then%>
      <table width="120" border="0" cellspacing="0" cellpadding="0" height="66" align="center">
        <tr> 
          <td width="120">
<a href="#" onClick="javascript:window.open('detail.asp?id=<%=p_id(1)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
		  <img src="proimages/<%=photo(1)%>" width="120" height="104" border="0">
		  </a>
		  </td>
        </tr>
        <tr> 
          <td height="25"> 
		  <div align="center">
		     <font size="2">
<a href="#" onClick="javascript:window.open('detail.asp?id=<%=p_id(1)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
		       <%=modelno(1)%>
			   </a>
		     </font>
		  </div>
		  </td>
        </tr>
      </table>
	  <%end if%>
	  </td>
                            <td width="180" height="181"> 
                              <%if commended_pros>=2 then%>
      <table width="120" border="0" cellspacing="0" cellpadding="0" height="66" align="center">
        <tr> 
          <td>
<a href="#" onClick="javascript:window.open('detail.asp?id=<%=p_id(2)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
		  <img src="proimages/<%=photo(2)%>" width="120" height="104" border="0">
		  </a>
		  </td>
        </tr>
        <tr> 
          <td height="25"> 
		  <div align="center">
		  <font size="2">
<a href="#" onClick="javascript:window.open('detail.asp?id=<%=p_id(2)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
		  <%=modelno(2)%>
		  </a>
		  </font>
		  </div></td>
        </tr>
      </table>
	  <%end if%>
	  </td>
                            <td width="180" height="130"> 
                              <%if commended_pros>=3 then%>
      <table width="120" border="0" cellspacing="0" cellpadding="0" height="66" align="center">
        <tr> 
          <td>
<a href="#" onClick="javascript:window.open('detail.asp?id=<%=p_id(3)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
		  <img src="proimages/<%=photo(3)%>" width="120" height="104" border="0">
		  </a>
		  </td>
        </tr>
        <tr> 
          <td height="25">
		   <div align="center">
		   <font size="2">
<a href="#" onClick="javascript:window.open('detail.asp?id=<%=p_id(3)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
		    <%=modelno(3)%>
			</a>
		   </font>
		   </div></td>
        </tr>
      </table>
	  <%end if%>
	  </td>
                            <td width="181" height="130"> 
                              <%if commended_pros>=4 then%>
      <table width="120" border="0" cellspacing="0" cellpadding="0" height="66" align="center">
        <tr> 
          <td>
<a href="#" onClick="javascript:window.open('detail.asp?id=<%=p_id(4)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
		  <img src="proimages/<%=photo(4)%>" width="120" height="104" border="0">
		  </A>
		  </td>
        </tr>
        <tr> 
          <td height="25"> 
		  <div align="center">
		  <font size="2">
<a href="#" onClick="javascript:window.open('detail.asp?id=<%=p_id(4)%>','','width=400,height=400,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
		  <%=modelno(4)%>
		  </A>
		  </font>
		  </div></td>
        </tr>
      </table>
	  <%end if%>
	  </td>
  </tr>
</table>
</td>
                    </tr>
                  </table> </td>
                <td height="563" valign="top" background="star/dd_boot_22.jpg">
				<img src="star/dd_boot_2.jpg" width="9" height="23"></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="1000" border="0" cellspacing="0" cellpadding="0" height="10">
  <tr> 
    <td background="star/dd_bor_1.jpg" height="10"></td>
  </tr>
</table>
<table width="1000" border="0" cellspacing="0" cellpadding="0" height="91">
  <tr> 
    <td bgcolor="#EAEAEA" height="32">
	<div align="center">
	<!--#INCLUDE FILE="inc/bottom.htm" -->
	</div></td>
  </tr>
</table>
<%
conn.close
set conn=nothing
%>

<script language="JScript">
<!--
function form2submit()
{
  document.form2.submit();
}

function check1(form1)
{
if (form1.keywords.value=="")
 {
    alert("Input your keywords!")
    return false
  }
  else
  {
  return true
  }
}
function check2(form2)
{
if (form2.model.value=="")
 {
    alert("Input your model!")
	form2.model.focus()
    return false
  }
  else
  {
  return true
  }
}
//-->
</script>
</body>
</html>
