<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->

<html>
<head>
<title>Welcome to SICTEC !</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
a:active {  text-decoration: none}
a:hover {  }
a:link {  text-decoration: none}
a:visited {  text-decoration: none}
.unnamed1 {
	font-family: "Arial", "Helvetica", "sans-serif";
	font-size: 12px;
	color: #000000;
}
-->
</style>
<style type="text/css">
<!--
.unnamed2 {
	font-family: "Arial", "Helvetica", "sans-serif";
	font-size: 14px;
}
-->
</style>
</head>

<body background="image/bg01.jpg" text="#000000"   link="#000000" vlink="#000000" alink="#000000" topmargin="0"  leftmargin="0">
<div align="left"> 
  <table border="0" cellpadding="0" cellspacing="7" bgcolor="#C4E1E1">
    <tr> 
      <td bgcolor="#FFFFFF"><img src="image/bg16.jpg" width="127" height="19"></td>
    </tr>
    <tr> 
      <td>
            <%
					    username=trim(session("username"))
						if username<>"" then
						  sql="select firstname,lastname,tel,fax,email from member where member_uid='"&username&"'"
						  set rs=conn.execute(sql)
						else
						  m=2
						end if
						%>
      
      <form name="form1" method="post" action="script/mail.asp" onSubmit="return check(this)">
          <table width="510" height="369" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
            <tr bgcolor="#8CC6FF"> 
              <td height="25" colspan="2"  class="unnamed1"><strong>Your contact 
                information</strong></td>
            </tr>
            <tr> 
              <td height="25"  class="unnamed1"><div align="right">Name:</div></td>
              <td  class="unnamed1"><input name="contact_name" type="text" id="contact_name" size="20" value="<%if m=2 then response.Write "" else response.Write rs("firstname")&" "&rs("lastname") end if%>"> 
                <font color="#FF0000">*</font></td>
            </tr>
            <tr> 
              <td height="25"  class="unnamed1"><div align="right">Tel: </div></td>
              <td  class="unnamed1"><input name="contact_tel" type="text" id="contact_tel" value="<%if m=2 then response.Write "" else response.Write rs("tel") end if%>"> 
                <font color="#FF0000">*</font></td>
            </tr>
            <tr> 
              <td height="25"  class="unnamed1"><div align="right">Fax:</div></td>
              <td  class="unnamed1"><input name="fax" type="text" id="fax" value="<%if m=2 then response.Write "" else response.Write rs("fax") end if%>"></td>
            </tr>
            <tr> 
              <td height="25"  class="unnamed1"><div align="right">E-mail:</div></td>
              <td  class="unnamed1"><input name="email" type="text" id="email" value="<%if m=2 then response.Write "" else response.Write rs("email") end if%>"> 
                <font color="#FF0000">*</font></td>
            </tr>
            <tr bgcolor="#B9FFFF"> 
              <td height="24" colspan="2"  class="unnamed1"><strong>Demand:</strong></td>
            </tr>
            <tr> 
              <td width="153" height="25"  class="unnamed1"><div align="right">Category:</div></td>
              <td width="357"  class="unnamed1"><FONT color=#000000> 
                <select name="category" id="category">
                  <option>Transportation </option>
                  <option>Construction &amp; Hardware</option>
                  <option>Medical &amp; Health care</option>
                  <option>Home Appliance </option>
                  <option>Gifts, Models &amp; Crafts</option>
                  <option>Sports &amp; Leisure </option>
                  <option>Security &amp; Protection </option>
                  <option>Office supplies </option>
                  <option>Electronics &amp; Electrical </option>
                  <option>Sourcing Services</option>
                  <option>New Promotional Items</option>
                  <option>Others</option>
                </select>
                </FONT></td>
            </tr>
            <tr> 
              <td height="25"  class="unnamed1"><div align="right">Product Name:</div></td>
              <td> <input name="pro_name" type="text" id="pro_name4" size="30"> 
                <span class="unnamed1"><font color="#FF0000">*</font></span></td>
            </tr>
            <tr> 
              <td height="25"  class="unnamed1"><div align="right">Order Volume 
                  (pcs):</div></td>
              <td> <input name="ordervolume" type="text" id="price4" size="30" class="unnamed1"></td>
            </tr>
            <tr> 
              <td height="25"  class="unnamed1"><div align="right">Expected delivery 
                  time:</div></td>
              <td> <input name="ed_time" type="text" id="ed_time" size="30"></td>
            </tr>
            <tr> 
              <td height="90"  class="unnamed1"><div align="right">Notes:</div></td>
              <td> <textarea name="notes" cols="50" rows="5" id="textarea3"></textarea></td>
            </tr>
            <tr> 
              <td>&nbsp;</td>
              <td height="30"> <input type="submit" name="Submit" value="Submit"> 
                &nbsp;&nbsp;&nbsp; <input type="reset" name="Submit2" value="Reset" onClick="javascript:document.form1.reset();return false;"> 
              </td>
            </tr>
          </table>
        </form></td>
    </tr>
  </table>
</div>
<script language="JScript">
<!--
function check(form1)
{
if (form1.contact_name.value=="")
 {
    alert("Pls input your name,thanks!")
	form1.contact_name.focus()
    return false
  }
if (form1.contact_tel.value=="")
 {
    alert("Pls give us your telephone number,thanks!")
	form1.contact_tel.focus()
    return false
  }
if (form1.email.value=="")
 {
    alert("Pls input your email address,thanks!")
	form1.email.focus()
    return false
  }
if (form1.pro_name.value=="")
 {
    alert("Pls input product name,thanks!")
	form1.pro_name.focus()
    return false
	}
}
//-->
</script>
</body>
</html>
