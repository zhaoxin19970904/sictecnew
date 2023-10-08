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
  <table border="0" cellpadding="0" cellspacing="4" bgcolor="#C4E1E1">
    <tr> 
      <td bgcolor="#FFFFFF"><img src="image/quotation.jpg" width="127" height="19"></td>
    </tr>
    <tr> 
      <td valign="top">
	  <%
	  prod_name=request("productname") 
	  %>
	   <%
					    username=trim(session("username"))
						if username<>"" then
						  sql="select company,firstname,lastname,tel,fax,email from member where member_uid='"&username&"'"
						  set rs=conn.execute(sql)
						else
						  m=2
						end if
						%>

<form name="form1" method="post" action="script03/mail.asp" onSubmit="return check(this)">
          <table width="515" height="329" border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="#FFFFFF">
            <tr bgcolor="#88C4C4"> 
              <td height="25" colspan="2"  class="unnamed1"><strong>　Demand:</strong></td>
            </tr>
            <tr> 
              <td width="139" height="25"  class="unnamed1"><div align="right">Product 
                  Name:<strong></strong> </div></td>
              <td width="371"  class="unnamed1"><strong><font color="#FF0000"> 
                </font><font color="#FF0000"> </font><font color="#0066CC"></font><font color="#FF0000"> 
                <input name="pro_name" type="text" class="unnamed1" id="ordervolume2" value="<%=prod_name%>" maxlength="80">
                </font></strong></td>
            </tr>
            <tr> 
              <td height="23"  class="unnamed1"> <div align="right">Order Volume 
                  (pcs):</div></td>
              <td  class="unnamed1"><font color="#FF0000"> 
                <input name="ordervolume" type="text" id="ordervolume" size="30" class="unnamed1">
                </font></td>
            </tr>
            <tr> 
              <td height="25"  class="unnamed1"><div align="right">Expected delivery 
                  time:</div></td>
              <td  class="unnamed1"> <input name="ed_time" type="text" id="ed_time2" size="30"></td>
            </tr>
            <tr> 
              <td height="25"  class="unnamed1"><div align="right">Notes:</div></td>
              <td  class="unnamed1"><font color="#FF0000"> 
                <textarea name="notes" cols="48" rows="4" id="textarea"></textarea>
                </font></td>
            </tr>
            <tr bgcolor="#88C4C4"> 
              <td height="24" colspan="2"  class="unnamed1"><strong>　Your contact 
                information:</strong></td>
            </tr>
            <tr>
              <td height="25"  class="unnamed1"><div align="right"><B>&nbsp;</B>Company 
                  Name:</div></td>
              <td><INPUT class=fonten id=company2 size=45 name=company_name value="<%if m=2 then response.Write "" else response.Write rs("company") end if%>"></td>
            </tr>
            <tr> 
              <td height="25"  class="unnamed1"><div align="right">Name:</div></td>
              <td><input name="contact_name" type="text" id="contact_name"  size="20" value="<%if m=2 then response.Write "" else response.Write rs("firstname")&" "&rs("lastname") end if%>" > 
                <span class="unnamed1"><font color="#FF0000">*</font></span></td>
            </tr>
            <tr> 
              <td height="25"  class="unnamed1"><div align="right">Tel: </div></td>
              <td><input name="contact_tel" type="text" id="contact_tel2" value="<%if m=2 then response.Write "" else response.Write rs("tel") end if%>">
                <span class="unnamed1"><font color="#FF0000">*</font> </span></td>
            </tr>
            <tr> 
              <td height="25"  class="unnamed1"><div align="right">Fax:</div></td>
              <td><input name="fax" type="text" id="fax2" value="<%if m=2 then response.Write "" else response.Write rs("fax") end if%>"> </td>
            </tr>
            <tr> 
              <td height="25"  class="unnamed1"> <div align="right">E-mail:</div></td>
              <td><input name="email" type="text" id="email2" value="<%if m=2 then response.Write "" else response.Write rs("email") end if%>">
                <span class="unnamed1"><font color="#FF0000">* </font> </span></td>
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
}
//-->
</script>
</body>
</html>
