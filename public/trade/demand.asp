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
	line-height: 20px;
}
-->
</style>
<style type="text/css">
<!--
.unnamed2 {
	font-family: "Arial", "Helvetica", "sans-serif";
	font-size: 14px;
}
.unnamed3 {	font-family: "Arial", "Helvetica", "sans-serif";
	font-size: 11px;
}
-->
</style>
</head>

<body background="image/bg01.jpg" text="#000000"   link="#000000" vlink="#000000" alink="#000000" topmargin="0" >
<div align="left">
<table width="800" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999">
    <tr>
      <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
          <tr>
            <td width="25%" height="318" valign="top" bgcolor="#D8D8D8"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td><img src="image/logo.jpg" width="201" height="66"></td>
                </tr>
                <tr> 
                  <td height="30" bgcolor="#D9E2E7" class="unnamed2"> <div align="center"><font color="#000033"><strong>MEMBER 
                      LOGIN</strong></font></div></td>
                </tr>
                <tr> 
                  <td> 
                    <%
                username=trim(session("username"))
                if username="" then%>
                    <table border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="200" height="134" valign="top" background="image/bg02.jpg"> 
                          <FORM name="form1" action="script/login.asp" method="post">
                            <table height=120 cellspacing=0 width="99%">
                              <tbody>
                                <tr> 
                                  <td width="41%" height=37 valign="bottom" class="unnamed1"> 
                                    <div align=right><br>
                                      User ID: </div></td>
                                  <td height=37 valign="bottom"> <div align=left><font color=#000000> 
                                      <input id=username 
                  style="BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; BORDER-LEFT: 1px double; BORDER-BOTTOM: 1px double; BACKGROUND-COLOR: #ffffff; TEXT-ALIGN: left" 
                  size=11 name="username" border-color:#bbecf6>
                                      </font></div></td>
                                </tr>
                                <tr> 
                                  <td width="41%" height=28 class="unnamed1"> 
                                    <div align=right>Password: </div></td>
                                  <td height=28> <div align=left><font color=#000000> 
                                      <input id=userpass 
                  style="BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; BORDER-LEFT: 1px double; BORDER-BOTTOM: 1px double; BACKGROUND-COLOR: #ffffff; TEXT-ALIGN: left" 
                  type=password size=11 name="userpass" border-color:#bbecf6>
                                      </font></div></td>
                                </tr>
                                <tr> 
                                  <td height=20 colspan=2 valign=bottom> <div align=center> 
                                      <a href="script/login.asp"> 
                                      <INPUT name="image" type=image 
                  onclick="return checkempty(form1)" src="star/new_pa1.gif" 
                  border=0>
                                      </a>&nbsp; </div></td>
                                </tr>
                                <tr> 
                                  <td height=17 colspan=2 valign=top> <div align=center class="unnamed1"><b><a href="lost_p_w.asp"><font color="#000000">Forgot</font></a> 
                                      ? <a href="register.asp"><font color="#000000">Register</font></a> 
                                      !</b> </div></td>
                                </tr>
                              </tbody>
                            </table>
                          </form></td>
                      </tr>
                    </table></td>
                </tr>
                <% else %>
                <%
					 set rs=server.createobject("adodb.recordset")
					sql="select firstname,lastname,email from member where member_uid='"&username&"'"
						  set rs=conn.execute(sql)
						  firstname=rs("firstname")
						  lastname=rs("lastname")
						  email=rs("email")
						%>
                <tr> 
                  <td><table border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="200" height="134" valign="top" background="image/bg02.jpg"> 
                          <FORM name="form2" action="script/login.asp" method="post">
                            <table height=120 cellspacing=0 width="99%">
                              <tbody>
                                <tr> 
                                  <td width="41%" height=37 valign="bottom" class="unnamed1" align="center"> 
                                    <div align=right><br>
                                      Hello,</div></td>
                                  <td height=37 valign="bottom" class="unnamed1"> 
                                    <div align=left><%=firstname%>&nbsp;&nbsp;<%=lastname%></div></td>
                                </tr>
                                <tr> 
                                  <td width="41%" height=28 class="unnamed1" colspan="2"> 
                                    <div align=center><%=email%> </div></td>
                                </tr>
                                <tr> 
                                  <td height=20 colspan=2 valign=bottom> <div align=center class="unnamed1"> 
                                      <b> <a href="changereg.asp"><font color="#000000"> 
                                      &nbsp;Update my profile</font></a> </b></div></td>
                                </tr>
                                <tr> 
                                  <td height=17 colspan=2 valign=top> <div align=center class="unnamed1"><b><a href="script/logout.asp"><font color="#000000">Exit</font></a> 
                                      </b> </div></td>
                                </tr>
                              </tbody>
                            </table>
                          </form></td>
                      </tr>
                    </table>
                    <% 
                              end if
                %>
                  </td>
                </tr>
                <tr> 
                  <td valign="top"><marquee id=ca onMouseOver=ca.stop() onMouseOut=ca.start() scrollamount=1 direction=up height=100>
                    <table width="88%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td valign="top"><img width="1" height="6"></td>
                      </tr>
                      
                      <tr> 
                        <td height="26" class="unnamed1"> <img src="image/b06.jpg" width="7" height="7"> 
                          <font color="#000066">Distributors/Agents welcome </font>                        </td>
                      </tr>
                    </table>
                    </marquee></td>
                </tr>
                <tr> 
                  <td height="18" bgcolor="#91B7BC"><font class="MB">　</font><span class="unnamed1"><strong>PROMOTION 
                    ITEMS</strong></span></td>
                </tr>
                <tr> 
                  <td bgcolor="#336666"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="1"><img width="2" height="1"></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td><table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td valign="top"><img width="1" height="6"></td>
                      </tr>
                      <tr> 
                        <td valign="top" class="unnamed1"><strong><img src="image/b07.jpg" width="7" height="7"></strong>　AB 
                          lounge<br> <strong><img src="image/b07.jpg" width="7" height="7"></strong>　digital 
                          voice recorder<br> <strong><img src="image/b07.jpg" width="7" height="7"></strong>　postal 
                          scales<br> <strong><img src="image/b07.jpg" width="7" height="7"></strong>　induction 
                          cooker<br> <strong><img src="image/b07.jpg" width="7" height="7"></strong>　vacuum 
                          flask</td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="10"><img width="1" height="10"></td>
                </tr>
                <tr> 
                  <td height="18" bgcolor="#91B7BC"><font class="MB">　</font><span class="unnamed1"><strong>HOT 
                    SALE</strong></span></td>
                </tr>
                <tr> 
                  <td valign="top" bgcolor="#336666"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="1"><img width="2" height="1"></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td valign="top"><table width="94%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td valign="top"><img width="1" height="10"></td>
                      </tr>
                      <tr> 
                        <td valign="top" class="unnamed1"><div align="center"><a href="product/newitems.asp"><img src="image/p06.jpg" width="137" height="45" border="0"></a></div></td>
                      </tr>
                      <tr> 
                        <td height="16" valign="top" class="unnamed1"> <div align="center" class="unnamed3"><a href="product/newitems.asp">New 
                            Items</a></div></td>
                      </tr>
                      
                      <tr> 
                        <td valign="top" class="unnamed1"><div align="center"><a href="http://www.sictec.com/tyreclock" target="_blank"><img src="image/p02.jpg" width="137" height="45" border="0"></a></div></td>
                      </tr>
                      <tr> 
                        <td height="16" valign="top" class="unnamed1"> <div align="center"><span class="unnamed3">Tire 
                            Clock</span> </div></td>
                      </tr>
                      <tr> 
                        <td valign="top" class="unnamed1"><div align="center"><a href="http://www.sictec.com/stationery" target="_blank"><img src="image/p03.jpg" width="135" height="43" border="0"></a></div></td>
                      </tr>
                      <tr> 
                        <td height="16" valign="top" class="unnamed1"> <div align="center" class="unnamed3">Stationery</div></td>
                      </tr>
                      <tr> 
                        <td valign="top" class="unnamed1"><div align="center"><a href="http://www.sictec.com/3WS" target="_blank"><img src="image/p01.jpg" width="137" height="45" border="0"></a></div></td>
                      </tr>
                      <tr> 
                        <td height="16" valign="top" class="unnamed1"> <div align="center" class="unnamed3">3 
                            Wheel Sliding Scooter<br>
                          </div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="18" bgcolor="#91B7BC" ><strong><font class="MB">　</font><span class="unnamed1">LINK</span></strong></td>
                </tr>
                <tr> 
                  <td height="1" valign="top" bgcolor="#336666"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="1"><img width="2" height="1"></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="10"><img width="1" height="10"></td>
                </tr>
                <tr> 
                  <td><div align="center"><span class="unnamed2"> 
                      <select name="select" onChange="javascript:window.open(this.options[this.selectedIndex].value)">
                        <option selected>======link======</option>
                        <option value="http://www.sourcing-universal.com">sourcing-universal</option>
                        <option value="http://www.sictec.com">sictec</option>
                      </select>
                      </span></div></td>
                </tr>
                <tr> 
                  <td height="10"><img width="1" height="10"></td>
                </tr>
              </table></td>
            <td width="75%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="25" background="image/bg03.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="0%">&nbsp;</td>
                        <td width="81%">&nbsp;</td>
                        <td width="19%" class="unnamed1"><a href="signin.asp"><font color="#666666">Sign                     
                          in</font></a><font color="#666666"> | </font><a href="register.asp"><font color="666666"> 
                          Join now</font></a><font color="#666666">&nbsp; </font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="37" background="image/bg04.jpg">&nbsp;</td>
                </tr>
                <tr> 
                  <td height="34" valign="top"><img src="image/banner04.jpg" width="599" height="34" border="0" usemap="#Map"></td>
                </tr>
                <tr> 
                  <td height="138" valign="top"><img src="image/pho16.jpg" width="599" height="139"></td>
                </tr>
                <tr> 
                  <td height="25" valign="top" background="image/bg12.jpg"><table width="95%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="6%" height="23">&nbsp;</td>
                        <td width="94%" height="34" valign="bottom"><img src="image/bg16.jpg" width="127" height="19"></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="9" valign="top"><TABLE height=315 cellSpacing=0 cellPadding=0 width="99%" bgColor=#ffffff 
      border=0>
                      <TBODY>
                        <TR> 
                          <TD vAlign=top width=562 height=273><table width="100%" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td>&nbsp;</td>
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
                                
                                <form name="form2" method="post" action="script/mail.asp" onSubmit="return check(this)">
                                    <table width="510" height="369" border="0" align="center" cellpadding="0" cellspacing="0">
                                      <tr bgcolor="#ACD6FF"> 
                                        <td height="25" colspan="2"  class="unnamed1"><strong>Your       
                                          contact information</strong></td>      
                                      </tr>
                                      <tr> 
                                        <td height="25"  class="unnamed1">Name:</td>
                                        <td  class="unnamed1"><input name="contact_name" type="text" id="contact_name" size="20" value="<%if m=2 then response.Write "" else response.Write rs("firstname")&" "&rs("lastname") end if%>"> 
                                          <font color="#FF0000">*</font></td>
                                      </tr>
                                      <tr> 
                                        <td height="25"  class="unnamed1">Tel: 
                                        </td>
                                        <td  class="unnamed1"><input name="contact_tel" type="text" id="contact_tel" value="<%if m=2 then response.Write "" else response.Write rs("tel") end if%>"> 
                                          <font color="#FF0000">*</font></td>
                                      </tr>
                                      <tr> 
                                        <td height="25"  class="unnamed1">Fax:</td>
                                        <td  class="unnamed1"><input name="fax" type="text" id="fax" value="<%if m=2 then response.Write "" else response.Write rs("fax") end if%>"></td>
                                      </tr>
                                      <tr> 
                                        <td height="25"  class="unnamed1">E-mail:</td>
                                        <td  class="unnamed1"><input name="email" type="text" id="email" value="<%if m=2 then response.Write "" else response.Write rs("email") end if%>"> 
                                          <font color="#FF0000">*</font></td>
                                      </tr>
                                      <tr bgcolor="#B9FFFF"> 
                                        <td height="24" colspan="2"  class="unnamed1"><strong>Demand:</strong></td>
                                      </tr>
                                      <tr> 
                                        <td width="153" height="25"  class="unnamed1">Category:</td>
                                        <td width="357"  class="unnamed1"><FONT color=#000000> 
                                          <select name="category" id="category">
                                            <option>Transportation </option>
                                            <option>Construction &amp; Hardware</option>
                                            <option>Medical &amp; Health care</option>
                                            <option>Home Appliance </option>
                                            <option>Gifts, Models &amp; Crafts</option>
                                            <option>Sports &amp; Leisure </option>
                                            <option>Security &amp; Protection 
                                            </option>
                                            <option>Office supplies </option>
                                            <option>Electronics &amp; Electrical 
                                            </option>
                                            <option>Sourcing Services</option>
                                            <option>New Promotional Items</option>
                                            <option>Others</option>
                                          </select>
                                          </FONT></td>
                                      </tr>
                                      <tr> 
                                        <td height="25"  class="unnamed1">Product 
                                          Name:</td>
                                        <td> <input name="pro_name" type="text" id="pro_name4" size="30"> 
                                          <span class="unnamed1"><font color="#FF0000">*</font></span></td>
                                      </tr>
                                      <tr> 
                                        <td height="25"  class="unnamed1">Order 
                                          Volume (pcs):</td>
                                        <td> <input name="ordervolume" type="text" id="price4" size="30" class="unnamed1"></td>
                                      </tr>
                                      <tr> 
                                        <td height="25"  class="unnamed1">Expected 
                                          delivery time:</td>
                                        <td> <input name="ed_time" type="text" id="ed_time" size="30"></td>
                                      </tr>
                                      <tr> 
                                        <td height="90"  class="unnamed1">Notes:</td>
                                        <td> <textarea name="notes" cols="50" rows="5" id="textarea3"></textarea></td>
                                      </tr>
                                     
                                      <tr> 
                                        <td>&nbsp;</td>
                                        <td height="30"> <input type="submit" name="Submit" value="Submit"> 
                                          &nbsp;&nbsp;&nbsp; <input type="reset" name="Submit2" value="Reset" onClick="javascript:document.form2.reset();return false;"> 
                                        </td>
                                      </tr>
                                    </table>
                                  </form></td>
                              </tr>
                            </table></TD>
                        </TR>
                      </TBODY>
                    </TABLE></td>
                </tr>
                <tr> 
                  <td height="9" valign="top">&nbsp;</td>
                </tr>
                <tr> 
                  <td>&nbsp;</td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
  </table>
  <table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td height="1" valign="top" bgcolor="#333333"><img width="2" height="5"></td>
    </tr>
    <tr>
      <td height="25" bgcolor="#999999" class="unnamed1"><div align="center"><a href="http://www.sourcing-universal.com" target="_blank">Sourcing    
          Request</a> | <a href="Becomeoursupplier.asp">Become Our Suppliers</a>    
          | <a href="career.asp">Career</a> | <a href="privacypolicy.asp">Privacy    
          Policy</a></div></td>
    </tr>
    <tr> 
      <td height="25" bgcolor="#999999"> <div align="center" class="unnamed1"><font color="#000000" face="Arial, Helvetica, sans-serif">Copyright    
          &copy; 2006 SICTEC Group Inc. All rights reserved.</font></div></td>              
    </tr>
  </table>
</div>
<map name="Map">
  <area shape="rect" coords="17,6,71,30" href="index.asp">
  <area shape="rect" coords="104,4,192,30" href="aboutus.asp">
  <area shape="rect" coords="223,5,358,31" href="index.asp">
  <area shape="rect" coords="387,9,472,31" href="demand.asp">
  <area shape="rect" coords="495,7,586,31" href="contactus.asp">
</map>
<script language="JScript">
<!--
function check(form2)
{
if (form2.contact_name.value=="")
 {
    alert("Pls input your name,thanks!")
	form2.contact_name.focus()
    return false
  }
if (form2.contact_tel.value=="")
 {
    alert("Pls give us your telephone number,thanks!")
	form2.contact_tel.focus()
    return false
  }
if (form2.email.value=="")
 {
    alert("Pls input your email address,thanks!")
	form2.email.focus()
    return false
  }
if (form2.pro_name.value=="")
 {
    alert("Pls input product name,thanks!")
	form2.pro_name.focus()
    return false
	}
}
//-->
</script>
</body>
</html>
