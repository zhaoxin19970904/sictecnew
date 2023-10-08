<%@ LANGUAGE="VBSCRIPT" CODEPAGE="936" EnableSessionState=True %>
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
	line-height: 20px;
}
-->
</style>
</head>

<body background="file:///E|/website/tire%20clock/image/bg01.jpg" text="#000000"   link="#000000" vlink="#000000" alink="#000000" topmargin="0" >
<div align="left">
  <table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td valign="top" bgcolor="#FFFFFF"> <table width="800" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999">
          <tr> 
            <td valign="top" bgcolor="#FFFFFF">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td><table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
                      <tr> 
                        <td width="25%" height="259" valign="top" bgcolor="#FFFFFF"> 
                          <table width="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td><img src="image/logo.jpg" width="201" height="66"></td>
                            </tr>
                            <tr> 
                              <td height="30" bgcolor="#D9E2E7" class="unnamed2"> 
                                <div align="center"><font color="#000033"><strong>MEMBER 
                                  LOGIN</strong></font></div></td>
                            </tr>
                            <tr> 
                              <td> <%
                username=trim(session("username"))
                if username="" then%> <table border="0" cellspacing="0" cellpadding="0">
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
                                              <td height=20 colspan=2 valign=bottom> 
                                                <div align=center> <a href="script/login.asp"> 
                                                  <INPUT name="image" type=image 
                  onclick="return checkempty(form1)" src="star/new_pa1.gif" 
                  border=0>
                                                  </a>&nbsp; </div></td>
                                            </tr>
                                            <tr> 
                                              <td height=17 colspan=2 valign=top> 
                                                <div align=center class="unnamed1"><b><a href="lost_p_w.asp"><font color="#000000">Forgot</font></a> 
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
                                              <td height=20 colspan=2 valign=bottom> 
                                                <div align=center class="unnamed1"> 
                                                  <b> <a href="changereg.asp"><font color="#000000"> 
                                                  &nbsp;Update my profile</font></a> 
                                                  </b></div></td>
                                            </tr>
                                            <tr> 
                                              <td height=17 colspan=2 valign=top> 
                                                <div align=center class="unnamed1"><b><a href="script/logout.asp"><font color="#000000">Exit</font></a> 
                                                  </b> </div></td>
                                            </tr>
                                          </tbody>
                                        </table>
                                      </form></td>
                                  </tr>
                                </table>
                                <% 
                              end if
                %> </td>
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
                                      Join now</font></a><font color="#666666">&nbsp; 
                                      </font></td>
                                  </tr>
                                </table></td>
                            </tr>
                            <tr> 
                              <td height="37" background="image/bg04.jpg">&nbsp;</td>
                            </tr>
                            <tr> 
                              <td height="34" valign="top"><img src="image/banner05.jpg" width="599" height="34" border="0" usemap="#MapMapMap"></td>
                            </tr>
                            <tr> 
                              <td height="138" valign="top"><img src="image/pho17.jpg" width="599" height="139"></td>
                            </tr>
                            <tr> 
                              <td height="23" valign="top" background="image/bg12.jpg">
<table width="95%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td width="6%" height="23">&nbsp;</td>
                                    <td width="94%" height="20" valign="bottom">&nbsp;</td>
                                  </tr>
                                </table></td>
                            </tr>
                          </table></td>
                      </tr>
                    </table>
                    <map name="MapMapMap">
                      <area shape="rect" coords="19,12,69,26" href="index.asp">
                      <area shape="rect" coords="110,10,187,27" href="aboutus.asp">
                      <area shape="rect" coords="219,13,364,27" href="index.asp">
                      <area shape="rect" coords="393,11,473,27" href="demand.asp">
                      <area shape="rect" coords="493,11,585,27" href="contactus.asp">
                    </map> 
                    <table width="85%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td height="30" valign="top" class="unnamed1"><img src="image/BecomeOurSuppliers.jpg" width="230" height="19"> 
                          <hr size="1"></td>
                      </tr>
                      <tr> 
                        <td valign="top" class="unnamed1"> 
						<form name="form1" method="post" action="script02/mail.asp" onSubmit="return check(this)">
                            <table width="100%" border="0" cellspacing="1" cellpadding="0">
                              <tr> 
                                <td><table width="100%" border="0" cellspacing="1" cellpadding="0">
                                    <tr> 
                                      <td width="25%" height="28" class="unnamed2"> 
                                        <div align="right">Company Name: </div></td>
                                      <td width="75%"><input name="company" type="text" id="company" size="40">
                                        <span class="unnamed2"><font color="#FF0000">*</font></span></td>
                                    </tr>
                                  </table></td>
                              </tr>
                              <tr> 
                                <td height="30" bgcolor="#BEDEDE" class="unnamed2"><strong><font color="#003333">CONTACT 
                                  INFORMATION</font></strong></td>
                              </tr>
                              <tr> 
                                <td><table width="100%" border="0" cellspacing="1" cellpadding="0">
                                    <tr bgcolor="#FFFFFF"> 
                                      <td height="25" class="unnamed2"> 
                                        <div align="right">Web site:</div></td>
                                      <td height="25"> 
                                        <input name="website" type="text" id="website" size="40"></td>
                                    </tr>
                                    <tr bgcolor="#FFFFFF"> 
                                      <td height="25" class="unnamed2"> 
                                        <div align="right">Adress:</div></td>
                                      <td height="25"> 
                                        <input name="adress" type="text" id="adress" size="50"></td>
                                    </tr>
                                    <tr bgcolor="#FFFFFF"> 
                                      <td height="25" class="unnamed2"> 
                                        <div align="right">Telephone:</div></td>
                                      <td height="25"> 
                                        <input name="tel" type="text" id="tel">
                                        <font color="#FF0000"><span class="unnamed2">* 
                                        </span></font></td>
                                    </tr>
                                    <tr bgcolor="#FFFFFF"> 
                                      <td height="25" class="unnamed2"> 
                                        <div align="right">Fax:</div></td>
                                      <td height="25"> 
                                        <input name="fax" type="text" id="fax"></td>
                                    </tr>
                                    <tr bgcolor="#FFFFFF"> 
                                      <td height="25" class="unnamed2"> 
                                        <div align="right">Contact 
                                          Person:</div></td>
                                      <td height="25"> 
                                        <input name="contactperson" type="text" id="contactperson"></td>
                                    </tr>
                                    <tr bgcolor="#FFFFFF"> 
                                      <td height="25" class="unnamed2"> 
                                        <div align="right">Position:</div></td>
                                      <td height="25"> 
                                        <input name="position" type="text" id="position"></td>
                                    </tr>
                                    <tr bgcolor="#FFFFFF"> 
                                      <td height="25" class="unnamed2"> 
                                        <div align="right">Mobile:</div></td>
                                      <td height="25"> 
                                        <input name="mobile" type="text" id="mobile"></td>
                                    </tr>
                                    <tr bgcolor="#FFFFFF"> 
                                      <td width="25%" height="25" class="unnamed2"> 
                                        <div align="right">E-mail:</div></td>
                                      <td width="75%" height="25"> 
                                        <input name="email" type="text" id="email" size="30">
                                        <font color="#FF0000"><span class="unnamed2">* 
                                        </span></font></td>
                                    </tr>
                                  </table></td>
                              </tr>
                              <tr> 
                                <td height="30" bgcolor="#BEDEDE" class="unnamed2"><strong><font color="#003333">COMPANY 
                                  PROFILE</font></strong></td>
                              </tr>
                              <tr> 
                                <td><table width="100%" border="0" cellspacing="1" cellpadding="0">
                                    <tr bgcolor="#FFFFFF"> 
                                      <td height="25" class="unnamed2"> <div align="right">Number 
                                          of employees:</div></td>
                                      <td><input name="no_employees" type="text" id="no_employees"></td>
                                    </tr>
                                    <tr bgcolor="#FFFFFF"> 
                                      <td class="unnamed2"> <div align="right">Mian 
                                          products:</div></td>
                                      <td><textarea name="mianproducts" cols="48" rows="5" id="mianproducts"></textarea>
                                        <font color="#FF0000">&nbsp;</font></td>
                                    </tr>
                                    <tr bgcolor="#FFFFFF"> 
                                      <td class="unnamed2"> <div align="right">Annual 
                                          production:</div></td>
                                      <td><textarea name="annualproduction" cols="48" rows="5" id="annualproduction"></textarea></td>
                                    </tr>
                                    <tr bgcolor="#FFFFFF"> 
                                      <td class="unnamed2"> <div align="right">Export 
                                          history:</div></td>
                                      <td><textarea name="exporthistory" cols="48" rows="5" id="exporthistory"></textarea>
                                        <font color="#FF0000">&nbsp;</font></td>
                                    </tr>
                                    <tr bgcolor="#FFFFFF"> 
                                      <td class="unnamed2"> <div align="right">Major 
                                          Export Area:</div></td>
                                      <td><textarea name="area" cols="48" rows="5" id="area"></textarea>
                                      </td>
                                    </tr>
                                    <tr bgcolor="#FFFFFF"> 
                                      <td height="25" class="unnamed2"> <div align="right">ISO 
                                          registered:</div></td>
                                      <td><input name="registered" type="text" id="registered" size="50">
                                        <font color="#FF0000">&nbsp;</font></td>
                                    </tr>
                                    <tr bgcolor="#FFFFFF"> 
                                      <td class="unnamed2"> <div align="right">Certificates/Approvals:</div></td>
                                      <td><textarea name="certificate" cols="48" rows="5" id="certificate"></textarea>
                                        <span class="unnamed2"> </span></td>
                                    </tr>
                                    <tr bgcolor="#FFFFFF"> 
                                      <td class="unnamed2"> <div align="right">Competitive 
                                          Strengh:</div></td>
                                      <td><textarea name="competitive" cols="48" rows="5" id="competitive"></textarea></td>
                                    </tr>
                                    <tr bgcolor="#FFFFFF">
                                      <td class="unnamed2"><div align="right">Notes/Comments:</div></td>
                                      <td><textarea name="notes" cols="48" rows="5" id="notes"></textarea> 
                                      </td>
                                    </tr>
                                    <tr bgcolor="#FFFFFF"> 
                                      <td width="25%" class="unnamed2"> <div align="right"></div></td>
                                      <td width="75%" height="35"> 　　　　　 
                                        <input type="submit" name="Submit" value="Submit">
                                        　 
                                        <input type="reset" name="Submit2" value="Reset" onClick="javascript:document.form1.reset();return false;"> 
                                      </td>
                                    </tr>
                                  </table></td>
                              </tr>
                            </table>
                          </form>
                          <p>&nbsp;</p></td>
                      </tr>
                    </table></td>
                </tr>
              </table>
              <font color="#FFFFCC">&nbsp;</font></td>
          </tr>
        </table>
        <div align="center"> 
          <map name="MapMap">
            <area shape="rect" coords="19,12,69,26" href="index.asp">
            <area shape="rect" coords="110,10,187,27" href="aboutus.asp">
            <area shape="rect" coords="219,13,364,27" href="index.asp">
            <area shape="rect" coords="393,11,473,27" href="request.asp">
            <area shape="rect" coords="493,11,585,27" href="contactus.asp">
          </map>
        </div></td>
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
  <area shape="rect" coords="19,12,69,26" href="explain.asp">
  <area shape="rect" coords="110,10,187,27" href="explain.asp">
  <area shape="rect" coords="220,13,365,27" href="explain.asp">
  <area shape="rect" coords="393,11,473,27" href="explain.asp">
  <area shape="rect" coords="493,11,585,27" href="contactus.asp">
</map>
<script language="JScript">
<!--
function check(form1)
{
if (form1.company.value=="")
 {
    alert("Pls input your company name,thanks!")
	form1.company.focus()
    return false
  }
if (form1.tel.value=="")
 {
    alert("Pls give us your telephone number,thanks!")
	form1.tel.focus()
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
