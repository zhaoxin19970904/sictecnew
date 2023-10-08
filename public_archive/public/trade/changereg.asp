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

      <%
					    username=trim(session("username"))
						if username<>"" then
						  sql="select * from member where member_uid='"&username&"'"
						  set rs=conn.execute(sql)
						else %>
						<script language="JavaScript">
alert("Time out.You have ben logout,pls login first!")
history.back()
</script>
						
						<%  
						response.end()
						end if
						%>


<body background="image/bg01.jpg" text="#000000"   link="#000000" vlink="#000000" alink="#000000" topmargin="0" >
<div align="left">
  <table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td bgcolor="#FFFFFF"> <table width="800" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999">
          <tr> 
            <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="25" valign="top" bgcolor="#FFFFFF"> 
                    <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
                      <tr> 
                        <td width="25%" valign="top" bgcolor="#FFFFFF"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td><img src="image/logo.jpg" width="201" height="66"></td>
                            </tr>
                            <tr> 
                              <td height="30" bgcolor="#D9E2E7" class="unnamed2"> 
                                <div align="center"><font color="#000033"><strong>MEMBER 
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
                %>
                              </td>
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
                              <td height="34" valign="top"><img src="image/pho14.jpg" width="599" height="34" border="0" usemap="#MapMapMapMap"></td>
                            </tr>
                            <tr> 
                              <td height="138" valign="top"><img src="image/pho20.jpg" width="599" height="139"></td>
                            </tr>
                          </table></td>
                      </tr>
                    </table>
                    <map name="MapMapMapMap">
                      <area shape="rect" coords="19,12,69,26" href="index.asp">
                      <area shape="rect" coords="110,10,187,27" href="aboutus.asp">
                      <area shape="rect" coords="220,13,365,27" href="index.asp">
                      <area shape="rect" coords="393,11,473,27" href="demand.asp">
                      <area shape="rect" coords="493,11,585,27" href="contactus.asp">
                    </map>
                    <table width="100%" border="0" cellpadding="0" cellspacing="0" background="image/bg12.jpg">
                      <tr> 
                        <td width="6%" height="23">&nbsp;</td>
                        <td width="94%" height="20" valign="bottom">&nbsp;</td>
                      </tr>
                    </table> </td>
                </tr>
                <tr> 
                  <td height="45" valign="top" bgcolor="#FFFFFF"> 
                    <div align="center"> 
                      <map name="MapMapMap">
                        <area shape="rect" coords="19,12,69,26" href="index.asp">
                        <area shape="rect" coords="110,10,187,27" href="aboutus.asp">
                        <area shape="rect" coords="220,13,365,27" href="index.asp">
                        <area shape="rect" coords="393,11,473,27" href="demand.asp">
                        <area shape="rect" coords="493,11,585,27" href="contactus.asp">
                      </map>
                      <span class="unnamed2"><FONT color=#31659c><B><font color="#244791" size="3">Register                        
                      as Sictec member make deal online with ease!</font></B></FONT><FONT color=#31659c size=4><B><BR>                       
                      </B></FONT><FONT color=#31659c>( <font color="#FF3300">*</font> 
                      update my personal information!)</FONT></span></div></td>
                </tr>
                <tr> 
                  <td bgcolor="#FFFFFF"><TABLE width="98%" height=930 border=0 align="center" cellPadding=0 cellSpacing=0 
      valign="top">
                      <TBODY>
                        <TR> 
                          <TD vAlign=top align=middle width="100%" height=930> 
                            <form name="form2" method="post" action="script04/mail2.asp" onSubmit="return check(this)">
                              <TABLE cellSpacing=0 cellPadding=0 width=700 border=0>
                                <TBODY>
                                  <TR> 
                                    <TD bgColor=#C4E1FF height=24>&nbsp;<img src="image/kp.gif" width="11" height="8"><B> 
                                      <span class="unnamed2">Member Information</span></B></TD>                       
                                  </TR>
                                </TBODY>
                              </TABLE>
                              <BR>
                              <TABLE width="92%" 
border=0 align=center cellPadding=4 cellSpacing=1 class="unnamed1">
                                <TBODY>
                                  <TR> 
                                    <TD width="24%" height=28 align=right> <B>Member ID:</B><BR></TD> 
                                    <TD width="76%" height=28 vAlign=top class="unnamed1">  &nbsp;&nbsp;<%= username %> </TD>
                                  </TR>
                                  <TR> 
                                    <TD align=right height=28> <b>Your former              
                                      password:</b><BR></TD>
                                    <TD height=28> <INPUT class=fonten id=psd3 type=password 
                  size=30 name=psd_f> (If you want to change your password, pls   
                                      input your former password. If not, pls   
                                      keep this item in blank. ) </TD>  
                                  </TR>
                                  <TR> 
                                    <TD align=right height=28> <B> 
                                      &nbsp;New&nbsp; password:</B><BR></TD>             
                                    <TD height=28><INPUT class=fonten id=psd2 type=password 
                  size=30 name=psd> </TD>
                                  </TR>
                                  <TR> 
                                    <TD align=right height=28><B>Re-enter new password:</B><BR></TD>  
                                    <TD height=28><INPUT class=fonten id=re_psd2 type=password 
                  size=30 name=re_psd></TD>
                                  </TR>
                                </TBODY>
                              </TABLE>
                              <TABLE cellSpacing=0 cellPadding=0 width=700 border=0>
                                <TBODY>
                                  <TR> 
                                    <TD bgColor=#C4E1FF height=25>&nbsp;<img src="image/kp.gif" width="11" height="8"> 
                                      <span class="unnamed2"><B>Contact information</B>                        
                                      </span></TD>
                                  </TR>
                                </TBODY>
                              </TABLE>
                              <BR>
                              <TABLE width="92%" 
border=0 align=center cellPadding=4 cellSpacing=1 class="unnamed1">
                                <TBODY>
                                  <TR> 
                                    <TD rowspan="2" align=right><FONT color=#ff3300>*</FONT><STRONG>Name: 
                                      <B> </B></STRONG></TD>
                                    <TD> <table width="200" height="12" border="0" cellpadding="0" cellspacing="0" class="unnamed1">
                                        <tr> 
                                          <td width="100" height="12"> <div align="center">First 
                                              Name</div></td>
                                          <td width="100"><div align="center">Last 
                                              Name</div></td>
                                        </tr>
                                      </table></TD>
                                  </TR>
                                  <TR> 
                                    <TD height=29> <table width="200" height="24" border="0" cellpadding="0" cellspacing="0">
                                        <tr> 
                                          <td width="100" height="24"><div align="center"> 
                                              <input name=firstname class=fonten id=firstname size="8" value="<%=rs("firstname")%>">
                                            </div></td>
                                          <td width="100"><div align="center"> 
                                              <input name=lastname class=fonten id=lastname size="8" value="<%=rs("lastname")%>">
                                            </div></td>
                                        </tr>
                                      </table></TD>
                                  </TR>
                                  <TR> 
                                    <TD align=right height=28><STRONG>Job title:<B></B></STRONG></TD>
                                    <TD height=28><INPUT class=fonten id=jobtitle2 name=jobtitle value="<%=rs("jobtitle")%>"> 
                                    </TD>
                                  </TR>
                                  <TR> 
                                    <TD align=right height=28><STRONG>Contact 
                                      tel:<B></B></STRONG></TD>
                                    <TD height=28><INPUT class=fonten id=tel2 
                name=tel value="<%=rs("tel")%>"></TD>
                                  </TR>
                                  <TR> 
                                    <TD align=right height=28><FONT color=#ff3300>*</FONT> 
                                      <B>E-Mail:</B></TD>
                                    <TD height=28> <INPUT class=fonten id=email2 size=45 
                name=email value="<%=rs("email")%>">
                                      (Correct E-mail helps find password) </TD>
                                  </TR>
                                  <TR> 
                                    <TD align=right width="24%" height=28><B>Country/region:</B></TD>
                                    <TD width="76%" height=28> <input name="country" type="text" id="country" size="35" value="<%=rs("country")%>"> 
                                    </TD>
                                  </TR>
                                  <TR> 
                                    <TD align=right width="24%" height=28><B>City:</B></TD>
                                    <TD width="76%" height=28> <input name="city" type="text" id="city" size="35" value="<%=rs("city")%>"> 
                                    </TD>
                                  </TR>
                                  <TR> 
                                    <TD align=right width="24%" height=28><B>Address:</B></TD>
                                    <TD width="76%" height=28><INPUT class=fonten 
                  id=address2 size=55 name=address value="<%=rs("address")%>"></TD>
                                  </TR>
                                  <TR> 
                                    <TD align=right width="24%" height=28><B>Postal 
                                      code:</B></TD>
                                    <TD width="76%" height=28><INPUT class=fonten 
                  id=post2 size=40 name=post value="<%=rs("post")%>"></TD>
                                  </TR>
                                </TBODY>
                              </TABLE>
                              <BR>
                              <TABLE cellSpacing=0 cellPadding=0 width=700 border=0>
                                <TBODY>
                                  <TR> 
                                    <TD bgColor=#C4E1FF height=23>&nbsp;<img src="image/kp.gif" width="11" height="8"> 
                                      <span class="unnamed2"><B>Company information</B></span></TD>                       
                                  </TR>
                                </TBODY>
                              </TABLE>
                              <BR>
                              <TABLE width="92%" 
border=0 align=center cellPadding=4 cellSpacing=1 class="unnamed1">
                                <TBODY>
                                  <TR> 
                                    <TD align=right width="24%" height=28><div align="right"><FONT 
                  color=#ff0000>*</FONT><B>&nbsp;Company name:</B></div></TD>                       
                                    <TD width="76%" colSpan=2 height=28> <INPUT class=fonten id=company2 size=45 name=company value="<%=rs("company")%>"> 
                                    </TD>
                                  </TR>
                                  <TR> 
                                    <TD vAlign=top align=right height=28><div align="right"><FONT 
                  color=#ff0000>*</FONT><B>Major business area:</B><BR>
                                      </div></TD>
                                    <TD height=28> <input name="area" type="text" id="area" size="45" value="<%=rs("main_biz_area")%>"> 
                                    </TD>
                                  </TR>
                                  <TR> 
                                    <TD align=right><div align="right"><B>&nbsp;Main                        
                                        products/services:</B></div></TD>
                                    <TD> <P> 
                                        <TEXTAREA name=main_products cols=50 rows=5 class=fonten id="textarea" ><%=rs("main_products")%></TEXTAREA>
                                      </P></TD>
                                  </TR>
                                  <TR> 
                                    <TD align=right width="24%" height=28><div align="right"><FONT 
                  color=#ff0000>*</FONT><B>&nbsp;Business nature:</B></div></TD>                       
                                    <TD width="76%" colSpan=2 height=28> <select name="biznature" id="select3">
                                        <OPTION value="" selected>==select==</OPTION>
                                        <option>Importer/Exporter</option>
                                        <option>Brand Manufacturer</option>
                                        <option>Buying Office</option>
                                        <option>Manufacturer</option>
                                        <option>Trading Company</option>
                                        <option>Service Company</option>
                                        <option>Chain Store</option>
                                        <option>Wholesaler/Distributor/Agent</option>
                                        <option>Retailer</option>
                                        <option>Contractor</option>
                                      </select> </TD>
                                  </TR>
                                  <TR> 
                                    <TD align=right height=28><div align="right"><FONT 
                  color=#ff0000>*</FONT><B><B>&nbsp;</B>Tel:</B></div></TD>
                                    <TD colSpan=2 height=28><FONT class=fpt13> 
                                      <INPUT class=fonten 
                  id=com_tel2 name=com_tel value="<%=rs("com_tel")%>">
                                      </FONT></TD>
                                  </TR>
                                  <TR> 
                                    <TD align=right height=28><div align="right"><STRONG>Fax:</STRONG></div></TD>
                                    <TD colSpan=2 height=28><INPUT class=fonten id=fax2 
              name=fax value="<%=rs("fax")%>"></TD>
                                  </TR>
                                  <TR> 
                                    <TD height=28> <div align="right"><B><B></B>Products 
                                        to purchase:</B></div></TD>
                                    <TD height=28><textarea name="productbuy" cols="50" rows="5" id="textarea2" value=""><%=rs("productbuy")%></textarea> 
                                    </TD>
                                  </TR>
                                  <TR> 
                                    <TD align=right width="24%" height=28><div align="right"><B>&nbsp;Products                        
                                        to sell:</B></div></TD>                       
                                    <TD width="76%" height=28><textarea name="productsell" cols="50" rows="5" id="textarea3" value=""><%=rs("productsell")%></textarea></TD>
                                  </TR>
                                  <TR> 
                                    <TD align=right width="24%" 
                height=28><div align="right"><STRONG>Homepage:</STRONG></div></TD>
                                    <TD width="76%" colSpan=2 height=28><INPUT name=homepage class=fonten 
                  id=homepage2 size="35" value="<%=rs("homepage")%>"> </TD>
                                  </TR>
                                  <TR> 
                                    <TD align=right width="24%" height=28>&nbsp;</TD>
                                    <TD width="76%" colSpan=2 height=28><INPUT id=module2 
                  type=hidden value=commodule.phtml name=module> </TD>
                                  </TR>
                                </TBODY>
                              </TABLE>
                              <DIV align=center> 
                                <TABLE cellSpacing=0 cellPadding=0 width="92%" border=0>
                                  <TBODY>
                                    <TR> 
                                      <TD align=middle> <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                                          <TBODY>
                                            <TR> 
                                              <TD align=right width="51%"><input type="submit" name="Submit2" value="Update"> 
                                              </TD>
                                              <TD width="5%">&nbsp;</TD>
                                              <TD align=left width="44%"><input type="reset" name="Submit22" value="Reset" onClick="javascript:document.form2.reset();return false;"></TD>
                                            </TR>
                                          </TBODY>
                                        </TABLE></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE>
                                <BR>
                                &nbsp;&nbsp; &nbsp; </DIV>
                            </FORM></TD>
                        </TR>
                      </TBODY>
                    </TABLE></td>
                </tr>
              </table></td>
          </tr>
        </table>
        <div align="center"> 
          <map name="MapMap">
            <area shape="rect" coords="19,12,69,26" href="explain.asp">
            <area shape="rect" coords="110,10,187,27" href="explain.asp">
            <area shape="rect" coords="220,13,365,27" href="explain.asp">
            <area shape="rect" coords="393,11,473,27" href="explain.asp">
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
if (form1.member_id.value=="")
 {
    alert("Pls input member ID,thanks!")
	form1.member_id.focus()
    return false
  }

if (form1.re_psd.value=="")
 {
    alert("Pls input re-enter password,thanks!")
	form1.re_psd.focus()
    return false
  }
if (form1.firstname.value=="")
 {
    alert("Pls input your firstname,thanks!")
	form1.firstname.focus()
    return false
	}
if (form1.lastname.value=="")
 {
    alert("Pls input your lastname,thanks!")
	form1.lastname.focus()
    return false
	}
if (form1.email.value=="")
 {
    alert("Pls input your email address,thanks!")
	form1.email.focus()
    return false
	}
if (form1.company.value=="")
 {
    alert("Pls input company name,thanks!")
	form1.company.focus()
    return false
	}
if (form1.area.value=="")
 {
    alert("Pls input major business area,thanks!")
	form1.area.focus()
    return false
	}
if (form1.com_tel.value=="")
 {
    alert("Pls give us your telephone number,thanks!")
	form1.com_tel.focus()
    return false
	}
}
//-->
</script>
</body>
</html>
