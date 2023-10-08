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
.unnamed3 {	font-family: "Arial", "Helvetica", "sans-serif";
	font-size: 11px;
}
-->
</style>
</head>

<body background="file:///E|/website/tire%20clock/image/bg01.jpg" text="#000000"   link="#000000" vlink="#000000" alink="#000000" topmargin="0" >
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
                  <td height="34" valign="top"><img src="image/banner06.jpg" width="599" height="34" border="0" usemap="#Map"></td>
                </tr>
                <tr> 
                  <td height="138" valign="top"><img src="image/pho18.jpg" width="599" height="139"></td>
                </tr>
                <tr> 
                  <td height="25" valign="top" background="image/bg12.jpg"><table width="95%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="6%" height="23">&nbsp;</td>
                        <td width="94%" height="35" valign="bottom"><img src="image/career.jpg" width="127" height="19"></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="9" valign="top"><TABLE height=180 cellSpacing=0 cellPadding=0 width="99%" bgColor=#ffffff 
      border=0>
                      <TBODY>
                        <TR> 
                          <TD width=562 height=24>&nbsp;</TD>
                        </TR>
                        <TR> 
                          <TD vAlign=top width=562 height=273><table width="100%" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td width="10%">&nbsp;</td>
                                <td width="82%" class="unnamed2">SICTEC Group 
                                  Inc. is a minority-owned equal opportunity employer，with 
                                  headquarter located in Toronto, Canada and office 
                                  in Beijing, China. Our team works closely with 
                                  our clients and partners to deliver products, 
                                  programs and services that surpass the highest 
                                  standards. SICTEC welcomes highly motivated, 
                                  energetic people who are interested in professional 
                                  growth and career development to contact us 
                                  at hr@sictec.com.</td>
                                <td width="8%">&nbsp;</td>
                              </tr>
                            </table> </TD>
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
  <area shape="rect" coords="19,12,69,26" href="index.asp">
  <area shape="rect" coords="110,10,187,27" href="contactus.asp">
  <area shape="rect" coords="220,13,365,27" href="index.asp">
  <area shape="rect" coords="393,11,473,27" href="demand.asp">
  <area shape="rect" coords="493,11,585,27" href="contactus.asp">
</map>
<script language="JScript">
<!--
function check(form1)
{
if (form1.subject.value=="")
 {
    alert("Pls input your name,thanks!")
	form1.subject.focus()
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
