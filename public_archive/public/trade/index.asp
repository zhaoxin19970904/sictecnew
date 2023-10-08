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
}
-->
</style>
<style type="text/css">
<!--
.unnamed2 {
	font-family: "Arial", "Helvetica", "sans-serif";
	font-size: 14px;
}
.unnamed3 {
	font-family: "Arial", "Helvetica", "sans-serif";
	font-size: 11px;
}
-->
</style>
</head>


<body background="image/bg01.jpg" text="#000000"   link="#000000" vlink="#000000" alink="#000000" topmargin="0" >
<!--DESIGNER MEGGE -->
<div align="left">
<table width="800" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999">
    <tr>
      <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
          <tr>
            <td width="25%" height="318" valign="top" bgcolor="#D8D8D8">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td><img src="image/logo.jpg" width="201" height="66"></td>
                </tr>
                <tr> 
                  <td height="30" bgcolor="#D9E2E7" class="unnamed2"> 
                    <div align="center"><font color="#000033"><strong>MEMBER LOGIN</strong></font></div></td>
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
                                      <INPUT 
                  onclick="return checkempty(form1)" type=image src="star/new_pa1.gif" 
                  border=0>
                                      </a>&nbsp;  
                                    </div></td>
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
                                  <td height=37 valign="bottom" class="unnamed1"> <div align=left><%=firstname%>&nbsp;&nbsp;<%=lastname%></div></td>
                                </tr>
                                <tr> 
                                  <td width="41%" height=28 class="unnamed1" colspan="2"> 
                                    <div align=center><%=email%> </div></td>
                                </tr>
                                <tr> 
                                  <td height=20 colspan=2 valign=bottom> <div align=center class="unnamed1"> <b>
                                     <a href="changereg.asp"><font color="#000000"> &nbsp;Update my profile</font></a>                                            
                                     </b></div></td>
                                </tr>
                                <tr> 
                                  <td height=17 colspan=2 valign=top> <div align=center class="unnamed1"><b><a href="script/logout.asp"><font color="#000000">Exit</font></a>              
                                      </b> </div></td>             
                                </tr>
                              </tbody>
                            </table>
                          </form></td>
                      </tr>
                    </table>  <% 
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
                          cooker<br>
                          <strong><img src="image/b07.jpg" width="7" height="7"></strong>　vacuum                                       
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
                        <td height="16" valign="top" class="unnamed1"> 
                          <div align="center" class="unnamed3">3 Wheel Sliding 
                            Scooter<br>
                          </div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="18" bgcolor="#91B7BC" ><strong><font class="MB">　</font><span class="unnamed1">LINK</span></strong></td>
                </tr>
                <tr>
                  <td height="1" valign="top" bgcolor="#336666"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
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
                  <td height="37" background="image/bg04.jpg" align=right><font color="666666"> 
                    <font color="#999999">&nbsp;<span class="unnamed1"><%=date() %></span></font><font color="#666666"> </font>&nbsp;</font></td>
                </tr>
                <tr> 
                  <td height="34" valign="top"><img src="image/pho12.jpg" width="599" height="34" border="0" usemap="#Map"></td>
                </tr>
                <tr> 
                  <td height="138" valign="top"><img src="image/bg09.jpg" width="599" height="139"></td>
                </tr>
                <tr> 
                  <td height="25" valign="top" background="image/bg12.jpg"><table width="95%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="7%" height="26">&nbsp;</td>
                        <td width="93%" height="34" valign="bottom"><img src="image/bg11.jpg" width="174" height="17"></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="9" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="25" colspan="3"> 
                          <hr width="520" size="1"></td>
                      </tr>
                      <tr> 
                        <td width="7%">&nbsp;</td>
                        <td width="46%" height="25" valign="top"> <table width="88%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td width="40%" rowspan="2" class="unnamed1"> <div align="left"><a href="product/transportation.asp"><img src="image/pho01.jpg" width="82" height="65" border="0"></a></div>
                                <div align="left"></div></td>
                              <td width="60%" height="20" valign="bottom" class="unnamed1"><a href="product/transportation.asp" > 
                                <font color="#003366"><strong>Transportation</strong></font> 
                                </a></td>
                            </tr>
                            <tr> 
                              <td height="18" valign="top" class="unnamed1"><font color="#666666">Bicycles 
                                &amp; Parts,Auto Parts &amp; Equipment,etc.</font></td>
                            </tr>
                            <tr> 
                              <td height="25" colspan="2" class="unnamed1"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td background="image/d02.jpg"><img src="image/d02.jpg" width="2" height="1"></td>
                                  </tr>
                                </table></td>
                            </tr>
                            <tr> 
                              <td rowspan="2" class="unnamed1"> <div align="left"></div>
                                <a href="product/Construction&hardware.asp"><img src="image/pho02.jpg" width="82" height="65" border="0"></a></td>
                              <td height="18" valign="bottom" class="unnamed2"><span class="unnamed1"><strong><a href="product/Construction&hardware.asp"><font color="#003366">Construction 
                                &amp; Hardware</font></a></strong></span><strong><font color="#003366"> 
                                </font></strong></td>
                            </tr>
                            <tr> 
                              <td height="18" valign="top" class="unnamed1"><font color="#666666">Toilet 
                                Appliance &amp; Plumbing &amp; Garden</font><font color="#666666">,etc.</font></td>
                            </tr>
                            <tr> 
                              <td height="25" colspan="2" class="unnamed1"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td background="image/d02.jpg"><img src="image/d02.jpg" width="2" height="1"></td>
                                  </tr>
                                </table></td>
                            </tr>
                            <tr> 
                              <td rowspan="2" class="unnamed1"> <div align="left"></div>
                                <a href="product/Medical&Healthcare.asp"><img src="image/pho03.jpg" width="82" height="65" border="0"></a></td>
                              <td height="18" valign="bottom" class="unnamed2"><span class="unnamed1"><strong><a href="product/Medical&Healthcare.asp"><font color="#003366">Medical 
                                &amp; Health care</font></a></strong></span><font color="#003366"><a href="product/Medical&Healthcare.asp"><strong> 
                                </strong></a></font></td>
                            </tr>
                            <tr> 
                              <td height="18" valign="top" class="unnamed1"><font color="#666666">Recovered 
                                Excerciser,<br>
                                Medical Equipment etc.</font></td>
                            </tr>
                            <tr> 
                              <td height="25" colspan="2" class="unnamed1"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td background="image/d02.jpg"><img src="image/d02.jpg" width="2" height="1"></td>
                                  </tr>
                                </table></td>
                            </tr>
                            <tr> 
                              <td rowspan="2" class="unnamed1"> <div align="left"></div>
                                <a href="product/HomeAppliance.asp"><img src="image/pho04.jpg" width="82" height="65" border="0"></a></td>
                              <td height="18" valign="bottom" class="unnamed1"><strong><a href="product/HomeAppliance.asp"><font color="#003366">Home 
                                Appliance</font></a></strong> </td>
                            </tr>
                            <tr> 
                              <td height="18" valign="top" class="unnamed1"><font color="#666666">Kitchen 
                                appliance,<br>
                                Refrigerator,etc.</font></td>
                            </tr>
                            <tr> 
                              <td height="25" colspan="2" class="unnamed1"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td background="image/d02.jpg"><img src="image/d02.jpg" width="2" height="1"></td>
                                  </tr>
                                </table></td>
                            </tr>
                            <tr> 
                              <td rowspan="2" class="unnamed1"><a href="product/CraftsGifts.asp"><img src="image/pho05.jpg" width="82" height="65" border="0"></a></td>
                              <td height="18" valign="bottom" class="unnamed2"><span class="unnamed1"><strong><a href="product/CraftsGifts.asp"><font color="#003366">Gifts, 
                                Models &amp; Crafts</font></a></strong></span><font color="#003366"><a href="product/CraftsGifts.asp"><strong> 
                                </strong></a></font></td>
                            </tr>
                            <tr> 
                              <td height="18" valign="top" class="unnamed1"><font color="#666666">Holiday 
                                Gifts &amp; Crafts,<br>
                                etc.</font></td>
                            </tr>
                            <tr> 
                              <td height="25" colspan="2" class="unnamed1"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td background="image/d02.jpg"><img src="image/d02.jpg" width="2" height="1"></td>
                                  </tr>
                                </table></td>
                            </tr>
                            <tr> 
                              <td rowspan="2" class="unnamed1"> <div align="left"><a href="product/newitems.asp"><img src="image/pho23.jpg" width="82" height="65" border="0"></a></div></td>
                              <td height="19" valign="bottom" class="unnamed1"><strong><a href="product/newitems.asp"><font color="#003366">New 
                                Promotional Items</font></a></strong></td>
                            </tr>
                            <tr> 
                              <td height="50" valign="top" class="unnamed1"><font color="#666666">Moving 
                                Display Stand &amp; Scales, <span class="unnamed3"><a href="http://www.sictec.com/tyreclock" target="_blank"><font color="#666666">Tire 
                                Clock</font></a></span><font color="#666666">,</font>etc.</font></td>
                            </tr>
                          </table></td>
                        <td width="47%" height="25" valign="top"><table width="85%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td width="41%" rowspan="2" class="unnamed1"> <div align="left"><a href="product/Sports&Leisure.asp"><img src="image/pho07.jpg" width="82" height="65" border="0"></a></div>
                                <div align="left"></div></td>
                              <td width="59%" height="20" valign="bottom" class="unnamed1"><a href="product/Sports&Leisure.asp" > 
                                <font color="#003366"><strong>Sports &amp; Leisure 
                                </strong></font></a></td>
                            </tr>
                            <tr> 
                              <td height="18" valign="top" class="unnamed1"><font color="#666666">Entertainment,<br>
                                Sports,Scooters,etc.</font></td>
                            </tr>
                            <tr> 
                              <td height="25" colspan="2" class="unnamed1">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td background="image/d02.jpg"><img src="image/d02.jpg" width="2" height="1"></td>
                                  </tr>
                                </table></td>
                            </tr>
                            <tr> 
                              <td rowspan="2" valign="top" class="unnamed1"> 
                                <div align="left"></div>
                                <a href="product/security&protection.asp"><img src="image/pho08.jpg" width="82" height="65" border="0"></a></td>
                              <td height="18" valign="bottom" class="unnamed2"><span class="unnamed1"><strong><a href="product/security&protection.asp"><font color="#003366">Security 
                                &amp; Protection</font></a></strong></span><font color="#003366"><strong> 
                                </strong></font><strong></strong></td>
                            </tr>
                            <tr> 
                              <td height="18" valign="top" class="unnamed1"><font color="#666666">Locks 
                                &amp; Safety &amp; Safety,etc.</font></td>
                            </tr>
                            <tr> 
                              <td height="25" colspan="2" class="unnamed1">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td background="image/d02.jpg"><img src="image/d02.jpg" width="2" height="1"></td>
                                  </tr>
                                </table></td>
                            </tr>
                            <tr> 
                              <td rowspan="2" class="unnamed1"> <div align="left"><a href="product/OfficeSupplies.asp"><img src="image/pho11.jpg" width="82" height="65" border="0"></a></div>
                              </td>
                              <td height="18" valign="bottom" class="unnamed1"><strong><a href="product/OfficeSupplies.asp"><font color="#003366">Office 
                                supplies </font></a></strong></td>
                            </tr>
                            <tr> 
                              <td height="18" valign="top" class="unnamed1"><font color="#666666">Stationery,Paper,<br>
                                Consumable,etc.</font></td>
                            </tr>
                            <tr> 
                              <td height="25" colspan="2" class="unnamed1">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td background="image/d02.jpg"><img src="image/d02.jpg" width="2" height="1"></td>
                                  </tr>
                                </table></td>
                            </tr>
                            <tr> 
                              <td rowspan="2" class="unnamed1"> <div align="left"></div>
                                <a href="product/Electronics&Electrical.asp"><img src="image/pho10.jpg" width="82" height="65" border="0"></a></td>
                              <td height="18" valign="bottom" class="unnamed1"><strong><a href="product/Electronics&Electrical.asp"><font color="#003366">Electronics 
                                &amp; Electrical</font></a></strong><font color="#003366"><a href="product/Electronics&Electrical.asp"><strong> 
                                </strong></a></font></td>
                            </tr>
                            <tr> 
                              <td height="18" valign="top" class="unnamed1"><font color="#666666">Audio 
                                &amp; Video, <br>
                                Fixtures,etc.</font></td>
                            </tr>
                            <tr> 
                              <td height="25" colspan="2" class="unnamed1">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td background="image/d02.jpg"><img src="image/d02.jpg" width="2" height="1"></td>
                                  </tr>
                                </table></td>
                            </tr>
                            <tr> 
                              <td rowspan="2" class="unnamed1"><a href="product/Others.asp"><img src="image/pho06.jpg" width="82" height="65" border="0"></a></td>
                              <td height="18" valign="bottom" class="unnamed1"><strong><a href="product/Others.asp"><font color="#003366">Others</font></a></strong> 
                              </td>
                            </tr>
                            <tr> 
                              <td height="18" valign="top" class="unnamed1"><font color="#666666">Conference 
                                System,<br>
                                Wines &amp; Liquors,etc.</font></td>
                            </tr>
                            <tr> 
                              <td height="25" colspan="2" class="unnamed1">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td background="image/d02.jpg"><img src="image/d02.jpg" width="2" height="1"></td>
                                  </tr>
                                </table></td>
                            </tr>
                            <tr> 
                              <td rowspan="2" valign="top" class="unnamed1"> 
                                <div align="left"><a href="http://www.sourcing-universal.com" target="_blank"><img src="image/pho13.jpg" width="82" height="65" border="0"></a></div></td>
                              <td height="19" valign="bottom" class="unnamed1"><strong><a href="http://www.sourcing-universal.com" target="_blank"><font color="#003366">Sourcing 
                                Services</font></a></strong></td>
                            </tr>
                            <tr> 
                              <td height="18" valign="top" class="unnamed1"><font color="#666666">chemicals,industrial 
                                supplies,etc.</font></td>
                            </tr>
                          </table></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="21"> <hr width="530" size="1"> </td>
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
<SCRIPT language=JavaScript>
	function checkempty(myform)
	{
		if(myform.username.value == "")
		{
			alert("Please input your UserID!");
			myform.username.focus();
			return false;
		}
		
		if(myform.userpass.value == "")
		{
			alert("Please input your Password!");
			myform.userpass.focus();
			return false;
		}
		return true;
	}
</SCRIPT>

<map name="Map">
  <area shape="rect" coords="19,12,69,26" href="index.asp">
  <area shape="rect" coords="110,10,187,27" href="aboutus.asp">
  <area shape="rect" coords="220,13,365,27" href="index.asp">
  <area shape="rect" coords="393,11,473,27" href="demand.asp">
  <area shape="rect" coords="493,11,585,27" href="contactus.asp">
</map>
</body>
</html>
