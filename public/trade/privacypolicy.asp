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
}
-->
</style>
</head>

<body background="file:///E|/website/tire%20clock/image/bg01.jpg" text="#000000"   link="#000000" vlink="#000000" alink="#000000" topmargin="0" >
<div align="left"> 
  <table width="800" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999">
    <tr> 
      <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
                <tr> 
                  <td width="25%" valign="top" bgcolor="#FFFFFF"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
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
                                        <td height=20 colspan=2 valign=bottom> 
                                          <div align=center> <a href="script/login.asp"> 
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
                                        <td height=20 colspan=2 valign=bottom> 
                                          <div align=center class="unnamed1"> 
                                            <b> <a href="changereg.asp"><font color="#000000"> 
                                            &nbsp;Update my profile</font></a> 
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
                        <td height="34" valign="top"><img src="image/banner08.jpg" width="599" height="34" border="0" usemap="#MapMapMapMap"></td>
                      </tr>
                      <tr> 
                        <td height="138" valign="top"><img src="image/pho21.jpg" width="599" height="139"></td>
                      </tr>
                      <tr> 
                        <td height="25" valign="top" background="image/bg12.jpg"><table width="95%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td width="6%" height="20">&nbsp;</td>
                              <td width="94%" height="20" valign="bottom">&nbsp;</td>
                            </tr>
                          </table></td>
                      </tr>
                    </table></td>
                </tr>
              </table>
              <map name="MapMapMapMap">
                <area shape="rect" coords="19,12,69,26" href="index.asp">
                <area shape="rect" coords="110,10,187,27" href="aboutus.asp">
                <area shape="rect" coords="219,13,364,27" href="index.asp">
                <area shape="rect" coords="393,11,473,27" href="demand.asp">
                <area shape="rect" coords="493,11,585,27" href="contactus.asp">
              </map>
              　　 <img src="image/bg23.jpg" width="158" height="19"><br>
              <hr width="710" size="1"> </td>
          </tr>
          <tr> 
            <td height="22" bgcolor="#FFFFFF"> 
              <div align="center"> 
                <map name="MapMapMap">
                  <area shape="rect" coords="19,12,69,26" href="index.asp">
                  <area shape="rect" coords="110,10,187,27" href="aboutus.asp">
                  <area shape="rect" coords="219,13,364,27" href="index.asp">
                  <area shape="rect" coords="393,11,473,27" href="demand.asp">
                  <area shape="rect" coords="493,11,585,27" href="contactus.asp">
                </map>
                <font size="2"><b>PRIVACY POLICY (Updated on Aug 18, 2000) </b></font></div></td>
          </tr>
          <tr> 
            <td valign="top" bgcolor="#FFFFFF"><table width="88%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td height="16"> <div align="center"><font size="2"></font></div></td>
                </tr>
                <tr> 
                  <td class="unnamed1"> <p>This is the Privacy Policy governing 
                      your use of the TradeUniversal Site. By using this Site, 
                      you consent to our Privacy Policy set out below. All terms 
                      not defined in this document have the meanings ascribed 
                      to them in the Terms of Use Agreement between you and TradeUniversal 
                      which by use of this Site you agree to accept. </p>
                    <p><strong>1. <span class="unnamed2"><font color="#990000">The 
                      Information We Collect </font></span></strong></p>
                    <p>1.1 Registration Information. At the time you register 
                      to become a Registered User of the Site, you will be asked 
                      to fill out a registration form which requires you provide 
                      information such as your name, address, phone/fax number, 
                      email address and other personal information as well as 
                      information about your business (&quot;Registration Information&quot;). 
                    </p>
                    <p>1.2 Publishing Information. If you submit any information 
                      to TradeUniversal to be published on the Site through the 
                      publishing tools, including but not limited to Company Profile, 
                      Product Catalog, Trade Leads, TrustPass Profile and any 
                      discussion forum, then you are deemed to have given consent 
                      to the publication of such information (&quot;Publishing 
                      Information&quot;). </p>
                    <p>1.3 Statistical Information. In addition, we gather aggregate 
                      statistical information about our Site and Users, such as 
                      IP addresses, browser software, operating system, pages 
                      viewed, number of sessions and unique visitors, etc. (&quot;Statistical 
                      Information&quot;). </p>
                    <p>1.4 Registration Information, Publishing Information, Statistical 
                      Information and any information we may collect from you 
                      through the use of cookies (see Section 5 below) or any 
                      other means shall collectively be referred to as &quot;Collected 
                      Information&quot;. </p>
                    <p class="unnamed2"><strong>2. <font color="#990000">How We 
                      May Use Information</font> </strong></p>
                    <p>2.1 General. We use your Collected Information to improve 
                      our marketing and promotional efforts, to statistically 
                      analyze site usage, to improve our content and product offerings 
                      and to customize our Site's content, layout and service 
                      specifically for you. We may use your Collected Information 
                      to service your Account with us, including but not limited 
                      to investigating problems, resolving disputes and enforcing 
                      agreements with us. We do not sell, rent, trade or exchange 
                      any personally-identifying information of our Users. We 
                      may share certain aggregate information based on analysis 
                      of Collected Information with our partners, customers, advertisers 
                      or potential Users. We may use your Collected Information 
                      to execute marketing campaigns, promotions or advertising 
                      messages on behalf of third parties; however, in these circumstances, 
                      your Collected Information will not be disclosed to such 
                      third parties unless you respond to the marketing, promotion 
                      or advertising message. </p>
                    <p>2.2 Registration Information. We may use your Registration 
                      Information to provide services that you request or to contact 
                      you regarding additional services about which TradeUniversal 
                      determines that you might be interested. Specifically, we 
                      may use your email address, mailing address, phone number 
                      or fax number to contact you regarding notices, surveys, 
                      product alerts, new service or product offerings and communications 
                      relevant to your use of our Site. We may generate reports 
                      and analysis based on the Registration Information for internal 
                      analysis, monitoring and marketing decisions. </p>
                    <p>2.3 Publishing Information. All of your Publishing Information 
                      will be publicly available on the Site and therefore accessible 
                      by any internet user. Any Publishing Information that you 
                      disclose to TradeUniversal becomes public information and 
                      you relinquish any proprietary rights (including but not 
                      limited to the rights of confidentiality and copyrights) 
                      in such information. You should exercise caution when deciding 
                      to include personal or proprietary information in the Publishing 
                      Information that you submit to us. </p>
                    <p>2.4 Statistical Information. We use Statistical Information 
                      to help diagnose problems with and maintain our computer 
                      servers, to manage our Site, and to enhance our Site and 
                      services based on the usage pattern data we receive. We 
                      may generate reports and analysis based on the Statistical 
                      Information for internal analysis, monitoring and marketing 
                      decisions. We may provide Statistical Information to third 
                      parties, but when we do so, we do not provide personally-identifying 
                      information without your permission. </p>
                    <p class="unnamed2"><strong>3. <font color="#990000">Disclosure 
                      of Information</font> </strong></p>
                    <p>3.1 We reserve the right to disclose your Collected Information 
                      to relevant authorities where we have reason to believe 
                      that such disclosure is necessary to identify, contact or 
                      bring legal action against someone who may be infringing 
                      or threatening to infringe, or who may otherwise be causing 
                      injury to or interference with, the title, rights, interests 
                      or property of TradeUniversal, our Users, customers, partners, 
                      other web site users or anyone else who could be harmed 
                      by such activities. </p>
                    <p>3.2 We also reserve the right to disclose Collected Information 
                      in response to a subpoena or other judicial order or when 
                      we reasonably believe that such disclosure is required by 
                      law, regulation or administrative order of any court, governmental 
                      or regulatory authority. </p>
                    <p>3.3 If we have reason to believe that a User is in breach 
                      of the Terms of User Agreement or any other agreement with 
                      us, then we reserve the right to make public or otherwise 
                      disclose such User's Collected Information in order to pursue 
                      our claim or prevent further injury to TradeUniversal or 
                      others. </p>
                    <p><strong>4. <font color="#990000"><span class="unnamed2">Co-Branded 
                      Relationships</span></font> </strong></p>
                    <p>We have established relationships with other parties to 
                      offer you the benefit of other products and services which 
                      we ourselves do not offer. We offer you access to these 
                      other parties either through the use of hyperlinks to their 
                      sites from our Site, or through offering &quot;co-branded&quot; 
                      sites in which both ourselves and these other parties share 
                      the same uniform resource locator, domain name or pages 
                      within a domain name on the Internet. In some cases you 
                      may be required to submit information for purposes of registering 
                      or applying for products or services provided by such third 
                      parties or co-branded partners. The privacy policy of such 
                      other parties may differ from ours, and we may not have 
                      any control over the information that you submit to such 
                      third parties or co-branded sites. We therefore encourage 
                      you to read that policy before responding to any offers, 
                      products or services provided by such other parties. </p>
                    <p><strong>5. <span class="unnamed2"><font color="#990000">Cookies</font></span> 
                      </strong></p>
                    <p>5.1 We use &quot;cookies&quot; to store specific information 
                      about you and track your visits to our Site. It is not uncommon 
                      for web sites to use cookies to enhance identification of 
                      their users. A &quot;cookie&quot; is a small amount of data 
                      that is sent to your browser and stored on your computer's 
                      hard drive. A cookie can be sent to your computer's hard 
                      drive only if you access our Site using the computer. If 
                      you do not de-activate or erase the cookie, each time you 
                      use the same computer to access our Site, our web servers 
                      will be notified of your visit to our Site and in turn we 
                      may have knowledge of your visit and the pattern of your 
                      usage. Generally, we use cookies to identify you and enable 
                      us to access your Registration Information, Publishing Information 
                      or Payment Information so you do not have to re-enter it; 
                      gather statistical information about usage by Users; research 
                      visiting patters and help target advertisements based on 
                      User interests; assist our partners to track User visits 
                      to the Site and process orders; and track progress and participation 
                      in promotions. </p>
                    <p>5.2 You can determine if and how a cookie will be accepted 
                      by configuring your browser's which is installed in the 
                      computer you are using to access the Site. If you desire, 
                      you can change those configurations in your browser. By 
                      setting your preferences in the browser, you can accept 
                      all cookies, you can be notified when a cookie is sent, 
                      or you can reject all cookies. If you reject all cookies 
                      by choosing the cookie-disabling function in your browser, 
                      you may be required to re-enter your information on our 
                      Site more often and certain features of our Site may be 
                      unavailable. </p>
                    <p><strong>6. <span class="unnamed2"><font color="#990000">Minors</font></span></strong> 
                    </p>
                    <p>The Site and its contents are not intended to be targeted 
                      to minors under applicable law and we do not intend to sell 
                      any of our products or services to minors. However, we have 
                      no way of distinguishing the age of individuals who access 
                      our Site and so we carry out the same Privacy Policy for 
                      individuals of all ages. If a minor has provided us with 
                      personal information without parental or guardian consent, 
                      the parent or guardian should contact us to remove the information. 
                    </p>
                    <p><span class="unnamed2"><strong>7. <font color="#990000">Security 
                      Measures</font></strong></span><strong> </strong></p>
                    <p>7.1 We employ commercially reasonable security methods 
                      to prevent unauthorized access, maintain data accuracy and 
                      ensure correct use of information. </p>
                    <p>7.2 As a Registered User, your Registration Information, 
                      Publishing Information and Payment Information (if any) 
                      can be viewed and edited through your Account which is protected 
                      by Password. We recommend that you do not divulge your Password 
                      to anyone. Our personnel will never ask you for your Password 
                      in an unsolicited phone call or in an unsolicited e-mail. 
                      If you share a computer with others, you should not choose 
                      to save your log-in information (e.g., User ID and Password) 
                      on the computer. Remember to sign out of your Account and 
                      close your browser window when you have finished your session. 
                    </p>
                    <p>7.3 No data transmission over the Internet or any wireless 
                      network can be guaranteed to be perfectly secured. As a 
                      result, while we try to protect your information, no web 
                      site or company, including ourselves, can absolutely ensure 
                      or guarantee the security of any information you transmit 
                      to us and you do so at your own risk. </p>
                    <p class="unnamed2"><strong>8. <font color="#990000">Changes 
                      to Privacy Policy</font> </strong></p>
                    <p>Any changes to this Privacy Policy will be communicated 
                      through our posting an amended and restated Privacy Policy 
                      on our Site. Our posting the amended and restated Privacy 
                      Policy will make such new Privacy Policy immediately effective. 
                      You agree that all Collected Information (whether or not 
                      collected prior to or after the new policy became effective) 
                      will be governed by the newest Privacy Policy then in effect. 
                      If you do not agree to the new changes in our Privacy Policy, 
                      you should contact TradeUniversal in writing (at the address 
                      set out in the Notice provision of the Terms of Use Agreement) 
                      and specifically request that TradeUniversal return and/or 
                      destroy all copies of all or part of your Collected Information 
                      in TradeUniversal's possession. This Privacy Policy was 
                      last amended on April 10, 2003. </p>
                    <p class="unnamed2"><strong>9. <font color="#990000">Correcting 
                      Your Information</font> </strong></p>
                    <p>You can access, view and edit your Registration Information, 
                      Publishing Information and Payment Information (if any) 
                      through your Account with TradeUniversal. To update or correct 
                      such information, please click here. </p>
                    <p class="unnamed2"><strong>10. <font color="#990000">Your 
                      feedback</font> </strong></p>
                    <p>TradeUniversal welcomes your continuous input regarding 
                      our Privacy Policy or our services provided to you. You 
                      may send us your comments and responses to sictec.bj@sictec.com<br>
                      <br>
                    </p></td>
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
  <area shape="rect" coords="19,12,69,26" href="explain.asp">
  <area shape="rect" coords="110,10,187,27" href="explain.asp">
  <area shape="rect" coords="220,13,365,27" href="explain.asp">
  <area shape="rect" coords="393,11,473,27" href="explain.asp">
  <area shape="rect" coords="493,11,585,27" href="contactus.asp">
</map>
</body>
</html>
