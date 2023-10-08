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
}
-->
</style>
</head>

<body background="image/bg01.jpg" text="#000000"   link="#000000" vlink="#000000" alink="#000000" topmargin="0" >
<div align="left">
  <table width="800" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999">
    <tr> 
      <td bgcolor="#FFFFFF"> <table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
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
                        <td><table border="0" cellspacing="0" cellpadding="0">
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
                          </table></td>
                      </tr>
                      <tr> 
                        <td valign="top"><table border="0" cellspacing="0" cellpadding="0">
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
                        <td height="34" valign="top"><img src="image/pho14.jpg" width="599" height="34" border="0" usemap="#MapMap"></td>
                      </tr>
                      <tr> 
                        <td height="138" valign="top"><img src="image/pho20.jpg" width="599" height="139"></td>
                      </tr>
                      <tr> 
                        <td height="25" valign="top" background="image/bg12.jpg"><table width="95%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td width="6%" height="23">&nbsp;</td>
                              <td width="94%" height="35" valign="bottom">&nbsp;</td>
                            </tr>
                          </table></td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
          </tr>
        </table>
        <div align="center"> 
          <table width="768" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr> 
              <td colspan="4"><div align="center"><strong> 
                  <%
 action=request.QueryString("action")
 user_name=request.form("user_name")
 email=trim(request.form("findpass"))
if action="find" then
   set rs=server.CreateObject("adodb.recordset")
   sql="select member_uid from member where member_uid='"&user_name&"'"
   rs.open sql,conn,1,1
   if rs.eof or rs.bof then 
%>
                  <SCRIPT language="JavaScript">
alert("Your userID is wrong!")
location.href="lost_p_w.asp"
</SCRIPT>
                  <%
		response.End()
     else
	   rs.close
       sql="select member_uid,psd,email from member where member_uid='"&user_name&"' and email='"&email&"'"
       rs.open sql,conn,1,1
	   
	   if not rs.eof or not rs.bof then
	   
	       okemail=rs("email")
           user_namefind=rs("member_uid")
		   user_passfind=rs("psd")
            
		 
		else
%>
                  <SCRIPT language="JavaScript">
alert("Your E-mail is wrong!")
location.href="lost_p_w.asp"
</SCRIPT>
                  <%
		     
			 response.End()
		end if
   end if
rs.close
%>
                  </strong> 
                  <table width="533" height="97" border="0" align="center" cellpadding="0" cellspacing="0" class="unnamed1">
                    <tr> 
                      <td width="265" height="28"><div align="center"><font color="#993333">Your       
                          User ID:</font></div></td>      
                      <td width="205"><%=user_namefind%></td>
                      <td width="63">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td height="37"> <div align="center"><font color="#993333">Your       
                          Password:</font></div></td>
                      <td><%=user_passfind%> </td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr> 
                      <td height="16" colspan="3"><div align="center">Please Don't       
                          Forget!</div></td>
                    </tr>
                    <tr> 
                      <td height="16" colspan="2"><div align="center">[ <a href="signin.asp">OK</a> 
                          ]</div></td>
                      <td>&nbsp;</td>
                    </tr>
                  </table>
                  <% 
 else
%>
                  <form name="form2" method="post" action="lost_p_w.asp?action=find" onsubmit="return check1(this)">
                    <table width="665" height="88" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td width="382" height="88"> <table width="571" height="99" border="0" align="center" cellpadding="0" cellspacing="0" class="unnamed1">
                            <tr> 
                              <td height="1" colspan="3"></td>
                            </tr>
                            <tr bgcolor="#3366CC"> 
                              <td height="1" colspan="3"></td>
                            </tr>
                            <tr> 
                              <td height="1" colspan="3"></td>
                            </tr>
                            <tr> 
                              <td width="140" height="27" bgcolor="#FFFFFF"> <div align="right">Your       
                                  User ID:</div></td>      
                              <td width=1></td>
                              <td bgcolor="#FFFFFF"> <div align="center"> 
                                  <input type="text" name="user_name" style="FONT-SIZE: 12px; WIDTH: 210px">
                                </div></td>
                            </tr>
                            <tr> 
                              <td height="1" colspan="3"></td>
                            </tr>
                            <tr> 
                              <td height="23" bgcolor="#FFFFFF"> <div align="right">Your 
                                  register E-mail:</div></td>
                              <td width=1></td>
                              <td bgcolor="#FFFFFF"><div align="center"> 
                                  <input type="text" name="findpass" style="FONT-SIZE: 12px; WIDTH: 210px">
                                </div></td>
                            </tr>
                            <tr> 
                              <td height="1" colspan="3"></td>
                            </tr>
                            <tr> 
                              <td height="35" bgcolor="#FFFFFF"> <div align="right"> 
                                </div></td>
                              <td width=1></td>
                              <td valign="bottom" bgcolor="#FFFFFF"> <div align="center"> 
                                  <input type="submit" name="Submit22" value="Submit">
                                  &nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;       
                                  <input type="reset" name="Submit2" value="Reset">
                                </div></td>
                            </tr>
                            <tr> 
                              <td height="1" colspan="3"></td>
                            </tr>
                          </table></td>
                      </tr>
                    </table>
                  </form>
                  <%end if%>
                </div></td>
            </tr>
            <tr> 
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
          </table>
          <map name="MapMap">
            <area shape="rect" coords="19,12,69,26" href="index.asp">
            <area shape="rect" coords="110,10,187,27" href="aboutus.asp">
            <area shape="rect" coords="219,13,364,27" href="index.asp">
            <area shape="rect" coords="393,11,473,27" href="demand.asp">
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


</body>
</html>
