<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- saved from url=(0028)http://en.ningbo-export.com/ -->
<html>
<head>
<title>Member Log On </title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
.input {
	border:1px solid #c3c3c3; padding:1px; FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: "Verdana", "Arial", "宋体"; HEIGHT: 20px; BACKGROUND-COLOR: #ffffff
}
</style>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr>
    <td>
      <table width="370" border="0" cellspacing="0" cellpadding="0" align="center" height="140">
        <tr> 
          <td width="17"><img src="images/jx46.jpg" width="32" height="229"></td>
          <td width="320" background="images/jx47.jpg"> 
		  <form name="form1" method="post" action="script/login.asp" onSubmit="return checkempty(this)">
              <table width="200" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="46"><img src="images/login_01.gif" width="45" height="33"></td>
                  <td width="65">&nbsp;</td>
                  <td width="89" valign="top">
<div align="right"><a href="index.asp"><font color="#939393"><strong>HOME</strong></font></a></div></td>
                </tr>
                <tr> 
                  <td><img src="images/login_03.gif" width="45" height="27"></td>
                  <td colspan="2" rowspan="2"><table width="154" height="54" border="0" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td width="105" height="27">
						<input name="username" type="text" size="13" class="input" id="username"></td>
                          <td width="49" rowspan="2"> 
                          <input type="image" src="images/login_05.gif" width="49" height="56" 
						  border="0">
                          </td>
                      </tr>
                      <tr> 
                        <td height="16">
						<input name="userpass" type="password" size="13" class="input" id="userpass"></td>
                      </tr>
                    </table> </td>
                </tr>
                <tr> 
                  <td><img src="images/login_06.gif" width="45" height="29"></td>
                </tr>
                <tr> 
                  <td><a href="index.asp"><font color="#939393"></font></a></td>
                  <td><a href="find_p_w.asp" target="_blank"><img src="images/login_08.gif" width="61" height="31" border="0"></a></td>
                  <td><a href="register.asp"> <img src="images/login_09.gif" width="89" height="31" border="0"></a></td>
                </tr>
              </table>
            </form></td>
          <td width="33"><img src="images/jx46_1.jpg" width="32" height="229"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
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
