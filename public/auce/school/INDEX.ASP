<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<html>
<head>
<title>商务校友录</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="同学、同学录、校友、老师、学校、班级、交流">
<STYLE type=text/css>
</STYLE>
<LINK href="scss.css" rel=stylesheet>
</head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="5" link="#003366" vlink="#006666"><CENTER>
<!--#include file=top2.htm-->
<table width="750" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>

  </tr>
</table>
<table width="759" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="#FFFFFF" width="181"> 
      <table width="180" border="0" cellspacing="2" cellpadding="2">
        <tr> 
        </tr>
        <tr> 
          
          <%
          set rs = createobject("ADODB.recordset")
          set rss = createobject("ADODB.recordset")
          sql="select Top 5 * from classinfo order by regdate "
          rs.open SQL,schooldb
         do while not rs.eof 
           classname=rs("classname")
          %>      
           
           
          <%rs.movenext
            i=i+1
        n=cstr(i)  
  if n=5 then exit do
           loop
            rs.close
         i=0
          %>    
            </td>
        </tr>
      </table>
      <br>
            
    </td>
    <td width="10" height="363"></td>
    <td bgcolor="#FFFFFF" align="center" valign="top" width="370"> 
      <table width="100%" border="0" cellspacing="2" cellpadding="2" bgcolor="#FFFFFF" height="339">
        <tr> 
          <td align="left" bgcolor="#FFFFFF" height="16" valign="bottom" width="50%">　</td>
          <td align="center" bgcolor="#FFFFFF" height="16" valign="bottom" width="50%" class="topic1">        </tr>
        <tr> 
          <td align="center" bgcolor="#FFFFFF" height="308" valign="top" colspan="2"><img src="school_images/china_map.gif" width="350" height="289" usemap="#Map" border="0" href="local.asp"><br>
            <br>
            请您点击想要浏览学校所在的省市名称！<font color="#006666"><a href="schoollist.asp"><font color="#0000E8">或者点击这里查找</font></a></font></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="2"></td>
        </tr>
      </table>
      
    </td>
    <td width="12"></td>
    <td bgcolor="#FFFFFF" valign="top" width="188" align="center"> 
      <table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td height="40" valign="bottom">　</td>
        </tr>
        <tr> 
          <td class="topic" valign="top" align="center">
            <table width="180" border="0" cellspacing="1" cellpadding="1">
              <form name="form2" method="post" action="classlogin.asp"> <tr>
                  <td valign="bottom" align="center">用户名： 
                    <input type="text" name="userid" size="15">
                  
                </td>
              </tr>
              <tr>
                  <td valign="bottom" align="center">密&nbsp;&nbsp;码： 
                    <input type="password" name="password" size="15">
                  </td>
              </tr>
              <tr>
                  <td align="center" height="40"> <a href="register.htm">
                  <font color="#0000FF">新用户注册</font></a> 
                    <input type="submit" name="Submit2" value="立刻登录">
                  </td>
              </tr></form>
            </table>
          </td>
        </tr>
      </table>
      </td>
  </tr>
  <tr>
    <td align="center" valign="top" bgcolor="#FFFFFF" width="761" colspan="5"> 
<form name="form1" method="post" action="class/popwindow_searchout.asp" target="_blank">      同学姓名： 
              <input type="text" name="realname" size="12">&nbsp;&nbsp; 所在学校： 
              <input type="text" name="schoolname" size="12">&nbsp;&nbsp; 所在班级： 
              <input type="text" name="classname" size="12">&nbsp;&nbsp; 入学年份：<select name=enterdate>
                <option value="0" selected>请选择…</option>
                <option value=1970>1970年</option>
                <option value=1971>1971年</option>
                <option value=1972>1972年</option>
                <option value=1973>1973年</option>
                <option value=1974>1974年</option>
                <option value=1975>1975年</option>
                <option value=1976>1976年</option>
                <option value=1977>1977年</option>
                <option value=1978>1978年</option>
                <option value=1979>1979年</option>
                <option value=1980>1980年</option>
                <option value=1981>1981年</option>
                <option value=1982>1982年</option>
                <option value=1983>1983年</option>
                <option value=1984>1984年</option>
                <option value=1985>1985年</option>
                <option value=1986>1986年</option>
                <option value=1987>1987年</option>
                <option value=1988>1988年</option>
                <option value=1989>1989年</option>
                <option value=1990>1990年</option>
                <option value=1991>1991年</option>
                <option value=1992>1992年</option>
                <option value=1993>1993年</option>
                <option value=1994>1994年</option>
                <option value=1995>1995年</option>
                <option value=1996>1996年</option>
                <option value=1997>1997年</option>
                <option value=1998>1998年</option>
                <option value=1999>1999年</option>
                <option value=2000>2000年</option>
                <option value=2001>2001年</option>
              </select>&nbsp; 
              <input type="submit" name="Submit" value="开始查找"></form></td>
    </tr>
</table>
<br>
<br>
      
<map name="Map">
  <area shape="circle" coords="81,86,15" href="local.asp?provinceid=32">
  <area shape="circle" coords="73,168,15" href="local.asp?provinceid=27">
  <area shape="circle" coords="130,142,15" href="local.asp?provinceid=30">
  <area shape="circle" coords="167,190,14" href="local.asp?provinceid=23">
  <area shape="circle" coords="160,231,12" href="local.asp?provinceid=26">
  <area shape="circle" coords="222,272,14" href="local.asp?provinceid=22">
  <area shape="circle" coords="304,229,13" href="local.asp?provinceid=24">
  <area shape="circle" coords="309,172,11" href="local.asp?provinceid=02">
  <area shape="circle" coords="186,151,11" href="local.asp?provinceid=29">
  <area shape="circle" coords="197,131,10" href="local.asp?provinceid=31">
  <area shape="circle" coords="243,81,14" href="local.asp?provinceid=16">
  <area shape="circle" coords="309,42,15" href="local.asp?provinceid=17">
  <area shape="circle" coords="311,75,11" href="local.asp?provinceid=08">
  <area shape="circle" coords="293,93,11" href="local.asp?provinceid=07">
  <area shape="rect" coords="257,86,272,103" href="local.asp?provinceid=01">
  <area shape="rect" coords="254,106,273,118" href="local.asp?provinceid=03">
  <area shape="rect" coords="242,115,254,128" href="local.asp?provinceid=05">
  <area shape="rect" coords="257,130,275,146" href="local.asp?provinceid=14">
  <area shape="rect" coords="225,125,241,144" href="local.asp?provinceid=06">
  <area shape="rect" coords="203,152,222,170" href="local.asp?provinceid=28">
  <area shape="rect" coords="233,150,255,169" href="local.asp?provinceid=15">
  <area shape="rect" coords="278,159,293,173" href="local.asp?provinceid=09">
  <area shape="rect" coords="260,167,275,184" href="local.asp?provinceid=11">
  <area shape="rect" coords="281,185,300,203" href="local.asp?provinceid=10">
  <area shape="rect" coords="190,173,207,192" href="local.asp?provinceid=04">
  <area shape="rect" coords="189,210,211,229" href="local.asp?provinceid=25">
  <area shape="rect" coords="203,234,226,252" href="local.asp?provinceid=21">
  <area shape="rect" coords="227,173,248,193" href="local.asp?provinceid=18">
  <area shape="rect" coords="222,200,244,219" href="local.asp?provinceid=19">
  <area shape="rect" coords="237,225,257,242" href="local.asp?provinceid=20">
  <area shape="rect" coords="246,246,263,265" href="local.asp?provinceid=34">
  <area shape="rect" coords="252,194,268,211" href="local.asp?provinceid=13">
  <area shape="rect" coords="269,211,288,228" href="local.asp?provinceid=12">
  <area shape="rect" coords="264,237,282,261" href="local.asp?provinceid=33">
  <area shape="circle" coords="210,111,10" href="local.asp?provinceid=16">
  <area shape="circle" coords="271,40,9" href="local.asp?provinceid=16">
</map>
<CENTER>
<p></p>
<!--#include file=end.htm--></body>
</html>