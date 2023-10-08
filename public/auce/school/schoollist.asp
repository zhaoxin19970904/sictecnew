<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
set rs = createobject("ADODB.recordset")
set rss = createobject("ADODB.recordset")
%>
<html>
<head>
<title>商务校友录-浏览学校</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="同学、同学录、校友、老师、学校、班级、交流">
<STYLE type=text/css>
</STYLE>
<LINK href="scss.css" rel=stylesheet>
</head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="5"><CENTER>
<!--#include file=top2.htm-->

<table width="750" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td align="center" valign="top" width="550"><br>
      <table width="520" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <%
            sql="select * from province order by provinceid"
            rs.open SQL,schooldb
            i=0
            while not rs.eof
               if i mod 4 =0 then response.write "</tr><tr>"
               sql1="select count(*) as schoolnum from schoolinfo where schoolid like '"&rs("provinceid")&"%'"
               rss.open SQL1,schooldb,1,3
               schoolnum=rss("schoolnum")
               rss.close
           %>    
          <td height="50" align="left" valign="top"> 
            <table width="130" border="0" cellspacing="1" cellpadding="1">
              <tr> 
                <td class="cn1" height="24" valign="top"><a href="local.asp?provinceid=<%=rs("provinceid")%>"><%=rs("provincename")%></a></td>
              </tr>
              <tr> 
                <td>共注册<%=schoolnum%>所学校</td>
              </tr>
            </table>
          </td>
          <%
          i=i+1
          rs.movenext
          wend
          rs.close
          %>          
    </table>
    </td>
    <td valign="top" width="200" align="center"> 
      <table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td height="40" valign="bottom" align="center">校友录登陆</td>
        </tr>
        <tr> 
          <td class="topic" valign="top" align="center"> 
            <table width="180" border="0" cellspacing="1" cellpadding="1">
              <form name="form2" method="post" action="classlogin.asp">
                <tr> 
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
                </tr>
              </form>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  </table>

</body>
</html>
       <CENTER>
<p></p>
<!--#include file=end.htm-->