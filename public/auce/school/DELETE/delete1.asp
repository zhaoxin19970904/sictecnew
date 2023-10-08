<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
curuserid=trim(request.form("userid"))
curpassword=trim(request.form("password"))
if curuserid<>"admin" then
   response.redirect "../error.asp?info=您无权进入！"
end if 
if curpassword<>"admin" then
   response.redirect "../error.asp?info=您无权进入！"
end if 
session("pass")="ok" 
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
<LINK href="../scss.css" rel=stylesheet>
</head><CENTER>
<!--#include file=../top2.htm-->
<body bgcolor="#FFFFFF" text="#000000" topmargin="5">
<table width="479" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td align="center" valign="top" bgcolor="#F6F6F6" width="550"><br>
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
                <td class="cn1" height="24" valign="top"><a href="delete2.asp?provinceid=<%=rs("provinceid")%>"><%=rs("provincename")%></a></td>
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
      <br>
      <br>
      <br>
    </td>
    <td width="1" background="school_images/vline.gif"></td>
  </tr>
</table>
<br>
<br> <CENTER>
<p></p>
<!--#include file=../end.htm-->      </body>
</html>