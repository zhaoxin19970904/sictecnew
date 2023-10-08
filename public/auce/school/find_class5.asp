<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%userid=session("userid")
realname=session("realname")
curclassid=trim(session("curclassid"))
if curclassid="" then
   curclassid=trim(request.form("classid"))
end if   

userstatus=trim(request.form("userstatus"))
if userstatus="" then
   userstatus="成员"
end if   

if userid="" then
  response.redirect "error.asp?info=对不起，您已经掉线了，请重新申请！"
end if
set rs = createobject("ADODB.recordset")
set rss = createobject("ADODB.recordset")
sql="select * from userjoinclassinfo where userid='"&userid&"' and classid='"&curclassid&"'"
rs.open SQL,schooldb
if not rs.eof then
  response.redirect "error.asp?info=对不起，您已经是这个班的成员了！"
end if
rs.close
sql="select * from classinfo where classid='"&curclassid&"'"
rs.open SQL,schooldb
if not rs.eof then
  classname=rs("classname")
else
  response.redirect "error.asp?info=对不起，该班级不存在！"
end if
rs.close

sql="select * from userjoinclassinfo"
rs.open SQL,schooldb,1,3
rs.addnew
rs("logintimes")=0
rs("joindate")=now()
rs("lastlogintime")=now()
rs("userid")=userid
rs("classid")=curclassid
rs("userstatus")=userstatus
rs.update
rs.close





schoolid=left(curclassid,8)
sql="select * from schoolinfo where schoolid='"&schoolid&"'"
rs.open SQL,schooldb
if not rs.eof then
   schoolname=rs("schoolname")
end if  
rs.close 

sql="select * from usercommunicationinfo where userid='"&userid&"'"
rs.open SQL,schooldb
if not rs.eof then
   mailto=rs("email")
end if
rs.close

sql="select * from userjoinclassinfo where classid='"&curclassid&"'"
rs.open SQL,schooldb
while not rs.eof 
  if userid<>rs("userid") then
    sql1="select * from userinfo where userid='"&userid&"'"
    rss.open SQL1,schooldb
    if not rss.eof then
      currealname=rss("realname")
    end if
    rss.close    
    sql1="select * from usercommunicationinfo where userid='"&rs("userid")&"'"
    rss.open SQL1,schooldb
    if not rss.eof then
      curemail=rss("email")
    end if
    rss.close
    Set MailObject = nothing
  end if   
  rs.movenext 
wend
rs.close  
Session.Abandon    
%>
<html>
<head>
<title>商务校友录-加入班级成功！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="同学、同学录、校友、老师、学校、班级、交流">
<STYLE type=text/css>
</STYLE>
<LINK href="scss.css" rel=stylesheet>
</head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="5"><CENTER>
<!--#include file=top2.htm-->

<div align="center">
  <center>
<table width="760" border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
  <tr> 
    <td align="center" valign="top" bgcolor="#F6F6F6"> <br>
      <br>
      <table width="98%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td height="40" valign="bottom" align="center">　</td>
        </tr>
      </table>
      <br>
      <table width="550" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="center" class="cn" height="30"><font color="#CC0000"> 恭喜您！您已经注册成为班级:<%=schoolname%> 
            <%=classname%>的一名成员了！</font></td>
        </tr>
      </table>
      <br>
      <table width="500" border="0" cellspacing="1" cellpadding="1">
        <form name="form1" method="post" action="ClassLogin.asp">
          <tr> 
            <td class="cn" height="28">请您使用用户名和密码登录同学录您的班级－－－－</td>
          </tr>
          <tr> 
            <td align="center" height="30" valign="bottom">用户名： 
              <input type="text" name="userid" size="20">
            </td>
          </tr>
          <tr> 
            <td align="center" height="30" valign="bottom">密&nbsp;&nbsp;码： 
              <input type="password" name="password" size="0" maxlength="12">
            </td>
          </tr>
          <tr> 
            <td height="40" align="center"> 
              <input type="submit" name="Submit" value="登录">
            </td>
          </tr>
        </form>
      </table>
      <br>
      <br>
      <table width="510" border="0" cellspacing="1" cellpadding="1">
        <tr align="center"> 
          <td><a href="index.asp"><font color="#000000">返回同学录首页</font></a></td>
        </tr>
      </table>
      <br>
      <br>
      <br>
    </td>
  </tr>
</table>
  </center>
</div>
<br>
<br>
      <CENTER>
<p></p>
<!--#include file=end.htm--> </body>
</html>