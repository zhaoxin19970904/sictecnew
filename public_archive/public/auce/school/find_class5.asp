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
   userstatus="��Ա"
end if   

if userid="" then
  response.redirect "error.asp?info=�Բ������Ѿ������ˣ����������룡"
end if
set rs = createobject("ADODB.recordset")
set rss = createobject("ADODB.recordset")
sql="select * from userjoinclassinfo where userid='"&userid&"' and classid='"&curclassid&"'"
rs.open SQL,schooldb
if not rs.eof then
  response.redirect "error.asp?info=�Բ������Ѿ��������ĳ�Ա�ˣ�"
end if
rs.close
sql="select * from classinfo where classid='"&curclassid&"'"
rs.open SQL,schooldb
if not rs.eof then
  classname=rs("classname")
else
  response.redirect "error.asp?info=�Բ��𣬸ð༶�����ڣ�"
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
<title>����У��¼-����༶�ɹ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="ͬѧ��ͬѧ¼��У�ѡ���ʦ��ѧУ���༶������">
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
          <td height="40" valign="bottom" align="center">��</td>
        </tr>
      </table>
      <br>
      <table width="550" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="center" class="cn" height="30"><font color="#CC0000"> ��ϲ�������Ѿ�ע���Ϊ�༶:<%=schoolname%> 
            <%=classname%>��һ����Ա�ˣ�</font></td>
        </tr>
      </table>
      <br>
      <table width="500" border="0" cellspacing="1" cellpadding="1">
        <form name="form1" method="post" action="ClassLogin.asp">
          <tr> 
            <td class="cn" height="28">����ʹ���û����������¼ͬѧ¼���İ༶��������</td>
          </tr>
          <tr> 
            <td align="center" height="30" valign="bottom">�û����� 
              <input type="text" name="userid" size="20">
            </td>
          </tr>
          <tr> 
            <td align="center" height="30" valign="bottom">��&nbsp;&nbsp;�룺 
              <input type="password" name="password" size="0" maxlength="12">
            </td>
          </tr>
          <tr> 
            <td height="40" align="center"> 
              <input type="submit" name="Submit" value="��¼">
            </td>
          </tr>
        </form>
      </table>
      <br>
      <br>
      <table width="510" border="0" cellspacing="1" cellpadding="1">
        <tr align="center"> 
          <td><a href="index.asp"><font color="#000000">����ͬѧ¼��ҳ</font></a></td>
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