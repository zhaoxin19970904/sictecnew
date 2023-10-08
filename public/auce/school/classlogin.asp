<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
curUserID = Trim(Request.form("UserID"))
addcurclassid=trim(request.form("curclassid"))
Password = Trim(Request.form("Password"))
userip = Request.ServerVariables("REMOTE_ADDR")
userid=trim(session("userid"))

if curuserid="" then
  response.redirect "error.asp?info=对不起，用户名不能为空，请重新输入！"
end if


set rs = createobject("ADODB.recordset")
set rss = createobject("ADODB.recordset")
sql="select * from userinfo where userid='"&curuserid&"'"
rs.open SQL,schooldb
if not rs.eof then
   if rs("userpassword")<>password then
      response.redirect "error.asp?info=对不起，密码不正确，请重新输入！"
   end if
else
   response.redirect "error.asp?info=对不起，用户名不存在，请重新输入！"
end if
rs.close

flag=0
sql="select * from online where userid='"&curuserid&"'"
rs.open SQL,schooldb
if not rs.eof  then
   flag=1
end if
rs.close
if flag=1 then
  sql1="delete from online where userid='"&curuserid&"'"
  rss.open SQL1,schooldb,1,3
  response.redirect "error.asp?info=对不起，你刚才已经登入了一次,请过10秒后再登录!"
end if

'if userid<>"" then
'  Session.Abandon
'end if


userid=curuserid
session("userid")=userid

sql="select * from userinfo where userid='"&userid&"'"
rs.open SQL,schooldb
if not rs.eof then
  session("realname")=rs("realname")
end if
rs.close

SQL = "select * from userjoinclassinfo where userid='"&userid&"' order by joindate"
rs.open SQL,schooldb
if rs.eof then
  sql1="select  * from online"
  rss.open SQL1,schooldb,1,3
  rss.AddNew
  rss("userid")=userid
  rss("userip")=userip
  rss("usercometime")=now()
  rss("classid")="nothing"
  rss.Update
  rss.Close
  response.redirect "find_class1.asp"
else
  if addcurclassid="" then
     curclassid=rs("classid")
  else
     sql1="select * from userjoinclassinfo where userid='"&userid&"' and classid='"&addcurclassid&"'"
     rss.open SQL1,schooldb
     if rss.eof then
        response.redirect "error.asp?info=对不起，您不是本班成员，请先注册在进入！"
     end if
     rss.close
     curclassid=addcurclassid
  end if
  sql1="select * from online"
  rss.open SQL1,schooldb,1,3
  rss.AddNew
  rss("userid")=userid
  rss("userip")=userip
  rss("usercometime")=now()
  rss("classid")=curclassid
  rss.Update
  rss.Close
  session("classid")=curclassid
  sql1="select * from userjoinclassinfo where userid='"&userid&"' and classid='"&curclassid&"'"
  rss.open SQL1,schooldb,1,3
  if not rss.eof then
     rss("logintimes")=rss("logintimes")+1
     rss("lastlogintime")=now
     rss.update
  end if
  rss.close
  response.redirect "class/class_index.asp"
end if

rs.close
%>
