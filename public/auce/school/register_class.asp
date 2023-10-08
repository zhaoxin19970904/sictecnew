<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
userid=session("userid")
if userid="" then
  response.redirect "error.asp?info=对不起，您已经掉线了，请重新申请！"
end if
classname=trim(request.form("classname"))
schoolid=trim(request.form("schoolid"))
enterdate=trim(request.form("enterdate"))
if right(classname,1)<>"班" then
   response.redirect "error.asp?info=对不起，班级名称最后必须以班结尾，请重新输入！"
end if
set rs = createobject("ADODB.recordset")
sql="select * from classinfo where classname='"&classname&"'"
rs.open SQL,schooldb
if not rs.eof then
   response.redirect "error.asp?info=对不起，该班级已经存在了！"
end if
rs.close
sql="select * from schoolinfo where schoolid='"&schoolid&"'"
rs.open SQL,schooldb
if rs.eof then
   response.redirect "error.asp?info=对不起，学校不存在！"
end if
rs.close
SQL = "select * from classinfo where schoolid='"&schoolid&"'order by seed desc"
rs.open SQL,schooldb
if not rs.eof then
  seed=rs("seed")+1
else
  seed=1  
end if
rs.close
if seed<10 then
   curclassid=cstr(schoolid&"000"&seed)
end if
if seed>=10 and seed<100 then
   curclassid=cstr(schoolid&"00"&seed)
end if   
if seed>=100 and seed<1000 then
   curclassid=cstr(schoolid&"0"&seed)
end if
if seed>=1000 then
   curclassid=cstr(schoolid&"000"&seed)
end if
sql="select * from classinfo"
  rs.open SQL,schooldb,1,3
  rs.AddNew    
  rs("classid")=curclassid
  rs("classname")=classname
  rs("schoolid")=schoolid
  rs("monitor")=userid
  rs("enterdate")=enterdate
  rs("seed")=seed
  rs("regdate")=now()
  rs.Update
  rs.Close 
session("curclassid")=curclassid
response.redirect "find_class5.asp"
%>
