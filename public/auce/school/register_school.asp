<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
userid=session("userid")
if userid="" then
  response.redirect "error.asp?info=对不起，您已经掉线了，请重新申请！"
end if
schoolname=trim(request.form("schoolname"))
schoolweb=trim(request.form("schoolweb"))
schooladdr=trim(request.form("schooladdr"))
schoolzip=trim(request.form("schoolzip"))
enterdate=trim(request.form("enterdate"))
classname=trim(request.form("classname"))
areaid=trim(request.form("areaid"))
schooltypeid=trim(request.form("schooltypeid"))
if schoolname="" then
   response.redirect "error.asp?info=对不起，学校名称不能为空，请重新输入！"
end if   
i = 1
len1 = 0
charinner = mid(schoolname,1,1)

while not charinner = ""
   
    if (charinner >="0" and charinner<="9") then
      
      response.redirect "error.asp?info=对不起，数字请用中文表示，请重新输入！"
   else
      len1 = len1 + 1

   end if  
   i = i + 1
   charinner = mid(schoolname,i,1)
wend
if len1 > 30 then
    response.redirect "error.asp?info=对不起，您所添入的学校名称字符不能多于3０个，请重新输入！"
end if
if right(classname,1)<>"班" then
   response.redirect "error.asp?info=对不起，班级名称最后必须以班结尾，请重新输入！"
end if
if enterdate="0" then
   response.redirect "error.asp?info=对不起，必须选择入校年份！"
end if  


set rs = createobject("ADODB.recordset")
sql="select * from schoolinfo where schoolname='"&schoolname&"'"
rs.open SQL,schooldb
if not rs.eof then
   response.redirect "error.asp?info=对不起，学校已经存在了！"
end if
rs.close

sql="select * from areainfo where areaid='"&areaid&"'"
rs.open SQL,schooldb
if rs.eof then
   response.redirect "error.asp?info=对不起，地区不存在！"
end if
rs.close
SQL = "select * from schoolinfo where areaid='"&areaid&"'order by seed desc"
rs.open SQL,schooldb
if not rs.eof then
  seed=rs("seed")+1
else
  seed=1  
end if
rs.close
if seed<10 then
   schoolid=cstr(areaid&"000"&seed)
end if
if seed>=10 and seed<100 then
   schoolid=cstr(areaid&"00"&seed)
end if   
if seed>=100 and seed<1000 then
   schoolid=cstr(areaid&"0"&seed)
end if
if seed>=1000 then
   schoolid=cstr(areaid&"000"&seed)
end if
sql="select * from schoolinfo"
  rs.open SQL,schooldb,1,3
  rs.AddNew    
  rs("regdate")=now()
  rs("schoolid")=schoolid
  rs("schoolname")=schoolname
  rs("schooltypeid")=schooltypeid
  rs("schoolweb")=schoolweb
  rs("schooladdr")=schooladdr
  rs("schoolzip")=schoolzip
  rs("buildman")=userid
  rs("seed")=seed
  rs("areaid")=areaid
  rs.Update
  rs.Close
  

curclassid=cstr(schoolid&"0001")
sql="select * from classinfo"
  rs.open SQL,schooldb,1,3
  rs.AddNew    
  rs("classid")=curclassid
  rs("classname")=classname
  rs("schoolid")=schoolid
  rs("monitor")=userid
  rs("enterdate")=enterdate
  rs("seed")=1
  rs("regdate")=now()
  rs.Update
  rs.Close 
session("curclassid")=curclassid
response.redirect "find_class5.asp" 
%>    
