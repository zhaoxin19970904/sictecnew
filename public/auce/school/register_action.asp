<%@ Language=VBScript %>
<!--#include file=globals.asp -->

<%
userid=cstr(trim(request.form("userid")))

if userid = "" then
    response.redirect "error.asp?info=对不起，您所添入的用户名不能为空，请重新输入！"
end if
set rs = createobject("ADODB.recordset")
SQL = "select * from userinfo where userid='"&userid&"'"
rs.open SQL,schooldb
if not rs.eof then
  response.redirect "error.asp?info=对不起，您所添入的用户名已经存在，请重新输入！"
end if
rs.close
i = 1
len1 = 0
charinner = mid(userid,1,1)

while not charinner = ""
   
    if ((charinner >= "a" and charinner <="z") or (charinner >="0" and charinner<="9") or (charinner >="A" and charinner<="Z") or charinner="_") then
      len1 = len1 + 1
      
   else
      response.redirect "error.asp?info=对不起，您所添入的用户名中有不合法的字符，请重新输入！"

   end if  
   i = i + 1
   charinner = mid(UserID,i,1)
wend

   
if len1 < 3 then
    response.redirect "error.asp?info=对不起，您所添入的用户名字符不能少于3个，请重新输入！"
   
elseif len1 > 15 then
    response.redirect "error.asp?info=对不起，您所添入的用户名字符不能大于15个，请重新输入！"
  
else
session("userid")=userid
response.redirect "register_info.asp"

end if   
%>    