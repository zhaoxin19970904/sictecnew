<%@ Language=VBScript %>
<!--#include file=globals.asp -->

<%
userid=cstr(trim(request.form("userid")))

if userid = "" then
    response.redirect "error.asp?info=�Բ�������������û�������Ϊ�գ����������룡"
end if
set rs = createobject("ADODB.recordset")
SQL = "select * from userinfo where userid='"&userid&"'"
rs.open SQL,schooldb
if not rs.eof then
  response.redirect "error.asp?info=�Բ�������������û����Ѿ����ڣ����������룡"
end if
rs.close
i = 1
len1 = 0
charinner = mid(userid,1,1)

while not charinner = ""
   
    if ((charinner >= "a" and charinner <="z") or (charinner >="0" and charinner<="9") or (charinner >="A" and charinner<="Z") or charinner="_") then
      len1 = len1 + 1
      
   else
      response.redirect "error.asp?info=�Բ�������������û������в��Ϸ����ַ������������룡"

   end if  
   i = i + 1
   charinner = mid(UserID,i,1)
wend

   
if len1 < 3 then
    response.redirect "error.asp?info=�Բ�������������û����ַ���������3�������������룡"
   
elseif len1 > 15 then
    response.redirect "error.asp?info=�Բ�������������û����ַ����ܴ���15�������������룡"
  
else
session("userid")=userid
response.redirect "register_info.asp"

end if   
%>    