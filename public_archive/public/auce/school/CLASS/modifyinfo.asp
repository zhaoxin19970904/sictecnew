<%@ Language=VBScript %>
<!--#include file=globals.asp -->

<%
'解析参数
userid=trim(session("userid"))
realname=trim(request.form("realname"))
dearname=trim(request.form("dearname"))
passwd=trim(request.form("passwd"))
confirmpasswd=trim(request.form("confirmpasswd"))
byear=trim(request.form("byear"))
bmonth=trim(request.form("bmonth"))
bday=trim(request.form("bday"))
passask=trim(request.form("passask"))
anwserpass=trim(request.form("anwserpass"))
communicationaddr=trim(request.form("communicationaddr"))
iscashow=trim(request.form("iscashow"))
telephone=trim(request.form("telephone"))
istelephoneshow=trim(request.form("istelephoneshow"))
mobile=trim(request.form("mobile"))
ismobileshow=trim(request.form("ismobileshow"))
bp=trim(request.form("bp"))
isbpshow=trim(request.form("isbpshow"))
qq=trim(request.form("qq"))
isqqshow=trim(request.form("isqqshow"))
workshop=trim(request.form("workshop"))
iswsshow=trim(request.form("iswsshow"))
email=trim(request.form("email"))
isemailshow=trim(request.form("isemailshow"))
addinfo=trim(request.form("addinfo"))
isafshow=trim(request.form("isafshow"))
homeaddr=trim(request.form("homeaddr"))
ishashow=trim(request.form("ishashow"))


if userid="" then
   response.redirect "../error.asp?info=对不起，您掉线了，请重新登陆！"
end if

if realname="" or passwd="" or confirmpasswd="" or byear="" or bmonth="" or bday="" or passask="" or anwserpass="" or communicationaddr="" or email="" or telephone="" then
  response.redirect "../error.asp?info=对不起，您的注册信息未填完整，请重新填写！"
end if

if len(realname)>5 or len(dearname)>5 then
  response.redirect "../error.asp?info=对不起,真实姓名或呢称不能超过５个字，请重新填写！"
end if

if len(passwd)<5 then
  response.redirect "../error.asp?info=对不起， 密码不能少于５位，请重新填写！"
end if

if passwd<>confirmpasswd then
  response.redirect "../error.asp?info=对不起,两次输入的密码不一致，请重新填写！"
end if

 '验证用户添入生日
if not IsNumeric(byear) then
    response.redirect "../error.asp?info=对不起,请您正确填写出生日期！"
else
   byear = CLng(Request("byear"))
end if
    
if not IsNumeric(bmonth) then
   response.redirect "../error.asp?info=对不起,请您正确填写出生日期！"
else
   bmonth = CLng(Request("bmonth"))
end if
    
if not IsNumeric(bday) then
   response.redirect "../error.asp?info=对不起,请您正确填写出生日期！"
else
   bday = CLng(Request("bday"))
end if

if communicationaddr<>"" and iscashow<>"" then
   iscashow=0
else
   iscashow=1
end if

if telephone<>"" and istelephoneshow<>"" then
   istelephoneshow=0
else
   istelephoneshow=1
end if
      
if mobile<>"" and ismobileshow<>"" then
   ismobileshow=0
else
   ismobileshow=1
end if    

if bp<>"" and isbpshow<>"" then
   isbpshow=0
else
   isbpshow=1
end if 

if qq<>"" and isqqshow<>"" then
   isqqshow=0
else
   isqqshow=1
end if

if workshop<>"" and iswsshow<>"" then
   iswsshow=0
else
   iswsshow=1
end if  

if email<>"" and isemailshow<>"" then
   isemailshow=0
else
   isemailshow=1
end if  
  
if addinfo<>"" and isafshow<>"" then
   isafshow=0
else
   isafshow=1
end if

if homeaddr<>"" and ishashow<>"" then
   ishashow=0
else
   ishashow=1
end if


'入库
set rs = createobject("ADODB.recordset")
SQL = "select * from UserInfo where userid='"&userid&"'"
rs.Open SQL,schooldb,1,3
if not rs.eof then    
  rs("userid")=userid
  rs("realname")=realname
  rs("dearname")=dearname
  rs("userpassword")=passwd
  if not (byear=0 or bmonth=0 or bday=0) then  
     rs("UserBirth") = DateSerial(byear,bmonth,bday)
  end if
  rs("passask")=passask
  rs("anwserpass")=anwserpass
  rs.Update
end if  
rs.Close


SQL = "select * from UsercommunicationInfo where userid='"&userid&"'"
rs.Open SQL,schooldb,1,3
if not rs.eof then    
  rs("userid")=userid
  rs("communicationaddr")=communicationaddr
  rs("iscashow")=iscashow
  rs("homeaddr")=homeaddr
  rs("ishashow")=ishashow
  rs("email")=email
  rs("isemailshow")=isemailshow
  rs("qq")=qq
  rs("isqqshow")=isqqshow
  rs("telephone")=telephone
  rs("istelephoneshow")=istelephoneshow
  rs("mobile")=mobile
  rs("ismobileshow")=ismobileshow
  rs("bp")=bp
  rs("isbpshow")=isbpshow
  rs("workshop")=workshop
  rs("iswsshow")=iswsshow
if trim(addinfo)="" then
addinfo="无"
end if
  rs("addinfo")=addinfo
  rs("isafshow")=isafshow
  rs.Update
end if  
rs.Close
set rs=nothing
response.redirect "class_index.asp"

%>