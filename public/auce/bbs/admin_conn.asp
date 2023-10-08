<!--#include file="connect.asp" -->
<%
dim userislogin,master,rsa,userloginlock,loginuser,errmess
if request.cookies("renwen")("user")<>"" and request.cookies("renwen")("passedok")="ofdkjduy" then 
userislogin=true
loginuser=request.cookies("renwen")("user")
end if
if userislogin then
set rsa=conn.execute("select user,jibie,kill from userinfo where user='"&loginuser&"'")
if rsa.eof or rsa.bof then
userislogin=false
else
if rsa("jibie")<>"∆’Õ®”√ªß" then master=true end if
end if
userloginlock=rsa("kill")
end if 

dim pagename,casename
pagename=split(request.ServerVariables("SCRIPT_NAME"),"/")
casename=pagename(ubound(pagename))
if (session("user")="" or session("adpassdf")<>"passed") and casename<>"admin_login.asp" then
 response.redirect"admin_login.asp" 
end if
%>