<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
userid=trim(session("userid"))
curclassid=trim(session("classid"))
mailtitle=trim(request.form("mailtitle"))
mailbody=trim(request.form("mailbody"))
if userid="" then
    response.redirect "../error.asp?info=对不起，您已经掉线了，请重新申请！"
end if
set rs = createobject("ADODB.recordset")
set rss = createobject("ADODB.recordset")
sql="select * from classinfo where classid='"&curclassid&"'"
rs.open SQL,schooldb
if not rs.eof then
   monitor=rs("monitor")
   monitor1=rs("monitor1")
end if
rs.close
if userid<>monitor and userid<>monitor1 then
   response.redirect "../error.asp?info=对不起，您不是本班班长无权进入！"
end if  
sql="select * from usercommunicationinfo where userid='"&userid&"'"
rs.open SQL,schooldb
if not rs.eof then
   mailto=rs("email")
end if
rs.close
sql="select * from userjoinclassinfo where classid='"&curclassid&"'"
rs.open SQL,schooldb
while not rs.eof 
 
    sql1="select * from userinfo where userid='"&rs("userid")&"'"
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
    MailFrom = "school.269.net"
    MailBody = MailBody &  CRLF
    MailBody = MailBody & "——269同学录"& CRLF
    MailBody = MailBody & "http://school.269.net"& CRLF
    Set MailOb = Server.CreateObject("CDONTS.NewMail")

   
    MailBody = replace(MailBody,"<br>","")
    MailBody = replace(MailBody,"  ",chr(13)) 
    MailOb.Send mailto,curemail,Mailtitle,MailBody
    
   
    Set MailObject = nothing
   
  rs.movenext 
wend
rs.close  
response.redirect "../success.asp?info= 您发信成功了！"
%> 