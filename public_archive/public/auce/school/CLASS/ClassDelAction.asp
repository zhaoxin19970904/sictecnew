<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
UserID=trim(session("UserID"))
ClassID=trim(session("ClassID"))
if UserID="" then
    response.redirect "../error.asp?info=�Բ������Ѿ������ˣ����������룡"
end if
set conn = server.createobject("ADODB.connection")
conn.open DBParams 
sql="select * from ClassInfo where ClassID='"&ClassID&"'"
set rs=conn.Execute(sql)
if not rs.eof then
   monitor=rs("monitor")
   monitor1=rs("monitor1")
end if
rs.close
Manager=false
if UserID=monitor or UserID=monitor1 or UserID="jackey" then Manager=true


select case request("Act")
   case 1
      if Manager then
         sql="update message set deleted=1 where ClassID='"&ClassID&"' and MsgID='"&request("MsgID")&"'"
         conn.Execute(sql)
         response.redirect "ClassBBSList.asp"
      else
         response.redirect "../error.asp?info=�Բ��������ǰ೤��û��Ȩ����"
      end if
   case 2
      if Manager or UserID=request("UserID") then
         sql="update message set deleted=1 where ClassID='"&ClassID&"' and MsgID='"&request("MsgID")&"'"
         conn.Execute(sql)
         response.redirect "popwindow_Perliuyan.asp"
      else
         response.redirect "../error.asp?info=�Բ��������ǰ೤��û��Ȩ����"
      end if   
   case 3
      if Manager then
         sql="update UpLoad set deleted=1 where ClassID='"&ClassID&"' and PicID='"&request("PicID")&"'"
         conn.Execute(sql)
         response.redirect "ClassPhotoList.asp"
      else
         response.redirect "../error.asp?info=�Բ��������ǰ೤��û��Ȩ����"
      end if
   case 4
      if Manager then
         sql="update UpLoadReply set deleted=1 where ClassID='"&ClassID&"' and PicID='"&request("PicID")&"' and MsgID='"&request("MsgID")&"'"
         conn.Execute(sql)
         response.redirect "ClassPhotoView.asp?PicID="&request("PicID")
      else
         response.redirect "../error.asp?info=�Բ��������ǰ೤��û��Ȩ����"
      end if   
end select
response.redirect Request.ServerVariables("http_referer")
%>