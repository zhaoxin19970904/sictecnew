<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
firstid=trim(request.form("firstid"))
secondid=trim(request.form("secondid"))
set rs = createobject("ADODB.recordset")
set rss = createobject("ADODB.recordset")
set rrs = createobject("ADODB.recordset")
if firstid="" or secondid="" then
  response.redirect "../error.asp?info=对不起，请填写完整！"
end if
firstschoolid=left(firstid,8)
sql="delete classinfo where classid='"&secondid&"'"
rs.open SQL,schooldb,1,3

sql="select * from userjoinclassinfo where classid='"&secondid&"'"
rs.open SQL,schooldb,1,3
while not rs.eof 
    sql1="select * from userjoinclassinfo where classid='"&firstid&"' and userid='"&rs("userid")&"'"
    rss.open SQL1,schooldb
    if rss.eof then
       sql2="update userjoinclassinfo set classid='"&firstid&"' where classid='"&secondid&"' and userid='"&rs("userid")&"' "
       rrs.open SQL2,schooldb,1,3
    end if
    rss.close
    rs.movenext
wend
rs.close

sql="delete message where classid='"&secondid&"'"
rs.open SQL,schooldb,1,3

sql="delete upload where classid='"&secondid&"'"
rs.open SQL,schooldb,1,3

sql="delete uploadreply where classid='"&secondid&"'"
rs.open SQL,schooldb,1,3  
response.redirect "../success.asp?info=恭喜您，班级合并成功！"
set rs=nothing
set rrs=nothing
set rss=nothing
%>

     
