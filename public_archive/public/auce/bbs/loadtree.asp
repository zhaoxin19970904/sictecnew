<!--#include file="CONN.ASP" -->
<!--#include file="ubbcode.asp" -->
<%
if request.QueryString("listonline")<>"" then
call listonlineuser()
else
call loadLTtree()
end if
sub loadLTtree()
dim id,rlyid,rs,i,sqltext
id=request.QueryString("id")
rlyid=request.QueryString("rlyid")
set rs=server.CreateObject("ADODB.Recordset")
sqltext=("select id,ly,user from rely where rid='"&id&"' order by top desc,id ")
rs.open sqltext,conn,1,1
response.Write("<link href=""DEFAULT.css"" rel=""stylesheet"" type=""text/css"">")
response.Write("<body background=""backimg/white.jpg"">")
response.Write("<table width=""100%"" border=0 bgcolor=#FFFFFF cellpadding=6 cellspacing=0>")
response.Write("<tr><td width=""5%"">&nbsp;</td>")
response.Write("<td width=""95%""><ol>")
if rs.eof or rs.bof then
response.Write("notree")
else
for i=1 to rs.recordcount
response.Write("<li><a href=type.asp?id="&id&"&rlyid="&rs("id").value&" target=""_parent"">"&ChkBadWords(left(rs("ly"),40))&"</a>...  --- "&rs("user")&"</li>")
rs.movenext
next
end if
response.Write("</ol></td> </tr></table>")
end sub
'显示在经用户列表
sub listonlineuser()
dim rs,lastout,i,title,href,img,ujibie,ip
i=1
set rs=conn.execute("select id,user,ip,time,ie,sys,jibie,address,where,hdtime from online order by id desc")
lastout="<table width=760 border=0 cellpadding=3 cellspacing=0><tr>"
do while not rs.eof 
if master then
img="<img src=pic/ao.gif>"
ip=rs("ip")
else
ip=" *.*.*.* "
end if 
title=("用户IP:"&ip&" "&rs("sys")&" <br>"&rs("ie")&"<br>上线时间:"&rs("time")&"<br>用户位置:"&rs("where")&"<br>活动时间:"&rs("hdtime"))
if rs("user")<>"客人" then
href="usertype.asp?user="&rs("user")
img="<img src=pic/ao1.gif>"
else
href="#"
img="<img src=pic/messages1.gif>"
end if
if cint(request.Cookies("rewin")("userid"))=rs("id") then
lastout=lastout&"<td>"&img&" <a href="&href&" title="""&title&"""><font color=#0000FF>"&rs("user")&"</font></a></td>"
else
lastout=lastout&"<td>"&img&" <a href="&href&" title="""&title&""">"&rs("user")&"</a></td>"
end if

  if (i mod 5 )=0 then
  lastout=lastout&"</tr><tr>"
  end if
rs.movenext
i=i+1
loop
lastout=lastout&"</tr></table>"
response.Write("<script>parent.document.all.listonline.innerHTML='"&lastout&"'</script>")
end sub
%>

