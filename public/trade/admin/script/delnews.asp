<%if session("admin_name")="" then response.end%>
<!--#include file="conn.asp"-->
<!--#include file="check.asp"--><head>
<link href=../css.css rel=STYLESHEET type=text/css>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
</head>
<%
set rs=server.createobject("adodb.recordset")
if request("action")="del" then
'    news_id=request("news_id")
    sql="select * from news where news_id="&request("news_id")
    rs.open sql,conn,3,3
    if rs.eof then
       response.redirect "delnews.asp"
    else
        rs.delete
        rs.update
        response.write "内容删除完毕"
        response.write "<Br>"
        response.write "<a href=delnews.asp>返回</a>"
    end if
    rs.close


else
%>




<table width="500" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>
      <div align="center">
        <center>
      <table border="1" width="550" cellspacing="0" bordercolor="#C0C0C0" style="border-collapse: collapse" cellpadding="0">
        <%
  page=request.querystring("page")
  if page="" then page=1
  if not(isnumeric(page)) then page=1
  if page<1 then page=1
  page=int(page)
  
  sql="select * from news order by news_id DESC"
  rs.open sql,conn,3,3
  if rs.eof then
  %>
        <tr bgcolor="#FFFFFF"> 
          <td width="55">编号</td>
          <td colspan="3">标题</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="4">没有内容</td>
        </tr>
        <%else
  rs.pagesize=20
  totalrec=rs.recordcount
  totalpage=rs.pagecount
  if page>totalpage then page=totalpage
  rs.absolutepage=page
  rs.cachesize=rs.pagesize
  i=0
  dim news_id(),news_title()
  do while not rs.eof and (i<rs.pagesize)
  i=i+1
  redim preserve news_title(i),news_id(i)
  news_id(i)=rs("news_id")
  news_title(i)=rs("news_title")
  rs.movenext
  loop
  rs.close
  %>
        <tr bgcolor="#6894d8"> 
          <td width="55" bgcolor="#808080"> 
            <div align="center"><font color="#FFFFFF">编号</font></div>
          </td>
          <td width="399" bgcolor="#808080"> 
            <div align="center"><font color="#FFFFFF">标题</font></div>
          </td>
          <td colspan="2" bgcolor="#808080"> 
            <div align="center"><font color="#FFFFFF">操作</font></div>
          </td>
        </tr>
        <%for i=1 to ubound(news_id)%>
        <tr bgcolor="#FFFFFF"> 
          <td width="55" height="20"><%=news_id(i)%></td>
          <td width="399" height="20"><%=news_title(i)%></td>
          <td width="43" height="20"> 
            <div align="center"><a href="delnews.asp?news_id=<%=news_id(i)%>&action=del" onClick="return confirmdel()">删除</a></div>
          </td>
          <td width="40" height="20"> 
            <div align="center"><a href="editnews.asp?news_id=<%=news_id(i)%>">修改</a></div>
          </td>
        </tr>
        <%next%>
        <%end if%>
      </table>
        </center>
      </div>
    </td>
  </tr>
</table>
<div align="center">共<font color=red><%=totalpage%></font>页 第<%=page%>页 <font color=666666> 
  <%if page-1>0 then%>
  <a href="delnews.asp?page=<%=page-1%>">上一页</a> 
  <%else%>
  <font color=666666>上一页</font> 
  <%end if%>
  　 
  <%if page+1<=totalpage then%>
  <a href="delnews.asp?page=<%=page+1%>">下一页</a> 
  <%else%>
  <font color=666666>下一页</font> 
  <%end if%>
  </font><br>
  <%end if
set rs=nothing
conn.close
set conn=nothing
%>
</div>