<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit %>
<!--#include file="conn.asp"-->
<%dim news_id,news_title,news_content,sql,rs
news_id=Request("news_id")
news_title=Request("news_title")
news_content=Request("news_content")
set rs=server.createobject("adodb.recordset")
sql="select * from news where news_id="&request("news_id")
rs.open sql,conn,3,3
rs("news_title")=news_title
rs("news_content")=news_content
rs.Update
response.write "商品类别修改成功!"
    response.write "<br>"
    response.write "<a href=delnews.asp>返回</a>"
    rs.close
%>