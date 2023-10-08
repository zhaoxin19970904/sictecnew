<%
if session("admin_name")="" then response.end
%>
<!--#include file="conn.asp"--><head>
<style type=text/css>
td  { font:normal 12px 宋体; }
img  { vertical-align:bottom; border:0px; }
a  { font:normal 12px 宋体; color:#000000; text-decoration:none; }
a:hover  { color:#428EFF;text-decoration:underline; }
.sec_menu  { border-left:1px solid white; border-right:1px solid white; border-bottom:1px solid white; overflow:hidden; background:#D6DFF7; }
.menu_title  { }
.menu_title span  { position:relative; top:2px; left:8px; color:#215DC6; font-weight:bold; }
.menu_title2  { }
.menu_title2 span  { position:relative; top:2px; left:8px; color:#428EFF; font-weight:bold; }
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<%
set rs=server.createobject("adodb.recordset")
if request("action")="save" then
    news_title=trim(request("news_title"))
    news_content=trim(request("news_content"))
    if news_title="" or news_content="" then
        response.write "输入数据不能为空"
        response.write "<br>"
        response.write "<a href=addnews.asp>返回</a>"
    else
        news_title=server.htmlencode(news_title)
        news_content=server.htmlencode(news_content)
        news_content=replace(news_content," ","&nbsp;")
        news_content=replace(news_content,chr(13)&chr(10),"<br>")
        sql="select * from news"
        rs.open sql,conn,3,3
        rs.addnew
        rs("news_title")=news_title
        rs("news_content")=news_content
		rs("news_upuser")=session("admin_name")
		rs("news_date")=now
    
        rs.update
        rs.close
        response.write "新闻添加完成"
        response.write "<br>"
        response.write "<a href=addnews.asp>返回"
    end if
else%>
<body bgcolor="#CCCCCC">
<table border="0" width="100%" cellspacing="1">
  <tr>
    <td width="100%">
      <form method="POST" action="addnews.asp?action=save">
        <table border="0" width="100%" cellspacing="1">
          <tr>
            <td width="100%">　</td>
          </tr>
          <tr>
            <td width="100%" valign="bottom"><font color="#000066">新闻标题</font>&nbsp;&nbsp;&nbsp;
<input type="text" name="news_title" size="20" class=input>
            </td>
          </tr>
          <tr>
            <td width="100%"><font color="#000066">新闻内容</font></td>
          </tr>
          <tr>
            <td width="100%"><textarea rows="14" name="news_content" cols="79" class=input></textarea></td>
          </tr>
        </table>
        <p><input type="submit" value="提交" name="B1" class=input><input type="reset" value="全部重写" name="B2" class=input></p>
      </form>
      <p>　</td>
  </tr>
</table>
<%end if
set rs=nothing
conn.close
set conn=nothing
%>