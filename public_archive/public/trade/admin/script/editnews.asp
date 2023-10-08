<%
if session("admin_name")="" then response.end
%>
<!--#include file="conn.asp"-->
<%
dim news_id,news_title,news_content
set rs=server.createobject("adodb.recordset")
sql="select * from news where news_id="&request("news_id") 
rs.open sql,conn,1,1
      news_id=rs("news_id")
      news_title=rs("news_title")
      news_content=rs("news_content")
      
rs.close
%><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href=../css.css rel=STYLESHEET type=text/css>
</head>



<table width="500" border="0" cellspacing="0" cellpadding="0" align="center" height="100">
  <tr> 
    <td height="175"> 
      <form name="form1" method="post" action="editnewsok.asp?news_id=<%=news_id%>">
        <table border="1" width="100%" cellspacing="0" bordercolor="#C0C0C0" style="border-collapse: collapse" cellpadding="0">
          <tr bgcolor="#FFFFFF"> 
            <td colspan="2"> 标题
<input type="text" name="news_title" size="60" class=input value="<%=news_title%>">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td colspan="2"> <font color="#FF0000">新闻内容</font></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td valign="top" colspan="2"> 
              <textarea rows="15" name="news_content" cols="70" class=input><%=news_content%></textarea>
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td colspan="2">&nbsp; </td>
          </tr>
          <tr bgcolor="#FFFFFF" align="center"> 
            <td colspan="2"> 
              <input type="submit" name="Submit" value="修改">
              &nbsp;&nbsp;&nbsp;&nbsp; &lt;&lt; 
              <a href="#" onClick="javascript:history.back()">返 回</a> &gt;&gt; </td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
</table>