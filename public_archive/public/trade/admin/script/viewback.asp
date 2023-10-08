<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%response.Buffer=true%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>分页显示图片记录</title>
<link rel="stylesheet" href="upload.css" type="text/css">
<script language="JScript">
<!--
function confirmdel()
{
var fig=false
  fig=confirm("确定要删除该回复信息吗？")
  return fig
}

function check2(form2)
{
var fig
 if (form2.proname.value=="")
 {
  fig=false
  window.alert("请先输入客户名称")
  }
  return fig
}
function check1(form1)
{
var fig
if (form1.inputpage.value=="")
 {
  fig=false
  window.alert("请先输入要转向的页码数")
  }
  return fig
}
  //-->
</script>

</head>
<body>
<%
if session("admin_name")="" then 
  response.Redirect("../index.htm")
  response.End()
end if

dim rs,sql
dim iPage
set rs=server.CreateObject("adodb.recordset")
proname=request.form("proname")
if request.form("proname")<>"" then
sql="select * from back where ask_member like '%"&request.form("proname")&"%' order by id desc"
 else
  sql="select * from back order by id desc"
end if
rs.pagesize=12
rs.open sql,conn,1,1
inputpage=request.form("inputpage")
if not isnumeric(inputpage) then 
response.write "没有您所需要的信息，请正确输入页码数！"
response.Write "<a href='#' onclick='javascript:history.go(-1)'>返 回</a>"
response.End()
end if
if cint(inputpage)>cint(rs.pagecount) then 
response.write "没有您所需要的信息，请正确输入页码数！"
response.Write "<a href='#' onclick='javascript:history.go(-1)'>返 回</a>"
response.End()
end if
'上一句的函数cint不能少，否则会出错。
if len(inputpage)<>0 then
   iPage=cint(inputpage)
   rs.absolutepage=inputpage
else   
if len(request("page"))=0 then
  iPage=1 
  else
   iPage=request("page")
 end if
  if not rs.eof then rs.absolutepage=iPage
end if 

%>
<p align="center" class="dateformat"><font color="#000066" size="4">回 复 信 息<br>
  <img src="../../images/line_01.jpg" width="338" height="20"> </font></p>
<p align="center" class="dateformat">第<font color="#FF0000">
  <%if len(inputpage)=0 and isnumeric(inputpage) then%>
  <%=iPage%> 
  <%else%>
  <%=inputpage%> 
  <%end if%></font>
  页&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; 总共 <font color="#FF0000"><%=rs.recordcount%></font> 条信息&nbsp; &nbsp;总共 <font color="#FF0000"><%=rs.pagecount%></font> 页</p>


<table width="1022" height="58" border="0" align="center"  cellpadding="4" cellspacing="1" bgcolor="#293863">
  <tr bgcolor="#FFFFFF"> 
    <td width="55" height="29" class="chinese1"> <div align="center">询问的客户</div></td>
    <td width="64" class="chinese1"> <div align="center">询问的主题</div></td>
    <td width="93" class="chinese1"> <div align="center">询问的时间</div></td>
    <td width="91" class="chinese1"><div align="center">询问的产品</div></td>
    <td width="91" class="chinese1"><div align="center">询问的内容</div></td>
    <td width="142" class="chinese1"> <div align="center">回复主题</div></td>
    <td width="144" class="chinese1"><div align="center">回复时间</div></td>
    <td width="89" class="chinese1"><div align="center">回复内容</div></td>
    <td width="89" class="chinese1"><div align="center">回复者</div></td>
    <td width="73" class="chinese1"> <div align="center">相关操作</div></td>
  </tr>
  <%
  for i=1 to rs.pagesize
   if not rs.eof then
 %>
  <tr bgcolor="#fafbfd"> 
    <td height="26" class="chinese"> <div align="center"><%=rs("ask_member")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("ask_subject")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("ask_date")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("ask_product")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("ask_toknow")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("back_subject")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("back_date")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("back_content")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("back_admin")%></div></td>
    <td  class="chinese"><div align="center"> <a href="delback.asp?id=<%=rs("id")%>" onClick="return confirmdel()">删除</a></div></td>
  </tr>
  <%
 end if
 if not rs.eof then rs.movenext
 next
 %>
</table>
<p align="center"> 
  <%
if cint(iPage)=1 or cint(inputpage)=1 then
%>
  <font class="dateformat">第一页 | 上一页 | </font> 
  <%
else
%>
  <a href="viewback.asp?page=1">第一页</a> | <a href="viewback.asp?page=<%=iPage-1%>">上一页</a> 
  | 
  <%end if%>
  <%
 if cint(iPage)=cint(rs.pagecount) or cint(inputpage)=cint(rs.pagecount) then
%>
  <font class="dateformat">下一页 | 最后一页 </font>
  <%else%>
  <a href="viewback.asp?page=<%=iPage+1%>">下一页</a> | <a href="viewback.asp?page=<%=rs.pagecount%>">最后一页</a> 
  <%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
  <br>
<form name="form1" method="post" action="" onsubmit="return check1(this)">
  <div align="center" class="chinese">转到第 
    <input name="inputpage" type="text" size="3" class="input">
    页 
    <input type="submit" name="Submit" value="GO">
  </div>
</form>

<form name="form2" method="post" action="viewback.asp?ifsearch=1" onsubmit="return check2(this)">
  <table width="508" height="16" border="0" align="center" cellpadding="0" cellspacing="0" class="chinese">
    <tr>
      <td width="268" height="16">按客户ID搜索： 
        <input name="proname" type="text" id="proname" size="15" class="input" value=<%=proname%>></td>
      <td width="106"> <div align="left">
          <input type="submit" name="Submit2" value="开始搜索" class="button">
          &nbsp;&nbsp;&nbsp;&nbsp;|</div></td>
      <td width="120"><input type="button" name="Submit3" value="显示全部" class="button" onClick="javascript:location.href='viewback.asp'"></td>
      <td width="14">&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>
