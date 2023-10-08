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
  fig=confirm("确定要删除该模型的所有信息吗？")
  return fig
}

function check2(form2)
{
var fig
 if (form2.proname.value=="")
 {
  fig=false
  window.alert("请先输入大类名称")
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
  response.Redirect("../../index.htm")
  response.End()
end if

dim rs,sql
dim iPage
set rs=server.CreateObject("adodb.recordset")
proname=request.form("proname")
if request.form("proname")<>"" then
sql="select * from modeldb where type like '%"&request.form("proname")&"%' order by update desc"
 else
  sql="select * from modeldb order by update desc"
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
<p align="center" class="datāeformat"><font color="#000066" size="4">所有已上传的模型图片<br>
  <img src="../../images/line_01.jpg" width="338" height="20"> </font></p>
<p align="center" class="dateformat">第<font color="#FF0000">
  <%if len(inputpage)=0 and isnumeric(inputpage) then%>
  <%=iPage%> 
  <%else%>
  <%=inputpage%> 
  <%end if%></font>
  页&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; 总共 <font color="#FF0000"><%=rs.recordcount%></font> 张模型图片&nbsp; &nbsp;总共 <font color="#FF0000"><%=rs.pagecount%></font> 页</p>


<table width="1581" height="58" border="0" align="center"  cellpadding="4" cellspacing="1" bgcolor="#293863">
  <tr bgcolor="#FFFFFF"> 
			    
    <td width="50" height="29" class="chinese1"> 
      <div align="center">Type</div></td>
                
    <td width="50" class="chinese1">
<div align="center">Name</div></td>
				
    <td width="131" class="chinese1">
<div align="center">Manufacturer</div></td>
				
    <td width="81" class="chinese1">
<div align="center">Country</div></td>
				
    <td width="42" class="chinese1">
<div align="center">Ratio</div></td>
				
    <td width="81" class="chinese1">
<div align="center">Model Size</div></td>
				
    <td width="65" class="chinese1">
<div align="center">Org-size</div></td>
				
    <td width="65" class="chinese1">
<div align="center">Material</div></td>
				
    <td width="42" class="chinese1">
<div align="center">Price</div></td>
                
    <td width="96" class="chinese1">
<div align="center">Miniquantity</div></td>
				
    <td width="65" class="chinese1">
<div align="center">Leadtime</div></td>
				
    <td width="35" class="chinese1">
<div align="center">Pack</div></td>
				
<td width="85" class="chinese1">
<div align="center">update</div></td>
				
    <td width="74" class="chinese1">
<div align="center">Small Photo</div></td>
				
    <td width="69" class="chinese1"> <div align="center">Big Photo</div></td>
				
    <td width="109" class="chinese1">
<div align="center">Sample Available</div></td>

<td width="73" class="chinese1">
<div align="center">Note</div></td>
				
    <td width="205" class="chinese1">
<div align="center">相关操作</div></td>
              </tr> 
  
 <%
  for i=1 to rs.pagesize
   if not rs.eof then
 %>
 
 
  <tr bgcolor="#fafbfd"> 
			    
       
    <td height="26" class="chinese"> 
      <div align="center"><%=rs("type")%></div></td>
       <td  class="chinese"><div align="center"><%=rs("name")%></div></td>
	   <td  class="chinese"><div align="center"><%=rs("manufacturer")%></div></td>
	   <td  class="chinese"><div align="center"><%=rs("original")%></div></td>
	   <td  class="chinese"><div align="center"><%=rs("ratio")%></div></td>
	   <td  class="chinese"><div align="center"><%=rs("size")%></div></td>
	   <td  class="chinese"><div align="center"><%=rs("org-size")%></div></td>
	   <td  class="chinese"><div align="center"><%=rs("material")%></div></td>
	   <td  class="chinese"><div align="center"><%=rs("price")%></div></td>
	   <td  class="chinese"><div align="center"><%=rs("miniquantity")%></div></td>
       <td  class="chinese"><div align="center"><%=rs("leadtime")%></div></td>
       <td  class="chinese"> <div align="center"><%=rs("packing")%></div></td>
	   <td  class="chinese"> <div align="center"><%=rs("update")%></div></td>
	   <td  class="chinese">
	   <div align="center">
	   <a href="photo1/<%=rs("photo1")%>" target="_blank">显示图片</a>
	   </div>
	   </td>
	   <td class="chinese"><div align="center">
	   <a href="photo2/<%=rs("photo2")%>" target="_blank">显示图片</a>
	   </div>
	   </td>
		<td  class="chinese"><div align="center"><%=rs("sample")%></div></td>
		<td  class="chinese"><div align="center">
		<a href="viewnote.asp?id=<%=rs("id")%>">查看内容</a>
		</div></td>
<td  class="chinese"><div align="center"> <a href="endup.asp?sign=<%=rs("sign")%>&deal=small" onClick="return confirmdel()">删除</a> 
        | <a href="editmodel.asp?id=<%=rs("id")%>">修改模型信息</a> 
        | <a href="upload.asp?id=<%=rs("id")%>">修改模型图片</a></div></td>
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
  <a href="viewmodel.asp?page=1">第一页</a> | <a href="viewmodel.asp?page=<%=iPage-1%>">上一页</a> 
  | 
  <%end if%>
  <%
 if cint(iPage)=cint(rs.pagecount) or cint(inputpage)=cint(rs.pagecount) then
%>
  <font class="dateformat">下一页 | 最后一页 </font>
  <%else%>
  <a href="viewmodel.asp?page=<%=iPage+1%>">下一页</a> | <a href="viewmodel.asp?page=<%=rs.pagecount%>">最后一页</a> 
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
    <input name="inputpage" type="text" size="3" class="input" style="ime-mode:disabled">
    页 
    <input type="submit" name="Submit" value="GO">
  </div>
</form>

<form name="form2" method="post" action="viewmodel.asp?ifsearch=1" onsubmit="return check2(this)">
  <table width="508" height="16" border="0" align="center" cellpadding="0" cellspacing="0" class="chinese">
    <tr>
      <td width="268" height="16">按模型大类(Type)搜索： 
        <input name="proname" type="text" id="proname" size="15" class="input" value=<%=proname%>></td>
      <td width="106"> <div align="left">
          <input type="submit" name="Submit2" value="开始搜索" class="button">
          &nbsp;&nbsp;&nbsp;&nbsp;|</div></td>
      <td width="120"><input type="button" name="Submit3" value="显示全部" class="button" onClick="javascript:location.href='viewmodel.asp'"></td>
      <td width="14">&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>
