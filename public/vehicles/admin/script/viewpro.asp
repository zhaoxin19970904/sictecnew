<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%response.Buffer=true%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>View Product Photo</title>
<link rel="stylesheet" href="upload.css" type="text/css">
<script language="JScript">
<!--
function confirmdel()
{
var fig=false
  fig=confirm("Make sure ？")
  return fig
}

function check2(form2)
{
var fig
 if (form2.proname.value=="")
 {
  fig=false
  window.alert("Please input the Type!")
  }
  return fig
}
function check1(form1)
{
var fig
if (form1.inputpage.value=="")
 {
  fig=false
  window.alert("Please input the page!")
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
sql="select * from products where type like '%"&request.form("proname")&"%' order by up_date desc"
 else
  sql="select * from products order by up_date desc"
end if
rs.pagesize=15
rs.open sql,conn,1,1
inputpage=request.form("inputpage")
if not isnumeric(inputpage) then 
response.write "No information，Please input right page！"
response.Write "<a href='#' onclick='javascript:history.go(-1)'>Back</a>"
response.End()
end if
if cint(inputpage)>cint(rs.pagecount) then 
response.write "No information，Please input right page！"
response.Write "<a href='#' onclick='javascript:history.go(-1)'>Back</a>"
response.End()
end if
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
<p align="center" class="datāeformat"><font color="#000066" size="4">All Product 
  Photo <br>
  <img src="../../images/line_01.jpg" width="338" height="20"> </font></p>
<p align="center" class="dateformat">Page No.<font color="#FF0000"> 
  <%if len(inputpage)=0 and isnumeric(inputpage) then%>
  <%=iPage%> 
  <%else%>
  <%=inputpage%> 
  <%end if%>
  </font> &nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; Total:<font color="#FF0000"><%=rs.recordcount%></font> Procucts&nbsp; &nbsp;Total:<font color="#FF0000"><%=rs.pagecount%></font> Pages</p>


<table width="1190" height="58" border="0" align="center"  cellpadding="4" cellspacing="1" bgcolor="#293863">
  <tr bgcolor="#FFFFFF"> 
    <td width="115" height="29" class="chinese1"> <div align="center">Type</div></td>
    <td width="97" class="chinese1"> <div align="center">Model</div></td>
    <td width="132" class="chinese1"><div align="center">Price</div></td>
    <td width="132" class="chinese1"><div align="center">Up Date</div></td>
    <td width="86" class="chinese1"> <div align="center">Up Mnanger</div></td>
    <td width="79" class="chinese1"> <div align="center">Porduct Photo</div></td>
    <td width="90" class="chinese1"> <div align="center">Introduction</div></td>
    <td width="386" class="chinese1"> <div align="center">Operate</div></td>
  </tr>
  <%
  for i=1 to rs.pagesize
   if not rs.eof then
 %>
  <tr bgcolor="#fafbfd"> 
    <td height="26" class="chinese"> <div align="center"><%=rs("type")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("modelno")%></div></td>
    <td  class="chinese"><div align="center">&yen; <%=rs("price")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("up_date")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("up_admin")%></div></td>
    <td  class="chinese"> <div align="center"> <a href="../../proimages/<%=rs("photo")%>" target="_blank">View 
        Photo</a> </div></td>
    <td  class="chinese"><div align="center"> <a href="viewnote.asp?id=<%=rs("id")%>">View 
        Content</a> </div></td>
    <td  class="chinese"> <a href="endup.asp?id=<%=rs("id")%>" onClick="return confirmdel()">Delete</a> 
      | <a href="editproinfo.asp?id=<%=rs("id")%>">Edit 
      pro info</a> | <a href="editphoto.asp?id=<%=rs("id")%>">Edit 
      pro photo</a> | 
      <%if rs("ifcommended")=false then%> <a href="deal.asp?id=<%=rs("id")%>&action=tj">Commend</a> 
      <%else%> <a href="deal.asp?id=<%=rs("id")%>&action=cxtj"> 
      <font color="#ff0000">Quash Commend</font> </a> <%end if%> </td>
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
  <font class="dateformat">First | Back | </font> 
  <%
else
%>
  <a href="viewpro.asp?page=1">First</a> | <a href="viewpro.asp?page=<%=iPage-1%>">Back</a> 
  | 
  <%end if%>
  <%
 if cint(iPage)=cint(rs.pagecount) or cint(inputpage)=cint(rs.pagecount) then
%>
  <font class="dateformat">Next | Last</font> 
  <%else%>
  <a href="viewpro.asp?page=<%=iPage+1%>">Next</a> | <a href="viewpro.asp?page=<%=rs.pagecount%>">Last</a> 
  <%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
  <br>
<form name="form1" method="post" action="" onsubmit="return check1(this)">
  <div align="center" class="chinese">GOTO 
    <input name="inputpage" type="text" size="3" class="input" style="ime-mode:disabled">
    Page 
    <input type="submit" name="Submit" value="GO">
  </div>
</form>

<form name="form2" method="post" action="viewpro.asp?ifsearch=1" onsubmit="return check2(this)">
  <table width="567" height="16" border="0" align="center" cellpadding="0" cellspacing="0" class="chinese">
    <tr>
      <td width="268" height="16">Search from Type： 
        <input name="proname" type="text" id="proname" size="15" class="input" value=<%=proname%>></td>
      <td width="140"> <div align="left">
          <input type="submit" name="Submit2" value="Start Search" class="button">
          &nbsp;&nbsp;&nbsp;&nbsp;|</div></td>
      <td width="125"><input type="button" name="Submit3" value="View All" class="button" onClick="javascript:location.href='viewpro.asp'"></td>
      <td width="34">&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>
