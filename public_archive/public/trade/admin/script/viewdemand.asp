<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<%response.Buffer=true%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ҳ��ʾͼƬ��¼</title>
<link rel="stylesheet" href="upload.css" type="text/css">
<script language="JScript">
<!--
function confirmdel()
{
var fig=false
  fig=confirm("ȷ��Ҫɾ����������Ϣ��")
  return fig
}

function check2(form2)
{
var fig
 if (form2.proname.value=="")
 {
  fig=false
  window.alert("��������ͻ�����")
  }
  return fig
}
function check1(form1)
{
var fig
if (form1.inputpage.value=="")
 {
  fig=false
  window.alert("��������Ҫת���ҳ����")
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
sql="select * from demand where memberid like '%"&request.form("proname")&"%' order by time desc"
 else
  sql="select * from demand order by time desc"
end if
rs.pagesize=12
rs.open sql,conn,1,1
inputpage=request.form("inputpage")
if not isnumeric(inputpage) then 
response.write "û��������Ҫ����Ϣ������ȷ����ҳ������"
response.Write "<a href='#' onclick='javascript:history.go(-1)'>�� ��</a>"
response.End()
end if
if cint(inputpage)>cint(rs.pagecount) then 
response.write "û��������Ҫ����Ϣ������ȷ����ҳ������"
response.Write "<a href='#' onclick='javascript:history.go(-1)'>�� ��</a>"
response.End()
end if
'��һ��ĺ���cint�����٣���������
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
<p align="center" class="dat��eformat"><font color="#000066" size="4">�ͻ���������Ϣ<br>
  <img src="../../images/line_01.jpg" width="338" height="20"> </font></p>
<p align="center" class="dateformat">��<font color="#FF0000">
  <%if len(inputpage)=0 and isnumeric(inputpage) then%>
  <%=iPage%> 
  <%else%>
  <%=inputpage%> 
  <%end if%></font>
  ҳ&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; �ܹ� <font color="#FF0000"><%=rs.recordcount%></font> ����Ϣ&nbsp; &nbsp;�ܹ� <font color="#FF0000"><%=rs.pagecount%></font> ҳ</p>


<table width="1424" height="58" border="0" align="center"  cellpadding="4" cellspacing="1" bgcolor="#293863">
  <tr bgcolor="#FFFFFF"> 
    <td width="59" height="29" class="chinese1"> <div align="center">Member</div></td>
    <td width="69" class="chinese1"> <div align="center">Type</div></td>
    <td width="100" class="chinese1"> <div align="center">Products Name</div></td>
    <td width="98" class="chinese1"><div align="center">Expected Price</div></td>
    <td width="153" class="chinese1"> <div align="center">Expected delivery time</div></td>
    <td width="155" class="chinese1"><div align="center">Search Time</div></td>
    <td width="96" class="chinese1"><div align="center">Contact Name</div></td>
    <td width="96" class="chinese1"><div align="center">Contact Telphone</div></td>
    <td width="96" class="chinese1"><div align="center">E-mail</div></td>
    <td width="96" class="chinese1"> <div align="center">Demand IP</div></td>
    <td width="96" class="chinese1"><div align="center">If Write Bback </div></td>
    <td width="96" class="chinese1"><div align="center">Note</div></td>
    <td width="96" class="chinese1"> <div align="center">��ز���</div></td>
  </tr>
  <%
  for i=1 to rs.pagesize
   if not rs.eof then
 %>
  <tr bgcolor="#fafbfd"> 
    <td height="26" class="chinese"> <div align="center"><%=rs("memberid")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("type")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("pro_name")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("price")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("time")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("demand_time")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("name")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("tel")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("email")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("demand_ip")%></div></td>
    <td  class="chinese"><div align="center"><%=rs("feedback")%></div></td>
    <td  class="chinese"><div align="center"><a href="viewclientnote.asp?id=<%=rs("id")%>&table=demand">View</a></div></td>
    <td  class="chinese"><div align="center"> 
	<%if rs("feedback")=false then%>
        <a href="backdemand.asp?id=<%=rs("id")%>">�ظ�</a> 
        |
<%else%>
	�ѻظ�
|	<%end if%> 
	<a href="dealclient.asp?id=<%=rs("id")%>&table=demand" onClick="return confirmdel()">ɾ��</a></div></td>
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
  <font class="dateformat">��һҳ | ��һҳ | </font> 
  <%
else
%>
  <a href="viewdemand.asp?page=1">��һҳ</a> | <a href="viewdemand.asp?page=<%=iPage-1%>">��һҳ</a> 
  | 
  <%end if%>
  <%
 if cint(iPage)=cint(rs.pagecount) or cint(inputpage)=cint(rs.pagecount) then
%>
  <font class="dateformat">��һҳ | ���һҳ </font>
  <%else%>
  <a href="viewdemand.asp?page=<%=iPage+1%>">��һҳ</a> | <a href="viewdemand.asp?page=<%=rs.pagecount%>">���һҳ</a> 
  <%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
  <br>
<form name="form1" method="post" action="" onsubmit="return check1(this)">
  <div align="center" class="chinese">ת���� 
    <input name="inputpage" type="text" size="3" class="input">
    ҳ 
    <input type="submit" name="Submit" value="GO">
  </div>
</form>

<form name="form2" method="post" action="viewdemand.asp?ifsearch=1" onsubmit="return check2(this)">
  <table width="508" height="16" border="0" align="center" cellpadding="0" cellspacing="0" class="chinese">
    <tr>
      <td width="268" height="16">���ͻ�ID������ 
        <input name="proname" type="text" id="proname" size="15" class="input" value=<%=proname%>></td>
      <td width="106"> <div align="left">
          <input type="submit" name="Submit2" value="��ʼ����" class="button">
          &nbsp;&nbsp;&nbsp;&nbsp;|</div></td>
      <td width="120"><input type="button" name="Submit3" value="��ʾȫ��" class="button" onClick="javascript:location.href='viewdemand.asp'"></td>
      <td width="14">&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>
