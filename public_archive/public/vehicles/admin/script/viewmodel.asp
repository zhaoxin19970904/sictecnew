<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<!--#include file="../../script/check.asp"-->
<%response.Buffer=true%>
<%
if session("admin_name")="" then response.end
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ҳ��ʾ���ݿ��¼</title>
<script language="JScript">
<!--
function confirmdel()
{
var fig=false
  fig=confirm("ȷ��Ҫɾ�����ͺ�������Ϣ��\n���ɾ�������������������ͺŶ�Ӧ�Ĳ�Ʒ\n�Ͳ�Ʒ�۸�һ��ɾ����\nȷ����")
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
<style type="text/css">
<!--
.input {
	border:1px solid #666666; padding:1px; FONT-SIZE: 12px; COLOR: #383805; FONT-FAMILY: "Verdana", "Arial", "����"; HEIGHT: 15px; BACKGROUND-COLOR: #ffffff
}
.dateformat {
	font-size: 13px;
}

     a:link {text-decoration:none; color:#383805; font-size=12px}
	 a:visited {text-decoration:none; color:#383805;font-size=12px}
     a:active {text-decoration:underline; color:#383805}
     a:hover {text-decoration:underline; color:#FF0033}
-->
</style>
</head>
<body>
<%
dim cn,rs,dbfile,sql
dim iPage
'dbfile=server.MapPath("allinfo.mdb")
'set cn=server.createobject("adodb.connection")
'cn.open "driver={microsoft access driver (*.mdb)};dbq="&dbfile
set rs=server.CreateObject("adodb.recordset")
if request.QueryString("change")<>"" then
  select1=request("select")
  sql="select * from model where prodir_name='"&select1&"'"
else
  sql="select * from model order by prodir_name desc"
end if
rs.pagesize=15
rs.open sql,conn,1,1
'rs.open sql,cn,1,1

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



'findprodir_name=request.QueryString("findprodir_name")
%>
<form name="form2" method="post" action="viewmodel.asp?change=change">
  <div align="center" class="dateformat">������ͣ� 
    <select name="select">
	 <%
	 sqldir="select * from prodir"
    set rsdir=server.createobject("adodb.recordset")
    rsdir.open sqldir,conn,1,3
	if rsdir.eof or rsdir.bof then
		 response.write "<option value='" +  "���������" + "'>" + "���������" + "</option>"
	  else
         Do while not rsdir.eof
            response.write "<option value='" +  trim(rsdir("prodir_name")) + "'>" + rsdir("prodir_name") + "</option>"
			'response.write "<option value='" +  trim(rsdir("prodir_name")) + "'" + " " + if trim(rsdir("prodir_name"))=findprodir_name then selected + ">" + rsdir("prodir_name") + "</option>"
            rsdir.MoveNext
         Loop
		 
		 if findprodir_name<>"" then
		    document.form2.select.value=findprodir_name
			end if
		 
		 
		 
	end if
	rsdir.close
	set rsdir=nothing
   %>
    </select>
    <input type="submit" name="Submit3" value="��ʼɸѡ" onClick="document.form2.findprodir_name.value=document.form2.select.options[document.form2.select.selectedIndex].value">
    <input type="hidden" name="findprodir_name" value="findprodir">
  </div>
</form>


<p align="center" class="dateformat">��<font color="#FF0000">
  <%if len(inputpage)=0 and isnumeric(inputpage) then%>
  <%=iPage%> 
  <%else%>
  <%=inputpage%> 
  <%end if%></font>
  ҳ&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; �ܹ� <font color="#FF0000"><%=rs.recordcount%></font> ������&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;�ܹ� <font color="#FF0000"><%=rs.pagecount%></font> ҳ</p>


  <table width="648" height="50" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC" bgcolor="fafbfd" class="dateformat">
    <tr> 
      <td width="115" height="18"><div align="center">�������</div></td>
      <td width="121"><div align="center">�ͺ���</div></td>
      <td width="153"><div align="center">����ʱ��</div></td>
      <td width="67"><div align="center">����</div></td>
      <td width="120"><div align="center">����</div></td>
      <td width="58"><div align="center">����</div></td>
    </tr>
    <%
  for i=1 to rs.pagesize
   if not rs.eof then
      
 %>
<form name="form3" method="post" action="changmodel.asp?model_name=<%=rs("model_name")%>&prodir_name=<%=rs("prodir_name")%>">
 
    <tr onmouseover="this.style.backgroundColor='#ffffff';" onmouseout="this.style.backgroundColor='';"> 
      <td height="25"><div align="center"><%=rs("prodir_name")%></div></td>
      <td>
	    <div align="center">
          <input name="model_namechg" type="text" size="15" value="<%=rs("model_name")%>">
        </div>
	 </td>
      <td><div align="center"><%=rs("model_date")%></div></td>
	  
      <td> <div align="center"><a href="addmodel.asp">���</a> </div></td>
      <td><div align="center"> 
          <input type="submit" name="Submit2" value="�޸�" onClick="return edit()">
        </div></td>
      <td><div align="center"><a href="delmodel.asp?model_name=<%=rs("model_name")%>&prodir_name=<%=rs("prodir_name")%>" onClick="return confirmdel()">ɾ��</a></div></td>
      <%'�Ͼ��delete������������������%>
    </tr>
	</form>
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
  <a href="viewmodel.asp?page=1">��һҳ</a> | <a href="viewmodel.asp?page=<%=iPage-1%>">��һҳ</a> 
  | 
  <%end if%>
  <%
 if cint(iPage)=cint(rs.pagecount) or cint(inputpage)=cint(rs.pagecount) then
%>
  <font class="dateformat">��һҳ | ���һҳ </font> 
  <%else%>
  <a href="viewmodel.asp?page=<%=iPage+1%>">��һҳ</a> 
  | <a href="viewmodel.asp?page=<%=rs.pagecount%>">���һҳ</a> 
  <%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
'cn.close
'set cn=nothing
%>
  <br>
<form name="form1" method="post" action="" onsubmit="return check1(this)">
  <div align="center" class="dateformat">ת���� 
    <input name="inputpage" type="text" size="3" class="input" style="ime-mode:disabled">
    ҳ 
    <input type="submit" name="Submit" value="GO">
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="viewmodel.asp">��ʾȫ�� </a></div>
</form>
</body>
</html>
