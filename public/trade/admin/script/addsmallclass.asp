<%
if session("admin_name")="" then response.end
%>
<!--#include file="conn.asp"-->
<!--#include file="check.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css>
.unnamed1 {
	FONT-SIZE: 9pt
}
</STYLE>
</head>
<%
    sql="select * from main"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
%>
<body topmargin="50">
<div align="center"><font color="#000066" size="5">产品小类管理</font> <br>
  <img src="../../images/line_01.jpg" width="338" height="20"></div>
<form method="POST" action="dealclass.asp?action=smallclass" onsubmit="return addnew()">
  <table border="0" cellpadding="0" cellspacing="0" width="44%" align="center"  class="unnamed1">

    <tr>
      <td height="31"> 
        <p align="center">在产品大类： 
          <select name="subject" size="1" style="font-size: 9pt">
            <%
	if rs.eof or rs.bof then
		 response.write "<option value='" +  "请增加类别" + "'>" + "请增加类别" + "</option>"
	  else
         Do while not rs.eof
             response.write "<option value='" +  trim(rs("main_id")) + "'>" + rs("main_type") + "</option>"
             rs.MoveNext
         Loop
	end if
%>
          </select>
          中 </td>
    </tr>
    <tr align="center">
      <td>
      </td>
    </tr>
    <tr align="center">
       
      <td height="34"> 
        <p>添加新的小类： 
          <input name="newclass" type="text" id="newclass" size="20">
          <input type="submit" value="确定" name="B1">
      </td>
    </tr>
    <tr align="center">
      <td>
      </td>
    </tr>
    <tr align="center">
      <td><p>&nbsp;
      </td>
    </tr>
  </table>
  <div align="center"><img src="../../images/line_01.jpg" width="338" height="20"> 
  </div>
</form>
<%
    rs.close
    sql="select distinct * from type"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
%>
<form method="POST" action="changetype.asp" onsubmit="return rename(this)">
  <table border="0" cellpadding="0" cellspacing="0" width="44%" align="center"  class="unnamed1">
    <tr> 
      <td height="27"> 
        <p align="center">把产品小类： 
          <select name="type_id" size="1" style="font-size: 9pt">
            <%
			rs.movefirst
	if rs.eof or rs.bof then
		 response.write "<option value='" +  "请增加类别" + "'>" + "请增加类别" + "</option>"
	  else
         Do while not rs.eof
             response.write "<option value='" +  trim(rs("type_id")) + "'>" + rs("type") + "</option>"
             rs.MoveNext
         Loop
	end if
%>
          </select>
          的名称
      </td>
    </tr>
    <tr align="center"> 
      <td> </td>
    </tr>
    <tr align="center"> 
      <td height="37"> 
        <p>重命名为： 
          <input name="newclass" type="text" id="newclass" size="20">
          <input type="submit" value="确定" name="B12">
      </td>
    </tr>
    <tr align="center"> 
      <td> </td>
    </tr>
    <tr align="center"> 
      <td><p>&nbsp; </td>
    </tr>
  </table>
  <div align="center"><img src="../../images/line_01.jpg" width="338" height="20"> 
  </div>
</form>

<form method="POST" action="changetype.asp" onsubmit="return del()">
  <table border="0" cellpadding="0" cellspacing="0" width="44%" align="center"  class="unnamed1">
    <tr> 
      <td height="27"> <p align="center">
          <input type="hidden" name="del" value="delete">
          把产品小类： 
          <select name="type_id" size="1" style="font-size: 9pt">
            <%
			rs.movefirst
	if rs.eof or rs.bof then
		 response.write "<option value='" +  "请增加类别" + "'>" + "请增加类别" + "</option>"
	  else
         Do while not rs.eof
             response.write "<option value='" +  trim(rs("type_id")) + "'>" + rs("type") + "</option>"
             rs.MoveNext
         Loop
	end if
%>
          </select>
          <input type="submit" value="删除" name="B122">
      </td>
    </tr>
    <tr align="center"> 
      <td> </td>
    </tr>
    <tr align="center"> 
      <td height="12"> 
        <p>&nbsp; </td>
    </tr>
    <tr align="center"> 
      <td> </td>
    </tr>
    <tr align="center"> 
      <td><p>&nbsp; </td>
    </tr>
  </table>
</form>


<%
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
%>
<script language="JavaScript">
function rename(form2)
{
   if(form2.newclass.value=="")
    {
	  alert("请输入新名称!");
	  form2.newclass.focus();
	  return false;
	}
	else
	{
    var fig;
	fig=confirm("确定要重命名该类吗?");
	return fig;
	}
}
</script>
<script language="JavaScript">
function del()
{
    var fig;
	fig=confirm("确定要删除该类吗?");
	return fig;
}
</script>