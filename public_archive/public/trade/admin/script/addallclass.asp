
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
    sql="select * from title"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
%>
<body topmargin="20">
<div align="center"><font color="#000066" size="5">��Ʒ�������</font> <br>
  <img src="../../images/line_01.jpg" width="338" height="20"></div>
<form method="POST" action="dealclass.asp?action=bigclass" onsubmit="return addnew()">
  <table border="0" cellpadding="0" cellspacing="0" width="44%" align="center"  class="unnamed1">

    <tr>
      <td height="33"> 
        <p align="center">�ڲ�Ʒ���ࣺ 
          <select name="subject" size="1" style="font-size: 9pt">
                <%
	if rs.eof or rs.bof then
		 response.write "<option value='" +  "���������" + "'>" + "���������" + "</option>"
	  else
         Do while not rs.eof
             response.write "<option value='" +  trim(rs("title_id")) + "'>" + rs("title_type") + "</option>"
             rs.MoveNext
         Loop
	end if
%>
              </select>
          ��
      </td>
    </tr>
    <tr align="center">
      <td>
      </td>
    </tr>
    <tr align="center">
       
      <td height="42"> 
        <p>����µĴ��ࣺ 
          <input name="newclass" type="text" id="newclass" size="20">
          <input type="submit" value="ȷ��" name="B1">
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
  <p align="center"><img src="../../images/line_01.jpg" width="338" height="20"></p>
</form>

<%
    rs.close
    sql="select * from main"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
	dim main_id(),main_type()
	i=1
	do while not rs.eof 
	  redim preserve main_id(i),main_type(i)
	  main_id(i)=rs("main_id")
	  main_type(i)=rs("main_type")
	  if not rs.eof then rs.movenext
	  i=i+1
	loop
%>
<form method="POST" action="changemain.asp" onsubmit="return rename(this)">
  <table border="0" cellpadding="0" cellspacing="0" width="44%" align="center"  class="unnamed1">
    <tr> 
      <td height="32"> 
        <p align="center">�Ѳ�Ʒ���ࣺ 
          <select name="main_id" size="1" style="font-size: 9pt">
            <%
			rs.movefirst
	if rs.eof or rs.bof then
		 response.write "<option value='" +  "���������" + "'>" + "���������" + "</option>"
	  else
         'Do while not rs.eof
		 for j=1 to ubound(main_id)
             response.write "<option value='" +  trim(main_id(j)) + "'>" + main_type(j) + "</option>"
             'rs.MoveNext
         'Loop
		 next
	end if
%>
          </select>
          ������
      </td>
    </tr>
    <tr align="center"> 
      <td> </td>
    </tr>
    <tr align="center"> 
      <td height="40"> 
        <p>������Ϊ�� 
          <input name="newclass2" type="text" id="newclass2" size="20">
          <input type="submit" value="ȷ��" name="B12">
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


<form method="POST" action="changemain.asp" onsubmit="return del()">
  <table border="0" cellpadding="0" cellspacing="0" width="44%" align="center"  class="unnamed1">
    <tr> 
      <td height="23"> <p align="center">
          <input type="hidden" name="del" value="delete">
          �Ѳ�Ʒ���ࣺ 
          <select name="main_id" size="1" style="font-size: 9pt">
            <%
			'rs.movefirst
	if rs.eof or rs.bof then
		 response.write "<option value='" +  "���������" + "'>" + "���������" + "</option>"
	  else
         'Do while not rs.eof
		 for j=1 to ubound(main_id)
             response.write "<option value='" +  trim(main_id(j)) + "'>" + main_type(j) + "</option>"
             'rs.MoveNext
         'Loop
		 next
	end if
%>
          </select>
          <input type="submit" value="ɾ��" name="B122">
      </td>
    </tr>
    <tr align="center"> 
      <td> </td>
    </tr>
    <tr align="center"> 
      <td height="12"> <p>&nbsp; </td>
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
   if(form2.newclass2.value=="")
    {
	  alert("������������!");
	  form2.newclass2.focus();
	  return false;
	}
	else
	{
    var fig;
	fig=confirm("ȷ��Ҫ������������?");
	return fig;
	}
}
</script>
<script language="JavaScript">
function del()
{
    var fig;
	fig=confirm("ȷ��Ҫɾ��������?");
	return fig;
}
</script>
