<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
firstid=trim(request.form("firstid"))
secondid=trim(request.form("secondid"))
set rs = createobject("ADODB.recordset")
if firstid="" or secondid="" then
  response.redirect "../error.asp?info=�Բ�������д������"
end if
if firstid=secondid then
  response.redirect "../error.asp?info=�Բ��������༶������һ����"
end if
sql="select * from classinfo where classid='"&firstid&"'"
rs.open SQL,schooldb
if not rs.eof then
   firstname=rs("classname")
   firstschoolid=rs("schoolid")
else
   response.redirect "../error.asp?info=�Բ��𣬺ϲ���༶id�����ڣ�"
end if
rs.close
sql="select * from classinfo where classid='"&secondid&"'"
rs.open SQL,schooldb
if not rs.eof then
   secondname=rs("classname")
   secondschoolid=rs("schoolid")
else
   response.redirect "../error.asp?info=�Բ��𣬱��ϲ��༶id�����ڣ�"
end if
rs.close   
sql="select * from schoolinfo where schoolid='"&firstschoolid&"'"
rs.open SQL,schooldb
if not rs.eof then
   firstschoolname=rs("schoolname")
end if
rs.close
sql="select * from schoolinfo where schoolid='"&secondschoolid&"'"
rs.open SQL,schooldb
if not rs.eof then
   secondschoolname=rs("schoolname")
end if
rs.close   
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#DFDFDF">
<p>&nbsp;</p>
<table width="75%" border="0" align="center" height="50">
  <tr> 
    <td width="50%"> 
      <div align="center"><b>�ϲ��༶ </b></div>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
 <form name="form1" method="post" action="class_action.asp">
<table width="75%" border="1" align="center" height="40%">
 
    <tr> 
      <td width="50%" height="151"> 
        <div align="center">�ϲ���İ༶���ƣ�</div>
      </td>
      <td width="50%" height="151"> 
        <div align="center">���ϲ��İ༶���ƣ�</div>
      </td>
    </tr>
    <tr> 
      <td width="50%" height="122"> 
        <div align="center"> <%=firstschoolname%><%=firstname%> </div>
      </td>
      <td width="50%" height="122"> 
        <div align="center"> <%=secondschoolname%><%=secondname%> </div>
      </td>
    </tr>

</table>
<table width="75%" border="1" align="center" height="60">
  <tr> 
    <td> 
      <div align="center">
        <input type="hidden" name="firstid" value="<%=firstid%>">
        <input type="hidden" name="secondid" value="<%=secondid%>">
        <input type="submit" name="Submit" value="��ȷ��Ҫ�ϲ��������༶">
      </div>
    </td>
  </tr>
</table>
  </form>
<p>&nbsp;</p>
</body>
</html>