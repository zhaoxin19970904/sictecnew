<!--#include file="CONN.ASP" -->
<%
if not userislogin then
errmess=errmess&error1
call founderror(errmess)
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=loadcopyc("ltname")%>--����Ϣ</title>
<link href="DEFAULT.css" rel="stylesheet" type="text/css">
</head>

<body  onkeydown="if(event.keyCode==13 && event.ctrlKey)messager.submit()">
<form language=javascript name="messager" action=messanger.asp method=post>
  <table width="898" border="0" align=center cellpadding=3 cellspacing=1 bgcolor="#92b9fb" class=tableborder1>
    <tbody>
      <tr> 
        <td height="27" colspan=3 background="backimg/bg1.gif"> <b><font color="#FFFFFF">���Ͷ���Ϣ 
          </font></b> </td>
      <tr bgcolor="#FFFFFF"> 
        <td valign=top class=tablebody1>������</td>
        <td valign=center class=tablebody1>
          <input name="touser" type="text" id="touser" size="50" maxlength="200">
          <SELECT name=fonta onchange="document.messager.touser.value+=fonta.value">
            <OPTION selected value="">�ռ���</OPTION>
<%
call listmyfriend()
sub listmyfriend()
dim rs
set rs=conn.execute("select me,fd from fd where me='"&loginuser&"'")
do while not rs.eof
response.Write("<OPTION value="""&rs("fd")&","">"&rs("fd")&"</OPTION>")
rs.movenext
loop
end sub
%>
          </SELECT> </td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td valign=top class=tablebody1>���⣺</td>
        <td valign=center class=tablebody1><font color="#0000FF">
          <input name="name" type="text" size="50" maxlength="50" >
          </font></td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td valign=top class=tablebody1>���ݣ�</td>
        <td valign=center class=tablebody1> <textarea rows="8" name="ly" cols="50" ></textarea> 
        </td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td colspan=2 class=tablebody1> ˵����<br>
          �� ������ʹ��Ctrl+Enter����ݷ��Ͷ���<br>
          �� ������Ӣ��״̬�µĶ��Ž��û�������ʵ��Ⱥ�������5���û�<br>
          �� �������50���ַ����������4000���ַ�</td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td colspan=2 align=middle valign=center class=tablebody2> <div align="center"> 
            <font color="#0000FF"> 
            <input name=FORM1 type=submit  value="�� ��">
            &nbsp; 
            <input type="reset" name="Reset" value="�� ��">
            &nbsp;&nbsp; 
            <input type="button" name="Submit2" value="�� ��"onClick="javascript:window.close()">
             &nbsp;&nbsp; 
            <input type="button" name="Submit" value="�����¼">
            </font></div></td>
      </tr>
  </table>
</form>
</body>
</html>
