<!--#include file="admin_conn.asp" -->
<%
dim lttitle
if request.Form("Submit")<>"" then
call addpagehtm()
end if
sub addpagehtm()
dim rs,sqltext
if request.Form("tophtm")="" or request.Form("copyc")="" or request.Form("ltname")="" or request.Form("offtitle")="" or request.Form("badwords")="" then
lttitle="<li><b>!�������ɹ�</b><li>ע���趨ÿ��Ϊ������.�粻Ҫ����ʾֻҪ�ӿո��HTML����!"
exit sub
end if
set rs=server.createobject("adodb.recordset")
sqltext="select * from admin_copyc where id="&request.Form("id")
rs.open sqltext,conn,1,3
rs("badwords")=request.Form("badwords")
rs("tophtm")=request.Form("tophtm")
rs("copyc")=request.Form("copyc")
rs("ltname")=request.Form("ltname")
rs("offtitle")=request.Form("offtitle")
rs.update
lttitle="<li><b>!�����ɹ�</b><li>��̳�����Ƕ����ݸ��³ɹ�"
end sub
dim rs,sqltext
set rs=conn.execute("select tophtm,copyc,ltname,id,offlt,offtitle,badwords from admin_copyc")
if not rs.eof then
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ҳ�����趨</title>
<link href="DEFAULT.css" rel="stylesheet" type="text/css">
</head>

<body>
      <form name="form1" method="post" action="<%=request.ServerVariables("SCRIPT_NAME")%>">
        
  <table width="573" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#92b9fb">
    <tr align="center" background="../backimg/bg1.gif"> 
      <td height="28" colspan="2"><font color="#FFFFFF"><strong>��ҳ�����趨</strong></font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="2"><%=lttitle%> </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="107">����̳״̬</td>
      <td> <%if rs("offlt")=0 then 
	  response.Write("<a href=admin_cz.asp?czlb=offlt&id="&rs("id")&" title=����ر���̳ onclick=""{if(confirm('��ȷʵҪ�ر���̳��ȷ����?')){return true;}return false;}""> <b>����</b></a>")
	  else
	  response.Write("<a href=admin_cz.asp?czlb=offlt&id="&rs("id")&" title=���������̳ onclick=""{if(confirm('��ȷʵҪ������̳��ȷ����?')){return true;}return false;}""><b>�ر�</b>---<font color=red>������̳���ڹر�״̬</font></a>")
	  end if
	  %>
        ---�������,�ر���̳ </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="107">��̳����</td>
      <td><input name="ltname" type="text" id="ltname3" value="<%=rs("ltname")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="107">��̳�����ַ�</td>
      <td><textarea name="badwords" cols="50" rows="6" id="badwords"><%=rs("badwords")%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="107">��̳�ر���ʾ<br>
        (֧��HTLM)</td>
      <td><textarea name="offtitle" cols="50" rows="6" id="offtitle"><%=rs("offtitle")%></textarea> 
        <br> </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td>��̳��������<br>
        (֧��HTLM) </td>
      <td width="451"><textarea name="tophtm" cols="50" rows="6"><%=rs("tophtm")%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td>��̳��Ȩ��Ϣ<br>
        (֧��HTLM) </td>
      <td> <textarea name="copyc" cols="50" rows="6"><%=rs("copyc")%></textarea> 
        <input name="id" type="hidden" id="id" value="<%=rs("id")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td>&nbsp;</td>
      <td align="center"> <input type="submit" name="Submit" value=" �� �� "> &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="reset" name="Submit2" value=" �� �� "> </td>
    </tr>
  </table>
        </form>
<%
end if
%>		
</body>
</html>
