<!-- #Include File=check.asp -->
<%
user=request("user")
if request("form1")<>"" then
set rs = Server.CreateObject("ADODB.Recordset")
sqltext = "select mess From user where user='"&request.form("user")&"' "
rs.open sqltext,conn,1,1
if rs.recordcount=0 then
response.write "�����˲�����!<br>"
response.write"<input type=""button"" name=""Submit"" value=""�رմ���"" onClick=""javascript:window.close()"">"
response.end
end if
if rs("mess")="1" then
response.write "sorry!�����˾ܾ����ն���Ϣ!<br>"
response.write"<input type=""button"" name=""Submit"" value=""�رմ���"" onClick=""javascript:window.close()"">"
response.end
end if
set rs = Server.CreateObject("ADODB.Recordset")
sqltext = "select *From mess "
rs.open sqltext,conn,3,3
from_address=address(Request.ServerVariables("REMOTE_ADDR"))
memo=request.form("ly")
memo=ubbcode(memo,vbCrlf,"<br>")
rs.addnew

rs("me")=request.cookies("renwen")("user")
rs("fd")=request.form("user")
rs("ly")=memo
rs("name")=request.form("name")
rs("date") = date
rs("time") = time
rs("ip")=from_address
rs.Update
Rs.Close
Conn.Close
Set Rs=Nothing
Set Conn=Nothing
response.write "<script language=JavaScript>" & chr(13) & "alert('���ŷ��ͳɹ���лл��ʹ�ñ�����');" & "window.close();" & "</script>"
end if
%>
<html>
<head>
<title>��<%=fd %>���ѷ�����Ϣ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<SCRIPT language=javascript id=clientEventHandlersJS>
<!--

function form1_onsubmit() 
{
  if(document.form1.name.value.length<1)
 {
   alert("������д���ű�����");
   document.form1.name.focus();
   return false;
 }
   if(document.form1.ly.value.length<1)
 {
   alert("����д����Ҫ�������ѵĶ���");
   document.form1.ly.focus();
   return false;
  }

}

//-->
</SCRIPT>
<link rel="stylesheet" href="../images/style.css" type="text/css">
<link href="DEFAULT.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#E0E4FE" background="backimg/03.gif">
<FORM language=javascript name=form1   onsubmit="return form1_onsubmit()"
      action=sendmess.asp method=post>
  <div align="center"></div>
  <table width="418" align=center cellpadding=3 cellspacing=1 bgcolor="#003366" class=tableborder1>
    <tbody>
      <tr align="center" bgcolor="#FFFFFF"> 
        <td height="27" colspan=3 background="backimg/bg1.gif"> <font color="#FFFFFF"> 
          ��</font><font size="3" color="#FFFFFF"><b><font size="4"><font size="4"><%=user%></font></font></b></font><font color="#FFFFFF">���Ͷ���Ϣ��������������Ϣ�� </font></td>
      <tr> </tr>
      <tr bgcolor="#FFFFFF"> 
        <td valign=top class=tablebody1><font color="#000000"><strong>������</strong></font></td>
        <td valign=center class=tablebody1><font color="#FFFFFF">
          <input  name="user" value="<%=request("user")%>" size="50">
          </font><font color="#0000FF">&nbsp; </font> </td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td valign=top class=tablebody1><b>���⣺</b></td>
        <td valign=center class=tablebody1><font color="#0000FF">
          <input type="text" name="name" size="50" >
          </font></td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td valign=top class=tablebody1><b>���ݣ�</b></td>
        <td valign=center class=tablebody1><font color="#0000FF"> 
          <textarea rows="8" name="ly" cols="50" ></textarea>
          </font> </td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td colspan=2 class=tablebody1> �������<b>50</b>���ַ����������<b>300</b>���ַ� </td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td colspan=2 align=middle valign=center class=tablebody2> <div align="center"> 
            <input type="submit" name="form1" value="�ͳ�">
            &nbsp; 
            <input type="reset" name="Submit3" value="��д">
            &nbsp; 
            <input type="button" name="Submit" value="�رմ���"onClick="javascript:window.close()">
          </div></td>
      </tr>
    </tbody>
  </table>
</FORM>


  