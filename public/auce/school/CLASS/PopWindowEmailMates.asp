<!--#include file=globals.asp -->
<%
UserID=Session("UserID")
'if ClassID="" then ClassID="GoodLuck"
if UserID="" then response.redirect "../Error.asp?Info=�����ȵ�½��лл��"

%> 
<HTML>
<HEAD>
<title>֪ͨУ��ע��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="ͬѧ��ͬѧ¼��У�ѡ���ʦ��ѧУ���༶������">
<meta name="web_designner" content="�����">
<STYLE type=text/css>
</STYLE>
<LINK href="../scss.css" rel=stylesheet>
</HEAD>
<BODY BGCOLOR=#FFFFFF leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#52707D"><img src="../school_images/poplogo.gif" width="108" height="34"></td>
  </tr>
  <tr>
    <td bgcolor="#52707D" background="../school_images/popwindow_03.gif">&nbsp;</td>
  </tr>
</table>
<table width="550" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center" valign="top"> <br>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td class="topic">֪ͨУ��ע��</td>
        </tr>
      </table>
<%
    '���� curUserID �� UserInfo �鵽eMail��ַ
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open DBParams
    SQL = "select U1.RealName,U2.Email from UserInfo as U1,UserCommunicationInfo as U2 where U1.UserID=U2.UserID and U1.UserID='"&UserID&"'"
    'response.write sql
    set rs=conn.Execute(sql)
    if not rs.eof then
       RealName = rs("RealName")
       Email = rs("Email")		
    end if  
    rs.close
    set rs=nothing
   if not  request("Submit")="" then
    
    Title=Request("Title")
    MailCopyRight = "��ӭ������������ http://www.88com.cn "
        
    Comment = Comment & "��Ŀ:" & Title & CRLF
    Comment = Comment & "����:" & Request("Writer") & CRLF
    Comment = Comment & "����:" & Request("Comment") & CRLF
    Comment = Comment & CRLF & CRLF & CRLF & MailCopyRight
   
    
    for i=1 to 5
       'Ҫ���û� eMail������Ч�Լ��
       Mailto=Trim(request("Email"&i))
       if not MailTo=""  then
          if instr(Mailto,"@")=0 or instr(MailTo,".")=0 or len(MailTo)<6 then
             Info=Info&MailTo&"�������󣬲��ܷ���<br>"
          else
             Set MailOb = Server.CreateObject("CDONTS.NewMail") 
             MailOb.Send Email,MailTo,Title,Comment
             Set MailOb = nothing
             Info=Info&MailTo&"�����ѷ���<br>"
          end if
       end if
    next
    
    
    response.redirect "../success.asp?Info="&Info&"���ͽ�����лл���٣�"
else  
%>
      <table width="500" border="0" cellspacing="1" cellpadding="1">
        <form name="form1" method="post" action="">
          <tr> 
            <td width="130" align="right" class="topic">У�ѵ����ʼ�1��<br>
            </td>
            <td height="15" width="370"> 
              <input type="text" name="Email1" size="30">
            </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic">У�ѵ����ʼ�2��</td>
            <td width="370"> 
              <input type="text" name="Email2" size="30">
            </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic">У�ѵ����ʼ�3��</td>
            <td width="370"> 
              <input type="text" name="Email3" size="30">
            </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic">У�ѵ����ʼ�4��</td>
            <td width="370"> 
              <input type="text" name="Email4" size="30">
            </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic">У�ѵ����ʼ�5��</td>
            <td width="370"> 
              <input type="text" name="Email5" size="30">
            </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic">&nbsp;</td>
            <td width="370">&nbsp; </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic">�� &nbsp;Ŀ��</td>
            <td width="370"> 
              <input type="text" name="Title" size="40" value="��ã�������">
            </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic">�����ˣ�</td>
            <td width="370"> 
              <input type="text" name="Writer" value="<%=RealName%>">
            </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic">�� &nbsp;�ݣ�</td>
            <td width="370"> 
              <textarea name="Comment" cols="40" rows="5">��������ɣ�
269��ͬѧ¼�� http://school.269.net ��</textarea>
            </td>
          </tr>
          <tr> 
            <td width="130" align="right">&nbsp;</td>
            <td width="370" height="40" valign="bottom"> 
              <input type="submit" name="Submit" value="����">
              <input type="reset" name="Submit2" value="ȫ����д">
            </td>
          </tr>
        </form>
      </table>
<%end if%>
      <br>
    </td>
  </tr>
</table>
</BODY>
</HTML>