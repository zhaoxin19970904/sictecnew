<!--#include file=globals.asp -->
<%
UserID=Session("UserID")
'if ClassID="" then ClassID="GoodLuck"
if UserID="" then response.redirect "../Error.asp?Info=请您先登陆，谢谢！"

%> 
<HTML>
<HEAD>
<title>通知校友注册</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="同学、同学录、校友、老师、学校、班级、交流">
<meta name="web_designner" content="孙梦昕">
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
          <td class="topic">通知校友注册</td>
        </tr>
      </table>
<%
    '根据 curUserID 在 UserInfo 查到eMail地址
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
    MailCopyRight = "欢迎访问商务世纪 http://www.88com.cn "
        
    Comment = Comment & "题目:" & Title & CRLF
    Comment = Comment & "作者:" & Request("Writer") & CRLF
    Comment = Comment & "内容:" & Request("Comment") & CRLF
    Comment = Comment & CRLF & CRLF & CRLF & MailCopyRight
   
    
    for i=1 to 5
       '要对用户 eMail进行有效性检测
       Mailto=Trim(request("Email"&i))
       if not MailTo=""  then
          if instr(Mailto,"@")=0 or instr(MailTo,".")=0 or len(MailTo)<6 then
             Info=Info&MailTo&"信箱有误，不能发送<br>"
          else
             Set MailOb = Server.CreateObject("CDONTS.NewMail") 
             MailOb.Send Email,MailTo,Title,Comment
             Set MailOb = nothing
             Info=Info&MailTo&"信箱已发送<br>"
          end if
       end if
    next
    
    
    response.redirect "../success.asp?Info="&Info&"发送结束，谢谢光临！"
else  
%>
      <table width="500" border="0" cellspacing="1" cellpadding="1">
        <form name="form1" method="post" action="">
          <tr> 
            <td width="130" align="right" class="topic">校友电子邮件1：<br>
            </td>
            <td height="15" width="370"> 
              <input type="text" name="Email1" size="30">
            </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic">校友电子邮件2：</td>
            <td width="370"> 
              <input type="text" name="Email2" size="30">
            </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic">校友电子邮件3：</td>
            <td width="370"> 
              <input type="text" name="Email3" size="30">
            </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic">校友电子邮件4：</td>
            <td width="370"> 
              <input type="text" name="Email4" size="30">
            </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic">校友电子邮件5：</td>
            <td width="370"> 
              <input type="text" name="Email5" size="30">
            </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic">&nbsp;</td>
            <td width="370">&nbsp; </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic">题 &nbsp;目：</td>
            <td width="370"> 
              <input type="text" name="Title" size="40" value="你好，老朋友">
            </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic">发送人：</td>
            <td width="370"> 
              <input type="text" name="Writer" value="<%=RealName%>">
            </td>
          </tr>
          <tr> 
            <td width="130" align="right" class="topic">内 &nbsp;容：</td>
            <td width="370"> 
              <textarea name="Comment" cols="40" rows="5">快来这儿吧！
269的同学录（ http://school.269.net ）</textarea>
            </td>
          </tr>
          <tr> 
            <td width="130" align="right">&nbsp;</td>
            <td width="370" height="40" valign="bottom"> 
              <input type="submit" name="Submit" value="发送">
              <input type="reset" name="Submit2" value="全部重写">
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