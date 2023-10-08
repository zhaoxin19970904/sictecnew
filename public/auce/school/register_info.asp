<%userid=trim(session("userid"))%>
<html>
<head>
<title>商务校友录-注册新用户</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="同学、同学录、校友、老师、学校、班级、交流">
<STYLE type=text/css>
</STYLE>
<LINK href="scss.css" rel=stylesheet>
</head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="5"><CENTER>
<!--#include file=top2.htm-->
<table border="0" width="750" cellspacing="0" cellpadding="0" align="center">
  <tr> 
      </tr>
</table>
<table width="750" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td align="center" valign="top" bgcolor="#FFFFFF"> <br>
      <table width="90%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td height="40" valign="bottom">新用户注册&gt;&gt;&gt;第二步：</td>
        </tr>
      </table>
      <br>
      <table width="90%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td class="topic">欢迎您！<%=userid%><br>
            &nbsp;&nbsp;&nbsp;&nbsp;为了方便您的同学联系您，请认真填写下面的个人信息，带（！！）的为必填信息。</td>
        </tr>
      </table>
      <br>
      <table width="500" border="0" cellspacing="2" cellpadding="2">
        <form name="form1" method="post" action="register_step2.asp">
          <tr> 
            <td width="130" align="right">真实姓名：<br>
            </td>
            <td height="15" width="370"> 
              <input type="text" name="realname" size="20">
              ！！（请填写中文）</td>
          </tr>
          <tr> 
            <td width="130" align="right">昵称（笔名）：</td>
            <td width="370"> 
              <input type="text" name="dearname" value="无" size="20">
            </td>
          </tr>
          <tr> 
            <td width="130" align="right">个人密码：</td>
            <td width="370"> 
              <input type="password" name="passwd" size="20">
              ！！（多于5位）</td>
          </tr>
          <tr> 
            <td width="130" align="right">重复密码：</td>
            <td width="370"> 
              <input type="password" name="confirmpasswd" size="20">
              ！！</td>
          </tr>
          <tr> 
            <td width="130" align="right">您的生日：</td>
            <td width="370"> 
              <input type="text" name="BYear" size="4" value="19××">
              年 
              <input type="text" name="BMonth" size="2" value="00">
              月 
              <input type="text" name="BDay" size="2" value="00">
              日 （！！）</td>
          </tr>
          <tr> 
            <td width="130" align="right">密码丢失提示问题：</td>
            <td width="370"> 
              <input type="text" name="passask" size="20">
              ！！（例如您的身份证号码?）</td>
          </tr>
          <tr> 
            <td width="130" align="right">提示答案：</td>
            <td width="370"> 
              <input type="text" name="anwserpass" size="20">
              ！！</td>
          </tr>
          <tr> 
            <td width="130" align="right">　</td>
            <td width="370">　</td>
          </tr>
          <tr> 
            <td width="130" align="right">通信地址(邮编)：</td>
            <td width="370"> 
              <input type="text" name="communicationaddr" size="40" value="不告诉你">
              ！！ 
              <input type="checkbox" name="iscashow" value="0">
              保密</td>
          </tr>
          <tr> 
            <td width="130" align="right">固定电话：</td>
            <td width="370"> 
              <input type="text" name="telephone" size="20">
              ！！&nbsp; 
              <input type="checkbox" name="istelephoneshow" value="0">
              保密</td>
          </tr>
          <tr> 
            <td width="130" align="right">移动电话：</td>
            <td width="370"> 
              <input type="text" name="mobile" value="无" size="20">
              &nbsp; 
              <input type="checkbox" name="ismobileshow" value="0">
              保密 </td>
          </tr>
          <tr> 
            <td width="130" align="right">传呼：</td>
            <td width="370"> 
              <input type="text" name="BP" value="无" size="20">
              &nbsp; 
              <input type="checkbox" name="isbpshow" value="0">
              保密 </td>
          </tr>
          <tr> 
            <td width="130" align="right">OICQ/ICQ 号码：</td>
            <td width="370"> 
              <input type="text" name="QQ" value="无" size="20">
              &nbsp; 
              <input type="checkbox" name="isqqshow" value="0">
              保密 </td>
          </tr>
          <tr> 
            <td width="130" align="right">工作单位：</td>
            <td width="370"> 
              <input type="text" name="workshop" size="40" value="无">
              &nbsp; 
              <input type="checkbox" name="iswsshow" value="0">
              保密 </td>
          </tr>
          <tr> 
            <td width="130" align="right">家庭居住地址：</td>
            <td width="370"> 
              <input type="text" name="homeaddr" value="无" size="20">
              &nbsp; 
              <input type="checkbox" name="ishashow" value="0">
              保密</td>
          </tr>
          <tr> 
            <td width="130" align="right">电子邮件地址：</td>
            <td width="370"> 
              <input type="text" name="email" size="20">
              ！！ &nbsp; 
              <input type="checkbox" name="isemailshow" value="0">
              保密</td>
          </tr>
          <tr> 
            <td width="130" align="right">附加信息：</td>
            <td width="370"> 
              <textarea name="addinfo" rows="5" cols="35">无</textarea>
              &nbsp; 
              <input type="checkbox" name="isafshow" value="0">
              保密</td>
          </tr>
          <tr> 
            <td width="130" align="right">　</td>
            <td width="370" height="50" valign="bottom"> 
              <input type="submit" name="Submit" value="完成">
              <input type="reset" name="Submit2" value="全部重写">
            </td>
          </tr>
        </form>
      </table>
      <br>
      <br>
      <br>
      <br>
    </td>
  </tr>
  </table>
<br>
<br>
<CENTER>
<p></p>
<!--#include file=end.htm-->      </body>
</html>