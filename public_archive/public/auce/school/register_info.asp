<%userid=trim(session("userid"))%>
<html>
<head>
<title>����У��¼-ע�����û�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="ͬѧ��ͬѧ¼��У�ѡ���ʦ��ѧУ���༶������">
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
          <td height="40" valign="bottom">���û�ע��&gt;&gt;&gt;�ڶ�����</td>
        </tr>
      </table>
      <br>
      <table width="90%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td class="topic">��ӭ����<%=userid%><br>
            &nbsp;&nbsp;&nbsp;&nbsp;Ϊ�˷�������ͬѧ��ϵ������������д����ĸ�����Ϣ��������������Ϊ������Ϣ��</td>
        </tr>
      </table>
      <br>
      <table width="500" border="0" cellspacing="2" cellpadding="2">
        <form name="form1" method="post" action="register_step2.asp">
          <tr> 
            <td width="130" align="right">��ʵ������<br>
            </td>
            <td height="15" width="370"> 
              <input type="text" name="realname" size="20">
              ����������д���ģ�</td>
          </tr>
          <tr> 
            <td width="130" align="right">�ǳƣ���������</td>
            <td width="370"> 
              <input type="text" name="dearname" value="��" size="20">
            </td>
          </tr>
          <tr> 
            <td width="130" align="right">�������룺</td>
            <td width="370"> 
              <input type="password" name="passwd" size="20">
              ����������5λ��</td>
          </tr>
          <tr> 
            <td width="130" align="right">�ظ����룺</td>
            <td width="370"> 
              <input type="password" name="confirmpasswd" size="20">
              ����</td>
          </tr>
          <tr> 
            <td width="130" align="right">�������գ�</td>
            <td width="370"> 
              <input type="text" name="BYear" size="4" value="19����">
              �� 
              <input type="text" name="BMonth" size="2" value="00">
              �� 
              <input type="text" name="BDay" size="2" value="00">
              �� ��������</td>
          </tr>
          <tr> 
            <td width="130" align="right">���붪ʧ��ʾ���⣺</td>
            <td width="370"> 
              <input type="text" name="passask" size="20">
              �����������������֤����?��</td>
          </tr>
          <tr> 
            <td width="130" align="right">��ʾ�𰸣�</td>
            <td width="370"> 
              <input type="text" name="anwserpass" size="20">
              ����</td>
          </tr>
          <tr> 
            <td width="130" align="right">��</td>
            <td width="370">��</td>
          </tr>
          <tr> 
            <td width="130" align="right">ͨ�ŵ�ַ(�ʱ�)��</td>
            <td width="370"> 
              <input type="text" name="communicationaddr" size="40" value="��������">
              ���� 
              <input type="checkbox" name="iscashow" value="0">
              ����</td>
          </tr>
          <tr> 
            <td width="130" align="right">�̶��绰��</td>
            <td width="370"> 
              <input type="text" name="telephone" size="20">
              ����&nbsp; 
              <input type="checkbox" name="istelephoneshow" value="0">
              ����</td>
          </tr>
          <tr> 
            <td width="130" align="right">�ƶ��绰��</td>
            <td width="370"> 
              <input type="text" name="mobile" value="��" size="20">
              &nbsp; 
              <input type="checkbox" name="ismobileshow" value="0">
              ���� </td>
          </tr>
          <tr> 
            <td width="130" align="right">������</td>
            <td width="370"> 
              <input type="text" name="BP" value="��" size="20">
              &nbsp; 
              <input type="checkbox" name="isbpshow" value="0">
              ���� </td>
          </tr>
          <tr> 
            <td width="130" align="right">OICQ/ICQ ���룺</td>
            <td width="370"> 
              <input type="text" name="QQ" value="��" size="20">
              &nbsp; 
              <input type="checkbox" name="isqqshow" value="0">
              ���� </td>
          </tr>
          <tr> 
            <td width="130" align="right">������λ��</td>
            <td width="370"> 
              <input type="text" name="workshop" size="40" value="��">
              &nbsp; 
              <input type="checkbox" name="iswsshow" value="0">
              ���� </td>
          </tr>
          <tr> 
            <td width="130" align="right">��ͥ��ס��ַ��</td>
            <td width="370"> 
              <input type="text" name="homeaddr" value="��" size="20">
              &nbsp; 
              <input type="checkbox" name="ishashow" value="0">
              ����</td>
          </tr>
          <tr> 
            <td width="130" align="right">�����ʼ���ַ��</td>
            <td width="370"> 
              <input type="text" name="email" size="20">
              ���� &nbsp; 
              <input type="checkbox" name="isemailshow" value="0">
              ����</td>
          </tr>
          <tr> 
            <td width="130" align="right">������Ϣ��</td>
            <td width="370"> 
              <textarea name="addinfo" rows="5" cols="35">��</textarea>
              &nbsp; 
              <input type="checkbox" name="isafshow" value="0">
              ����</td>
          </tr>
          <tr> 
            <td width="130" align="right">��</td>
            <td width="370" height="50" valign="bottom"> 
              <input type="submit" name="Submit" value="���">
              <input type="reset" name="Submit2" value="ȫ����д">
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