<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%userid=session("userid")
provinceid=trim(request.form("provinceid"))
%>
<html>
<head>
<title>����У��¼-���Ұ༶</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="ͬѧ��ͬѧ¼��У�ѡ���ʦ��ѧУ���༶������">
<STYLE type=text/css>
</STYLE>
<LINK href="scss.css" rel=stylesheet>
</head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="5"><CENTER>
<!--#include file=top2.htm-->
<table width="750" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td align="center" valign="top"> <br>
      <br>
      <table width="600" border="0" cellspacing="1" cellpadding="1" bgcolor="#111111">
        <tr> 
          <td bgcolor="#F5F5F5" align="center" class="topic"> <br>
            <table width="400" border="0" cellspacing="0" cellpadding="0">
              <form name="form1" method="post" action="find_class3.asp">
                <tr> 
                  <td height="32" valign="middle" width="20"><img src="school_images/topic_inco1.gif" width="15" height="15"> 
                  </td>
                  <td height="32" valign="middle" width="380">�ڶ�����ѡ��ѧУ���ڵ����� 
                    <select name="areaid">
                      <%set rs = createobject("ADODB.recordset")
                        sql="select * from areainfo where provinceid='"&provinceid&"'"  '" order by seed "
                        rs.open SQL,schooldb
                        while not rs.eof
                      %>
                      <option value="<%=rs("areaid")%>"><%=rs("areaname")%></option>
                      <%rs.movenext
                        wend
                        rs.close
                      %>
                    </select>
                  </td>
                </tr>
                <tr> 
                  <td align="left" height="32" valign="middle"><img src="school_images/topic_inco1.gif" width="15" height="15"> 
                  </td>
                  <td align="left" height="32" valign="middle">��������ѡ��ѧУ���ͣ� 
                    <select name="schooltypeid">
                      <option selected value=1>��ѧ</option>
                      <option value=2>��ר</option>
                      <option value=3>��ѧ</option>
                      <option value=4>��ר</option>
                      <option value=5>��У</option>
                    
                      <option value=6>Сѧ</option>
                    </select>
                  </td>
                </tr>
                <tr> 
                  <td align="left" height="32" valign="middle"><img src="school_images/topic_inco1.gif" width="15" height="15"></td>
                  <td align="left" height="32" valign="middle">���Ĳ��� ѧУ���ƹؼ��֣� 
                    <input type="text" name="keyword" size="12" maxlength="6">
                    �����</td>
                </tr>
                <tr> 
                  <td align="center" height="40" valign="bottom" colspan="2"> 
                    <input type="Hidden" name="provinceid" value="<%=provinceid%>">
                    <input name="opration" type="submit" value=" ��һ�� ">
                  </td>
                </tr>
              </form>
            </table>
            <br>
            <br>
            <table width="400" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td class="topic"><font color="#993300">ע�⣺ <br>
                  1.Ϊ�˱���һ��ѧУ������Ƶ��������дѧУ���ƹؼ��ֵ�ʱ����Ҫ���ҡ������еڰ�ʮ����ѧ���� �����롰��ʮ�塱����Ҫ��д��85���� 
                  <br>
                  2.�����ı�������Ӣ�ģ� </font></td>
              </tr>
            </table>
            <br>
            <br>
          </td>
        </tr>
      </table>
      <br>
      <br>
      <br>
    </td>
  </tr>
</table><CENTER>
<p></p>
<!--#include file=end.htm-->
<br>
      </body>
</html>