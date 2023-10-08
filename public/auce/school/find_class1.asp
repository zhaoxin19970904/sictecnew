<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%userid=session("userid")%>
<html>
<head>
<title>商务校友录-加入新班级</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="同学、同学录、校友、老师、学校、班级、交流">
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
          <td bgcolor="#F5F5F5" align="center" class="topic"><br>
            <br>
            <table width="400" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td align="center" class="topic">欢迎您！<%=userid%>:<br>
                  请按照下面的步骤查找和创建您的班级！</td>
              </tr>
            </table>
            <br>
            <hr width="400" size="1" noshade align="center">
            <br>
            <table width="400" border="0" cellspacing="0" cellpadding="0">
              <form name="form1" method="post" action="find_class2.asp">
                <tr> 
                  <td height="30" width="20"><img src="school_images/topic_inco1.gif" width="15" height="15"> 
                  </td>
                  <td height="30" width="380">第一步：查找到你的学校</td>
                </tr>
                <tr> 
                  <td align="center" height="30" colspan="2">请选择学校所在的省市： 
                    <select name="Provinceid" size="1" style="font-size: 9pt">
                      <%set rs = createobject("ADODB.recordset")
                        sql="select * from province order by seed "
                        rs.open SQL,schooldb
                        while not rs.eof
                           temp=""
                        if rs("provincename")="陕西省" then
                           temp="selected"
                        end if   
                      %>
                      <option <%=temp%> value="<%=rs("provinceid")%>"><%=rs("provincename")%></option>
                      <%rs.movenext
                        wend
                        rs.close
                      %>
                    </select>
                  </td>
                </tr>
                <tr> 
                  <td align="center" height="40" valign="bottom" colspan="2"> 
                    <input type="Hidden" name="nId" value="1857069">
                    <input name="opration" type="submit" value=" 下一步 ">
                  </td>
                </tr>
              </form>
            </table>
            <br>
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
</table>
<br>
<br><CENTER>
<p></p>
<!--#include file=end.htm-->
      </body>
</html>