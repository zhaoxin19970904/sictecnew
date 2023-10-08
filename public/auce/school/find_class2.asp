<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%userid=session("userid")
provinceid=trim(request.form("provinceid"))
%>
<html>
<head>
<title>商务校友录-查找班级</title>
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
          <td bgcolor="#F5F5F5" align="center" class="topic"> <br>
            <table width="400" border="0" cellspacing="0" cellpadding="0">
              <form name="form1" method="post" action="find_class3.asp">
                <tr> 
                  <td height="32" valign="middle" width="20"><img src="school_images/topic_inco1.gif" width="15" height="15"> 
                  </td>
                  <td height="32" valign="middle" width="380">第二步：选择学校所在地区： 
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
                  <td align="left" height="32" valign="middle">第三步：选择学校类型： 
                    <select name="schooltypeid">
                      <option selected value=1>大学</option>
                      <option value=2>大专</option>
                      <option value=3>中学</option>
                      <option value=4>中专</option>
                      <option value=5>技校</option>
                    
                      <option value=6>小学</option>
                    </select>
                  </td>
                </tr>
                <tr> 
                  <td align="left" height="32" valign="middle"><img src="school_images/topic_inco1.gif" width="15" height="15"></td>
                  <td align="left" height="32" valign="middle">第四步： 学校名称关键字： 
                    <input type="text" name="keyword" size="12" maxlength="6">
                    （必填）</td>
                </tr>
                <tr> 
                  <td align="center" height="40" valign="bottom" colspan="2"> 
                    <input type="Hidden" name="provinceid" value="<%=provinceid%>">
                    <input name="opration" type="submit" value=" 下一步 ">
                  </td>
                </tr>
              </form>
            </table>
            <br>
            <br>
            <table width="400" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td class="topic"><font color="#993300">注意： <br>
                  1.为了避免一个学校多个名称的情况，填写学校名称关键字的时候，如要查找“西安市第八十五中学”， 请填入“八十五”、不要填写“85”； 
                  <br>
                  2.用中文表达，请勿用英文； </font></td>
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