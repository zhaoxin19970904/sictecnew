<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%userid=session("userid")
provinceid=trim(request.form("provinceid"))
areaid=trim(request("areaid"))
schooltypeid=int(trim(request("schooltypeid")))
keyword=trim(request("keyword"))
if keyword="" then
  response.redirect "error.asp?info=������ѧУ���Ĺؼ��֣�"
end if
%>
<html>
<head>
<title>����У��¼-����༶</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="ͬѧ��ͬѧ¼��У�ѡ���ʦ��ѧУ���༶������">
<STYLE type=text/css>
</STYLE>
<LINK href="scss.css" rel=stylesheet>
</head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="5"><CENTER>
<!--#include file=top2.htm-->

<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td align="center" valign="top" bgcolor="#F6F6F6" width="760"> <br>
      <br>


      <%
            if Request("Page")<>"" then
               Page = CLng(Request("Page"))
            end if
            If Page < 1 Then
                Page = 1
            end if
       %>
      <table width="600" border="0" cellspacing="0" cellpadding="0">
        <tr align="left" valign="top">
          <td height="26" width="20" valign="middle"><img src="school_images/topic_inco1.gif" width="15" height="15"></td>
          <%
                    PageSize=5
                    set rs=Server.CreateObject("ADODB.RecordSet")
                    sql="select count(*) as a from schoolinfo where schooltypeid="&schooltypeid&" and schoolid like '"&areaid&"%' and schoolname like '%"&keyword&"%'"
                    rs.open SQL,schooldb,1,3
                    count=rs("a")
                    PageCount=CInt(rs("a")/PageSize+0.5)
                rs.Close
      %>
          <td height="20" valign="middle">�������Ĺؼ��������� <%=count%> ��ѧУ����ѡ����룺</td>
        </tr>
      </table>
      <table width="600" border="0" cellspacing="1" cellpadding="1">
        <form name="form1" method="post" action="find_class4.asp">
          <%
            if page=1 then
           SQL="select top 5 * from schoolinfo where schooltypeid="&schooltypeid&" and schoolid like '"&areaid&"%' and schoolname like '%"&keyword&"%' order by schoolid  "
         else
            SQL="select top 5 * from schoolinfo where schooltypeid="&schooltypeid&" and schoolid like '"&areaid&"%' and schoolname like '%"&keyword&"%' and schoolid not in (Select top "&Cstr(PageSize*(page-1))&" schoolid from schoolinfo where schooltypeid="&schooltypeid&" and schoolid like '"&areaid&"%' and schoolname like '%"&keyword&"%' order by schoolid  ) order by schoolid"
            end if

            rs.open SQL,schooldb
            while not rs.eof
            schoolname=replace(rs("schoolname"),keyword,"<font color=red>"&keyword&"</font>")
          %>
          <tr>
            <td height="28" bgcolor="#EBEBEB">
              <input type="radio" name="schoolid" value="<%=rs("schoolid")%>">
              <%=schoolname%></td>
          </tr>
          <%
            rs.movenext
            wend
            rs.close
          %>
          <tr>
            <td height="28" bgcolor="#CCCCCC" align="center" class="topic">��ǰҳ��<%=page%>/<%=pagecount%>ҳ
              | <a href="find_class3.asp?page=1&&areaid=<%=areaid%>&schooltypeid=<%=schooltypeid%>&keyword=<%=keyword%>">��һҳ</a>
              | <a href="find_class3.asp?page=<%=page-1%>&areaid=<%=areaid%>&schooltypeid=<%=schooltypeid%>&keyword=<%=keyword%>">��һҳ</a>
              | <a href="find_class3.asp?page=<%=page+1%>&areaid=<%=areaid%>&schooltypeid=<%=schooltypeid%>&keyword=<%=keyword%>">��һҳ</a>
              | <a href="find_class3.asp?page=<%=pagecount%>&areaid=<%=areaid%>&schooltypeid=<%=schooltypeid%>&keyword=<%=keyword%>">ĩҳ</a>
              | </td>
          </tr>
          <% if count=0 then
               response.write "</center><font color=red>�Բ���û����Ҫ�ҵ�ѧУ��</font></cener>"
             end if
          %>
          <tr>
            <td height="30" bgcolor="#F6F6F6">&nbsp;&nbsp;��Ҫ����İ༶����һ���ģ�������ѧ���Ϊ׼����
              <select name=enterdate>
                <option value="0" selected>��ѡ��</option>
                <option value=1970>1970��</option>
                <option value=1971>1971��</option>
                <option value=1972>1972��</option>
                <option value=1973>1973��</option>
                <option value=1974>1974��</option>
                <option value=1975>1975��</option>
                <option value=1976>1976��</option>
                <option value=1977>1977��</option>
                <option value=1978>1978��</option>
                <option value=1979>1979��</option>
                <option value=1980>1980��</option>
                <option value=1981>1981��</option>
                <option value=1982>1982��</option>
                <option value=1983>1983��</option>
                <option value=1984>1984��</option>
                <option value=1985>1985��</option>
                <option value=1986>1986��</option>
                <option value=1987>1987��</option>
                <option value=1988>1988��</option>
                <option value=1989>1989��</option>
                <option value=1990>1990��</option>
                <option value=1991>1991��</option>
                <option value=1992>1992��</option>
                <option value=1993>1993��</option>
                <option value=1994>1994��</option>
                <option value=1995>1995��</option>
                <option value=1996>1996��</option>
                <option value=1997>1997��</option>
                <option value=1998>1998��</option>
                <option value=1999>1999��</option>
                <option value=2000>2000��</option>
                <option value=2001>2001��</option>
                <option value=2000>2002��</option>
                <option value=2001>2003��</option>
              </select>
              <input type="submit" name="Submit" value="����">
            </td>
          </tr>
          <tr>
            <td height="28" bgcolor="#FFFFFF" align="center"> <br>
              <table width="560" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td class="topic"><font color="#993300">ע�⣺<br>
                    ��������㷢�����ѧУ����ע������ѧУ�ظ�ע�ᣬ�뷢��֪ͨ����Ա��Ϊ�����ظ�ע�ᣬ��ͬУУ����ɵĲ���Ҫ�鷳����ע����ڵ�ѧУ�����У��������Ը�֪����Ա�޸ġ�
                    </font></td>
                </tr>
              </table>
              <br>
            </td>
          </tr>
        </form>
      </table>
      <br>
      <table width="600" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="20"><img src="school_images/topic_inco1.gif" width="15" height="15"></td>
          <td width="480" valign="middle" class="topic" height="30"> <font color="#CC0000">����ϱ���û����Ҫ�����ѧУ������д�������Ϣ������</font></td>
        </tr>
      </table>
      <table width="600" border="0" cellspacing="1" cellpadding="1">
        <form name="form2" method="post" action="register_school.asp">
          <tr>
            <td height="28">ѧУ���ƣ�
              <input type="text" name="schoolname" size="40">
            </td>
          </tr>
          <tr>
            <td height="28">ѧУ��ַ�� http://
              <input type="text" name="schoolweb" size="40" value="��">
            </td>
          </tr>
          <tr>   <input type="Hidden" name="provinceid" value="<%=provinceid%>">
            <td height="28">ѧУ��ַ��
              <input type="text" name="schooladdr" size="40" value="��֪��">
              &nbsp;&nbsp;&nbsp; �ʱࣺ
              <input type="text" name="schoolzip" size="6" value="��">
            </td>
          </tr>
          <tr>
            <td height="28">��Ҫ����İ༶����һ���ģ�������ѧ���Ϊ׼����
              <select name=enterdate>
                <option value="0" selected>��ѡ��</option>
                <option value=1970>1970��</option>
                <option value=1971>1971��</option>
                <option value=1972>1972��</option>
                <option value=1973>1973��</option>
                <option value=1974>1974��</option>
                <option value=1975>1975��</option>
                <option value=1976>1976��</option>
                <option value=1977>1977��</option>
                <option value=1978>1978��</option>
                <option value=1979>1979��</option>
                <option value=1980>1980��</option>
                <option value=1981>1981��</option>
                <option value=1982>1982��</option>
                <option value=1983>1983��</option>
                <option value=1984>1984��</option>
                <option value=1985>1985��</option>
                <option value=1986>1986��</option>
                <option value=1987>1987��</option>
                <option value=1988>1988��</option>
                <option value=1989>1989��</option>
                <option value=1990>1990��</option>
                <option value=1991>1991��</option>
                <option value=1992>1992��</option>
                <option value=1993>1993��</option>
                <option value=1994>1994��</option>
                <option value=1995>1995��</option>
                <option value=1996>1996��</option>
                <option value=1997>1997��</option>
                <option value=1998>1998��</option>
                <option value=1999>1999��</option>
                <option value=2000>2000��</option>
                <option value=2001>2001��</option>
                <option value=2002>2002��</option>
                <option value=2003>2003��</option>
              </select>
            </td>
          </tr>
          <tr>
            <td height="28">�༶ȫ�ƣ�
              <input type="text" name="classname" size="40">
              ���༶ȫ�Ʊ�����'��'�ֽ�β��</td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF" align="center"> <br>
              <table width="560" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td class="topic"><font color="#993300">ע�⣺ <br>
                    1.У����дȫ�ƣ��Ա���һ��ѧУ������Ƶ���������á������еڰ�ʮ����ѧ����<br>
                    �������á�������ʮ���С��� <br>
                    2.�����ı�������Ӣ�ģ� <br>
                    3.�����������ı�ʾ�����á������еڰ�ʮ����ѧ���������á������е�85��ѧ���� </font></td>
                </tr>
              </table>
              <br>
            </td>
          </tr>
          <tr>
            <td height="30" align="center">
              <input type="hidden" name="schooltypeid" value="<%=schooltypeid%>">
              <input type="hidden" name="areaid" value="<%=areaid%>">
              <input type="submit" name="Submit3" value="����">
              <input type="submit" name="Submit32" value="��д">
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
    <CENTER>
<p></p>
<!--#include file=end.htm-->  </body>
</html>