<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%userid=session("userid")
schoolid=trim(request("schoolid"))
enterdate=trim(request("enterdate"))
if schoolid="" then
  response.redirect "error.asp?info=��ѡ��ѧУ��"
end if
if enterdate="0" then
  response.redirect "error.asp?info=��ѡ����У��ݣ�"
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
<div align="center">
  <center>
<table width="760" border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
  <tr> 
    <td align="center" valign="top" bgcolor="#F6F6F6"> <br>
      <br>
      <br>
      <table width="600" border="0" cellspacing="0" cellpadding="0">
        <tr align="left" valign="top"> 
          <td height="26" width="20" valign="middle"><img src="school_images/topic_inco1.gif" width="15" height="15"></td>
          <td height="20" valign="middle">��ѡ���������İ༶���룺</td>
        </tr>
      </table>
      <%
            if Request("Page")<>"" then 
       		Page = CLng(Request("Page"))
    	    end if 
            If Page < 1 Then 
                Page = 1
            end if
       %>
      <table width="600" border="0" cellspacing="1" cellpadding="1">
        <form name="form1" method="post" action="find_class5.asp">
          <%
                    PageSize=5
                    set rs=Server.CreateObject("ADODB.RecordSet")
                    sql="select count(*) as a from classinfo where classid like '"&schoolid&"%' and enterdate='"&enterdate&"'"  
                    rs.open SQL,schooldb,1,3
                    count=rs("a")
                    PageCount=CInt(rs("a")/PageSize+0.5)
	            rs.Close
	            if page=1 then
	       SQL="select top 5 * from classinfo where classid like '"&schoolid&"%' and enterdate='"&enterdate&"' order by classid  "
 	    else
 	       SQL="select top 5 * from classinfo where classid like '"&schoolid&"%' and enterdate='"&enterdate&"' and classid not in (Select top "&Cstr(PageSize*(page-1))&" classid from classinfo where classid like '"&schoolid&"%' and enterdate='"&enterdate&"' order by classid ) order by classid" 
            end if
            
            rs.open SQL,schooldb
            while not rs.eof 
            
          %>
          <tr> 
            <td height="28" bgcolor="#EBEBEB"> 
              <input type="radio" name="classid" value="<%=rs("classid")%>">
              <%=rs("classname")%></td>
          </tr>
          <%
            rs.movenext
            wend
            rs.close
          %>
          <tr> 
            <td height="28" bgcolor="#CCCCCC" align="center" class="topic">��ǰҳ��<%=page%>/<%=pagecount%>ҳ 
              | <a href="find_class4.asp?page=1&schoolid=<%=schoolid%>&enterdate=<%=enterdate%>">��һҳ</a> 
              | <a href="find_class4.asp?page=<%=page-1%>&schoolid=<%=schoolid%>&enterdate=<%=enterdate%>">��һҳ</a> 
              | <a href="find_class4.asp?page=<%=page+1%>&schoolid=<%=schoolid%>&enterdate=<%=enterdate%>">��һҳ</a> 
              | <a href="find_class4.asp?page=<%=pagecount%>&schoolid=<%=schoolid%>&enterdate=<%=enterdate%>">ĩҳ</a> 
              | </td>
          </tr>
          <% if count=0 then
               response.write "</center><font color=red>�Բ���û����Ҫ�ҵİ༶��</font></cener>"
             end if
          %>
          <tr> 
            <td height="30" bgcolor="#EBEBEB" align="center">&nbsp;&nbsp;��ѡ����ݼ��룺 
              <select name="userstatus">
                <option selected value="��Ա">��Ա</option>
                <option value="��ʦ">��ʦ</option>
                <option value="����">����</option>
              </select>
              <input type="submit" name="Submit" value="����">
            </td>
          </tr>
        </form>
      </table>
      <br>
      <table width="600" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="20"><img src="school_images/topic_inco1.gif" width="15" height="15"></td>
          <td width="480" valign="middle" class="topic" height="30"> <font color="#CC0000">����ϱ���û����Ҫ����İ༶������д�������Ϣ������</font></td>
        </tr>
      </table>
      <table width="600" border="0" cellspacing="1" cellpadding="1">
        <form name="form2" method="post" action="register_class.asp">
          <tr> 
            <td height="28">�༶ȫ�ƣ� 
              <input type="text" name="classname" size="40">
              ���༶ȫ�Ʊ�����'��'�ֽ�β��</td>
          </tr>
          <tr> 
            <td height="30" align="center"> 
              <input type="hidden" name="enterdate" value="<%=enterdate%>">
              <input type="hidden" name="schoolid" value="<%=schoolid%>">
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
  </center>
</div>
<br>
<br>
     <CENTER>
<p></p>
<!--#include file=end.htm-->    </body>
</html>