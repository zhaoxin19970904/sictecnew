<!--#include file=globals.asp -->
<%
ClassID=Session("ClassID")
if ClassID="" then ClassID="GoodLuck"
id=Request("id")
if PicID="" then PicID="2"
set conn=server.CreateObject("ADODB.connection")
set   rs=server.createobject("adodb.recordset")
conn.ConnectionString="driver={Microsoft Access Driver (*.mdb)};DBQ=" &server.MapPath("pic/images.mdb") & ";uid=;PWD=;"
conn.Open
sql="select * from images where id="&id
rs.open sql,conn,1,1

if not rs.eof then
   UserID=rs("UserID")
   Title=rs("Title")
   Comment=rs("Comment")
   PicType=rs("PicType")
   RegTime=rs("RegTime")
end if
rs.close
%>
<HTML>
<HEAD>
<title>����У��¼-�༶���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="ͬѧ��ͬѧ¼��У�ѡ���ʦ��ѧУ���༶������">
<meta name="web_designner" content="�����">
<STYLE type=text/css>
</STYLE>
<LINK href="../scss.css" rel=stylesheet>
<script language="javascript">
function Del(URL)
   {
    if (confirm("��ȷ��Ҫɾ����"))
       {
        //window.open("notebook.asp?curID="+curid+"&act=del","top","toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=640,height=400")
       window.location.href=URL
       }
   }
function personalinfo(userid){
        window.open("popwindow_perinfo.asp?userID="+userid,"top","toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=580,height=400")
    }
</script>
</HEAD>
<BODY BGCOLOR=#FFFFFF leftmargin="0" topmargin="0"><CENTER><!--#include file=../top.htm-->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center" valign="top"> <br>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="20" align="left" valign="top" class="topic">�༶���</td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="5" cellpadding="5">
        <tr>
          <td align="center">
<IMG SRC="pic/show.asp?id=<%=id%>" border=1>
</td>
        </tr>
      </table>
      <table width="510" border="0" cellspacing="1" cellpadding="1">
        <tr>
          <td width="80" align="left" height="22">��Ƭ���⣺</td>
          <td width="430" class="topic2"><%=Title%></td>
        </tr>
        <tr>
          <td width="80" align="left" height="22">��Ƭ�ṩ��</td>
          <td width="430" class="topic"><a href="javascript:personalinfo('<%=UserID%>');"><%=UserID%></a></td>
        </tr>
          <tr>
            <td width="80" align="left" height="22">��Ƭ���ͣ�</td>
            <td width="430" class="topic"> <%=PicType%> </td>
          </tr>
          <tr>
            <td width="80" align="left">��Ƭ˵����</td>
            <td width="430" class="topic"><%=Comment%> </td>
          </tr>
          <tr>
            <td width="80" align="left">����ʱ�䣺</td>

          <td width="430" class="topic"><%=RegTime%> </td>
          </tr>

      </table>
      <hr width="510" size="1" noshade align="center">
      <br>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td align="center" class="topic"><a href="javascript:window.close();">�رմ���</a></td>
        </tr>
      </table>
      <br>
    </td>
  </tr>
</table>
</BODY>
</HTML>