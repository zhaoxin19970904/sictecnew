<%
title=request("title")
pictype=request("pictype")
comment=request("comment")
set connGraph=server.CreateObject("ADODB.connection")
connGraph.ConnectionString="driver={Microsoft Access Driver (*.mdb)};DBQ=" &server.MapPath("pic/images.mdb") & ";uid=;PWD=;"
connGraph.Open
set rec=server.createobject("ADODB.recordset")
rec.Open "SELECT * FROM images where id is null",connGraph,1,3
rec.addnew
rec("regtime")=now()
rec("userid")=session("userid")
rec("classid")=session("classid")
rec("title")=title
rec("pictype")=pictype
rec("comment")=comment
rec.update
session("photoid")=rec("id")
rec.close
set rec=nothing
set connGraph=nothing
%>











<HTML>
<HEAD>
<Script language="javascript">
function mysubmit(theform)
{
    if(theform.big.value=="")
    {
    alert("���������ť��ѡ����Ҫ�ϴ���jpg��gif�ļ�!")
    theform.big.focus;
    return (false);
    }
    else
    {
    str= theform.big.value;
    strs=str.toLowerCase();
    lens=strs.length;
    extname=strs.substring(lens-4,lens);
    if(extname!=".jpg" && extname!=".gif")
    {
    alert("��ѡ��jpg��gif�ļ�!");
    return (false);
    }
    }
    return (true);
}
</script>

<title>�ϴ���Ƭ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="ͬѧ��ͬѧ¼��У�ѡ���ʦ��ѧУ���༶������">
<meta name="web_designner" content="У��¼">
<STYLE type=text/css>
</STYLE>
<LINK href="../scss.css" rel=stylesheet>
</HEAD>
<BODY BGCOLOR=#FFFFFF leftmargin="0" topmargin="0"><!--#include file=../top1.htm-->
<table width="550" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center" valign="top"> <br>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="20" align="left" valign="top" class="topic">�ϴ���Ƭ</td>
        </tr>
      </table>
      <table width="510" border="0" cellspacing="1" cellpadding="1">
        <form name="mainForm" enctype="multipart/form-data" action="pic/process.asp" method=post onSubmit="return mysubmit(this)">
          <tr> 
            <td width="62" align="right" height="26">���⣺</td>
            <td width="441"><%=title%> </td>
          </tr>
          <tr> 
            <td width="62" align="right" height="26">��Ƭ���ͣ�</td>
            <td height="26" width="441"><%=pictype%></td>
          </tr>
          <tr>
            <td width="62" align="right" height="26">��Ƭ˵����</td>
            <td width="441"><%=comment%></td>
          </tr>
          <tr> 
            <td width="62" align="right" height="26">ͼƬ·����</td>
            <td width="441"> 
              <input type="file" name="big" size="20">
            </td>
          </tr>
          <tr align="center"> 
            <td colspan="2" height="26"> <br>
              <input type="submit" name="Submit2" value="��ʼ�ϴ�">
              <br>
            </td>
          </tr>
          <tr align="center"> 
            <td colspan="2" height="26" class="di">&nbsp; </td>
          </tr>
          <tr align="center"> 
            <td colspan="2" height="26" class="di">֧��jpg,gif�ļ�������ͼƬ��С�����30k���ڣ��ļ�������Ӣ�ģ�</td>
          </tr>
        </form>
      </table>
      <br>
    </td>
  </tr>
</table>
</BODY>
</HTML>