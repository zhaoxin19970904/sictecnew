<html>
<%
if session("admin_name")="" then Response.Redirect("../index.asp")
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href=style.css rel=STYLESHEET type=text/css>
</head>
<body bgcolor=#ffffff topmargin="100">
<form name="form1" method="post" action="savepic.asp" enctype="multipart/form-data">
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="25" align="center"> <p>�ϴ�<font color="#FF0000"> ��</font> ( Big 
          )ͼƬ&nbsp;&nbsp;&nbsp;</p>
        <p>&nbsp; 
          <input name="file1" type=file class=form_pane id="file1" size="50">
          <input name="submit" type=submit class=form_pane id="submit" value="�ϴ�">
        </p></td>
    </tr>  </table>
  </form> 
</body>  
</html>