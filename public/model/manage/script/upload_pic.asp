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
      <td height="25" align="center"> <p>上传<font color="#FF0000"> 大</font> ( Big 
          )图片&nbsp;&nbsp;&nbsp;</p>
        <p>&nbsp; 
          <input name="file1" type=file class=form_pane id="file1" size="50">
          <input name="submit" type=submit class=form_pane id="submit" value="上传">
        </p></td>
    </tr>  </table>
  </form> 
</body>  
</html>