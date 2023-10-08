<html>
<%
if session("admin_name")="" then Response.Redirect("../index.asp")
session("id")=request.querystring("id")
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href=style.css rel=STYLESHEET type=text/css>
<script language="jscript">
<!--
  function check(form1)
  {
    var fig=true
	
    if(form1.file1.value=="")
	  {
	    window.alert("没有要上传的图片,请先选择！")
		fig=false
	  }
	return fig
  }  
  
//-->
</script>

</head>
<body bgcolor=#ffffff topmargin="100">

<form name="form1" method="post" action="saveuploadpic.asp" enctype="multipart/form-data" onsubmit="return check(this)">
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="25" align="center"> <p>上传<font color="#FF0000"> 大</font> ( Big 
          )图片&nbsp;&nbsp;&nbsp;</p>
        <p>&nbsp; 
          <input name="file1" type=file class=form_pane id="file1" size="50">
          <input name="submit" type=submit class=form_pane id="submit" value="确定修改">
        </p>
        <p><a href="upload1.asp">修改小图片</a> </p></td>
    </tr>  </table>
  </form> 
</body>  
</html>