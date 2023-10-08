<!--#include file="admin_conn.asp" -->
<%
dim rs,title
set rs=conn.execute("select ltname from admin_copyc")
title=rs("ltname")
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 3.0">
<link rel="stylesheet" type="text/css" href="css/forum.css">
<title><%=title%>-管理页面</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<style>
.navPoint
	{
	font-family: Webdings;
	font-size:9pt;
	color:white;
	cursor:hand;
	}
p{
	font-size:9pt;
}
</style>
<script language="javascript">
var iFm
iFm=1

function switchSysBar(){
	if (switchPoint.innerText==3){
		switchPoint.innerText=4
		document.all("frmTitle").style.display="none"
	}
	else{
		switchPoint.innerText=3
		document.all("frmTitle").style.display=""
	}
}
</script>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="MARGIN: 0px" scroll="no" oncontextmenu=self.event.returnValue=false >
<table border="0" cellspacing="0" cellpadding="0" width="100%" height="100%" align=center>
  <tr> 
    <td rowspan="4" align="center" valign="top" nowrap id="frmTitle" name="frmTitle"> 
      <Iframe id=leftn name=leftn style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 120; Z-INDEX: 2" scrolling=no frameborder=0 src="admin_leftmen.asp" ></Iframe></td>
    <td height="100%" rowspan="4" valign="bottom" bgcolor="#006699" style="width:5pt" > 
      <table cellspacing=0 border=0 height=100%>
        <tr> 
          <td style="height:100%" onclick="javascript:switchSysBar()"><span class="navPoint" id="switchPoint" title="关闭/打开左栏">3</span> 
      </table></td>
    <td height="100%" style="width:100%">
	<Iframe id="frmright" name="frmright" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1" scrolling="auto" frameborder="0" src="admin_index.asp"> 
      </Iframe></td>
  </tr>
  <tr> 
</table>
</body>
</html>