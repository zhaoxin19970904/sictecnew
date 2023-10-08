<!--#include file="conn.asp" -->
<!--#include file="mymem.asp" -->
<table width="898" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#92b9fb">
  <tr bgcolor="#92b9fb" background="backimg/bg1.gif"> 
    <td height="27" colspan="2" align="center" background="backimg/bg1.gif"><font color="#FFFFFF"><strong>论坛信息</strong></font></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="70" colspan="2"><ul>
        <%
	dim errlb
	errlb=request.QueryString("founderr")
	select case(errlb)
	case "mess"
	response.Write(request.cookies("rewin")("errmess"))
	'response.cookies("rewin")("errmess")=""
	case "link"
	%>
        <br>
        <li><b>五秒钟后自动返回主页</b> 
        <li>发生参数错误请从正确的链接进入. </li>
        <li><a href="index.asp">返回首页</a></li>
         <script>window.tm = setInterval("location.href='index.asp'", 5000)</script>
        <%end select%>
		<li></li>
        <li><a href="#" onClick="history.back()"> 后退返刚才页面</a>
        <li><li> 
          <%call listchangeLt()%>
      </ul></td>
  </tr>
</table>
<!--#include file="end.asp" -->