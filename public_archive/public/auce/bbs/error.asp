<!--#include file="conn.asp" -->
<!--#include file="mymem.asp" -->
<table width="898" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#92b9fb">
  <tr bgcolor="#92b9fb" background="backimg/bg1.gif"> 
    <td height="27" colspan="2" align="center" background="backimg/bg1.gif"><font color="#FFFFFF"><strong>��̳��Ϣ</strong></font></td>
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
        <li><b>�����Ӻ��Զ�������ҳ</b> 
        <li>�����������������ȷ�����ӽ���. </li>
        <li><a href="index.asp">������ҳ</a></li>
         <script>window.tm = setInterval("location.href='index.asp'", 5000)</script>
        <%end select%>
		<li></li>
        <li><a href="#" onClick="history.back()"> ���˷��ղ�ҳ��</a>
        <li><li> 
          <%call listchangeLt()%>
      </ul></td>
  </tr>
</table>
<!--#include file="end.asp" -->