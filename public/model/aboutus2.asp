<!--#include file="conn.asp" -->
<html>
<head>
<title>aboutus</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<link rel="stylesheet" href="css/style.css" type="text/css">
<link rel="stylesheet" href="css/css.css" type="text/css">
<SCRIPT language=Javascript>
<!--
function onChange(i){
childSort=document.all("child" + i);
//theTd=document.all("td" + i);
	if(childSort.style.display=="none"){
//		theTd.bgcolor="#ffffff";
		childSort.style.display="";}
	else{
//		theTd.bgcolor="#000000";
		childSort.style.display="none";}
}
//-->
</SCRIPT>
<script language="vbscript">
sub search_Onclick
<!--
dim frmtmp
set frmtmp=document.form1
frmtmp.submit
end sub
-->
</script>
<STYLE type=text/css>
.menu {
	line-height:30px;
}
</STYLE>
</head>
<body bgcolor="#FFFFFF" topMargin=0 text="#000000">
<form name="form1" method="post" action="research.asp">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="98%">
  <tr>
    <td valign="top"> 
      <table width="900" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td><!--#include file="0727.asp"--></td>
        </tr>
      </table>
        <table width="947" border="0" cellspacing="0" cellpadding="0" align="center" height="279">
          <tr> 
            <td valign="top"> 
              <table width="947" border="0" cellspacing="0" cellpadding="0" align="center" height="251">
                <tr> 
                  <td width="159" valign="top" height="290"> 
                    <table width="221" border="0" cellspacing="0" cellpadding="0" height="17">
                      <tr> 
                        <td><img src="star/hyh_1_5.jpg" width="221" height="58"></td>
                      </tr>
                    </table>
                    <table width="221" border="0" cellspacing="0" cellpadding="0" height="17">
                      <tr> 
                        <td><img src="star/hyh_1_6.jpg" width="221" height="20"></td>
                      </tr>
                    </table>
                    <table width="221" border="0" cellspacing="0" cellpadding="0" height="17">
                      <tr> 
                        <td background="star/hyh_1_66.jpg"> 
                          <table width="181" border="0" cellspacing="0" cellpadding="0" align="center" height="235">
                            <tr> 
                              <td bgcolor="#FFFFFF" height="106" valign="top"> 
                                <table width="153" border="0" align="center" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td width="153" background="../imgs/bg_03.gif"> 
                                      <p> 
                                        <%
	set rs= server.createobject("adodb.recordset")
	sql="select class_name from class" 
	rs.open sql,conn,1,1 
	j=241
	do while not rs.eof and not rs.bof %>
                                        <%  
	j=j+1%>
                                        <a href="javascript:onChange(<%=j%>)"> 
                                        <b><font size="3" color="#3399CC">+</font></b> 
                                        </a>&nbsp; <a href='showprodir.asp?class_name=<%=rs("class_name")%>'><font size="2" color="#003399"><%=rs("class_name")%></font></a><br>
                                      <table>
                                        <tr> 
                                          <td height="4"></td>
                                        </tr>
                                      </table>
                                      <% childid=cstr("child"&j) %>
                                      <span id="<%=childid%>" style="DISPLAY: none"> 
                                      <%
set rs1= server.createobject("adodb.recordset")
mo_name=rs("class_name")
sql1="select distinct original from modeldb where type='"&mo_name&"'"
rs1.open sql1,conn,1,1
%>
                                      <%
if  not rs1.bof or rs1.eof  then
do while not rs1.eof and not rs1.bof	
original=rs1("original")
%>
                                      <img align="absMiddle" src="images/line_tri.gif" width="19" height="16"> 
                                      <img align="absMiddle" src="images/reddot.gif" width="6" height="11">&nbsp; 
                                      <a href="showmodels.asp?original=<%=original%>&mo_name=<%=mo_name%>"> 
                                      <font size="2" color="#008BAC"><b><%=original%></b></font></a> 
                                      <br>
                                      <%
rs1.movenext
loop
end if
rs1.close
set rs1=nothing
%>
                                      </span> 
                                      <% 
	rs.movenext 
	loop 
	rs.close 
	m_type=request("m_type")
	m_name1=request("m_name")
	country=request("country")
%>
                                    </td>
                                  </tr>
                                </table>
                              </td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                    </table>
                    <table width="221" border="0" cellspacing="0" cellpadding="0" height="17">
                      <tr> 
                        <td><img src="star/hyh_1_7.jpg" width="221" height="20"></td>
                      </tr>
                    </table>
                    <table width="221" border="0" cellspacing="0" cellpadding="0" height="17">
                      <tr> 
                        <td><img src="star/hyh_1_66_1.jpg" width="221" height="62"></td>
                      </tr>
                    </table>
                  </td>
                  <td width="788" valign="top" height="290"> 
                    <table width="720" border="0" cellspacing="0" cellpadding="0" align="center" height="22">
                      <tr> 
                        <td>&nbsp;</td>
                      </tr>
                    </table>
                    <table width="720" border="0" cellspacing="0" cellpadding="0" align="center" height="22">
                      <tr> 
                        <td width="217" height="41"> 
                          <div align="center"><font size="2"><img src="star/aboutus.jpg" width="73" height="13"></font></div>
                        </td>
                        <td width="503" height="41"> 
                          <table width="500" border="0" cellspacing="0" cellpadding="0" align="center">
                            <tr> 
                              <td width="500" background="star/hyh_2_5.jpg" height="5"></td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                    </table>
                    <table width="709" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="315" width="42">&nbsp;</td>
                        <td height="315" width="667" valign="top"> 
                          <div align="left"><font size="3"><br>
                            Sictec International Group Ltd. </font> </div>
                          <p align="left"><font size="3">Beijing Office:</font></p>
                          <p align="left"><font size="3">The Exchange Beijing, 
                            Suite 2110<br>
                            118 Jianguolu Yi, Chaoyang Beijing, P. R. China 100022<br>
                            Tel: 010-5129-0065<br>
                            Fax: 010-5129-0065-210<br>
                            E-mail: sictec.bj@sictec.com</font></p>
                          <p align="left"><font size="3">Toronto Office:</font></p>
                          <p align="left"><font size="3">7050 Woodbine Ave., Suite 
                            201 <br>
                            Markham (Greater Toronto Area) ON<br>
                            Canada L3R 4G8 <br>
                            Phone: 1-905-943-7571<br>
                            Fax: 1-905-943-7573<br>
                            Email: gxu@sictec.com</font></p>
                          <p align="left"><font size="3">http://www.superior-models.com<br>
                            </font></p>
                        </td>
                      </tr>
                    </table>
                    <p><font size="2"> </font></p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
        <!--#include file="0919.asp"-->
      </td>
  </tr>
</table>
</form>
</body>
</html>
