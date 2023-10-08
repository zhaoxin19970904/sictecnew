<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Premium Superior Models |Rockets |Spacecraft | Satellites| Airplane | Space station| lunar module | Launcher | Professional Aerospace Models Provider</title>
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
.text {
	FONT-SIZE: 12px;
}
.input {
border:1px solid #666666; padding:1px; FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: "Verdana", "Arial"; HEIGHT: 19px; BACKGROUND-COLOR: #FAFEE3
}

</STYLE>
<link rel="stylesheet" href="css/style.css" type="text/css">
<link rel="stylesheet" href="css/css.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" topMargin=0 text="#000000">
<form name="form1" method="post" action="research.asp">
  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="80%">
    <tr>
      <td valign="top"> 
        <table width="900" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr>
            <td><!--#include file="0727.asp"--></td>
          </tr>
        </table>
        <table width="947" border="0" cellspacing="0" cellpadding="0" align="center" height="481">
          <tr>
          <td valign="top">
            <table width="947" border="0" cellspacing="0" cellpadding="0" align="center" height="538">
              <tr> 
                  <td width="159" valign="top" height="518"> 
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
                    <table width="221" border="0" cellspacing="0" cellpadding="0" height="364">
                      <tr> 
                      <td background="star/hyh_1_66.jpg" valign="top"> 
                        <table width="181" border="0" cellspacing="0" cellpadding="0" align="center" height="192">
                          <tr>
                              <td bgcolor="#FFFFFF" height="106" valign="top"> 
                                <table width="161" border="0" align="center" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td width="161" background="../imgs/bg_03.gif"> 
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
                                        <a href="showmodels.asp?original=<%=original%>&mo_name=<%=mo_name%>" target="_blank"> 
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
%>
                                    </td>
                                  </tr>
                                </table></td>
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
                    <table width="221" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td valign="top"><img src="star/hyh_1_66_1.jpg" width="221" height="133"></td>
                    </tr>
                  </table>
                </td>
                  <td width="788" valign="top" height="518"> 
                    <table width="720" border="0" cellspacing="0" cellpadding="0" align="center" height="595">
                      <tr> 
                      <td><table width="720" border="0" cellspacing="0" cellpadding="0" align="center" height="22">
                            <tr> 
                              <td width="217" height="41"> <div align="center"><font size="3"> 
                                  <img src="star/modelshowroom.jpg" width="138" height="13"></font></div></td>
                              <td width="503" height="41"> <table width="500" border="0" cellspacing="0" cellpadding="0" align="center">
                                  <tr> 
                                    <td width="500" background="star/hyh_2_5.jpg" height="5"></td>
                                  </tr>
                                </table></td>
                            </tr>
                          </table></td>
                    </tr>
					<tr> 
                        <td height="508" valign="top"> 
                          <%
					 'set rs=server.CreateObject("adodb.recordset")
					 sql="select id,photo1,ratio,name from modeldb order by id ASC"
					 rs.open sql,conn,1,1
	 if rs.eof then
          response.write "Sorry,there is no model of this kind for the moment。"
          response.end
     end if

if len(request("page"))=0 then
  iPage=1 
  else
   iPage=request("page")
 end if

 rs.pagesize=9
  rs.absolutepage=iPage
 rs.cachesize=rs.pagesize
  i=0
dim id(),photo1(),ratio(),m_name()
do while not rs.eof and (i<rs.pagesize)
          i=i+1
redim preserve id(i),photo1(i),ratio(i),m_name(i)
          id(i)=rs("id")
		  m_name(i)=rs("name")
          photo1(i)=rs("photo1")
		  ratio(i)=rs("ratio")
		  rs.movenext
       loop
		i=1 
Do while i<=ubound(id)
%>
                          <table width="100" border="0" cellspacing="0" cellpadding="0" align="center">
                    <tr> 
                      <td width="500" height="5"></td>
                    </tr>
                  </table>
                  <table width="720" border="0" cellspacing="0" cellpadding="0" align="center" height="122">
                    <tr> 
                      <td width="355" height="107" valign="top"> 
                        <table width="100" border="0" cellspacing="0" cellpadding="0" align="center">
                          <tr> 
                            <td width="500" height="5"></td>
                          </tr>
                        </table>
                        <table width="180" border="0" cellspacing="0" cellpadding="0" height="131">
                          <tr> 
                            <td height="150">
                                      <table width="159" border="0" cellspacing="1" cellpadding="0" height="139" align="center" background="star/hyh_d.jpg">
                                        <tr> 
                                          <td bgcolor="#FFFFFF"> 
                                            <div align="center">
<a href="#" onClick="javascript:window.open('showmodel.asp?id=<%=id(i)%>','','width=600,height=700,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
                                              <img src="manage/script/photo1/<%=trim(photo1(i))%>" width="151" height="131" border="0"></a></div>
                                          </td>
								</tr>
                              </table> </td>
                          </tr>
                        </table>
                        <table width="181" border="0" cellspacing="0" cellpadding="0" align="left" height="24">
                          <tr> 
                            <td> 
                              <div align="center"><font size="2">
							  <%=m_name(i)%>&nbsp;&nbsp;&nbsp;
							  

							  <%i=i+1%>
							  </font></div>
                            </td>
                          </tr>
                        </table>
                      </td>
                      <td width="355" height="107" valign="top"> 
                        <table width="100" border="0" cellspacing="0" cellpadding="0" align="center">
                          <tr> 
                            <td width="500" height="5"></td>
                          </tr>
                        </table>
						
<%if i<=ubound(id) then %>						
                        <table width="180" border="0" cellspacing="0" cellpadding="0" height="131">
                          <tr> 
                            <td height="150"> 
                                      <table width="159" border="0" height="139" cellspacing="1" cellpadding="0" align="center" background="star/hyh_d.jpg">
                                        <tr> 
                                          <td width="153" bgcolor="#FFFFFF"> 
                                            <div align="center">
<a href="#" onClick="javascript:window.open('showmodel.asp?id=<%=id(i)%>','','width=600,height=700,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
                                              <img src="manage/script/photo1/<%=trim(photo1(i))%>" width="151" height="131" border="0"></a></div>
                                          </td>
                                </tr>
                              </table>
                            </td>
                          </tr>
                        </table>
                        <table width="181" border="0" cellspacing="0" cellpadding="0" align="left" height="24">
                          <tr> 
                            <td> 
                              <div align="center"><font size="2">
							  <%=m_name(i)%>&nbsp;&nbsp;&nbsp;
							  
							  </font></div>
                            </td>
                          </tr>
                        </table>
						
<% 
i=i+1
end if
%>						
                      </td>
                      <td width="355" height="107" valign="top"> 
                        <table width="100" border="0" cellspacing="0" cellpadding="0" align="center">
                          <tr> 
                            <td width="500" height="5"></td>
                          </tr>
                        </table>
						
						
						
<%if i<=ubound(id) then %>						
                        <table width="180" border="0" cellspacing="0" cellpadding="0" height="131">
                          <tr> 
                            <td height="150"> 
                                      <table width="159" border="0" cellspacing="1" cellpadding="0" height="139" align="center" background="star/hyh_d.jpg">
                                        <tr> 
                                          <td bgcolor="#FFFFFF"> 
                                            <div align="center">
<a href="#" onClick="javascript:window.open('showmodel.asp?id=<%=id(i)%>','','width=600,height=700,toolbar=no,location=no,directories=no,status=no,resizable=no,scrollbars=yes','')"> 
                                              <img src="manage/script/photo1/<%=trim(photo1(i))%>" width="151" height="131" border="0"></a></div>
                                          </td>
                                </tr>
                              </table>
                            </td>
                          </tr>
                        </table>
                        <table width="181" border="0" cellspacing="0" cellpadding="0" align="left" height="24">
                          <tr> 
                            <td> 
                              <div align="center"><font size="2">
							  <%=m_name(i)%>&nbsp;&nbsp;&nbsp;
							 
							  </font></div>
                            </td>
                          </tr>
                        </table>
<% 
i=i+1
end if
%>						
                      </td>
                    </tr>
                  </table>
<%
  loop
%>
</td>
                    </tr>
					<tr> 
                      <td><table align="right" class="text">
                            <tr> 
                              <td height="21"> <div align="right">Page No.<font color="#FF0000"> 
                                  <%if len(inputpage)=0 and isnumeric(inputpage) then%>
                                  <%=iPage%> 
                                  <%else%>
                                  <%=inputpage%> 
                                  <%end if%>
                                  </font> &nbsp;&nbsp;&nbsp;&nbsp;Total <font color="#FF0000"><%=rs.recordcount%></font> 
                                  Models&nbsp;&nbsp;&nbsp;&nbsp;Total <font color="#FF0000"><%=rs.pagecount%></font> 
                                  Pages &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                                  <%
if cint(iPage)=1 or cint(inputpage)=1 then
%>
                                  <font class="dateformat">First Page | Previous 
                                | </font> 
                                  <%
else
%><font class="dateformat"><a href="index.asp?page=1">First Page</a></font> 
                                  |
            <font class="dateformat"><a href="index.asp?page=<%=iPage-1%>">Previous</a></font> 
                                  | 
                                  <%end if%>
                                  <%
 if cint(iPage)=cint(rs.pagecount) or cint(inputpage)=cint(rs.pagecount) then
%>
                                  <font class="dateformat">Next | Last Page </font> 
                                  <%else%> <a href="index.asp?page=<%=iPage+1%>">
                                Next</a> 
                                  | <a href="index.asp?page=<%=rs.pagecount%>">
                                Last Page</a> 
                                  <%
end if
%>
                                </div></td>
                            </tr>
                          </table></td>
                    </tr>
                  </table>
                  </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
        <table width="948" border="0" cellspacing="0" cellpadding="0" align="center" height="60" class="text">
          <tr valign="top"> 
            <td width="9"><!--#include file="0919.asp"--></td>
          </tr>
        </table>
      </td>
  </tr>
</table>
  </form>
<%
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
%>
</body>
</html>