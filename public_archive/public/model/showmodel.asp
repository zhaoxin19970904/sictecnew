<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>model showroom</title>
<link rel="stylesheet" href="css/style.css" type="text/css">
<link rel="stylesheet" href="css/css.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" topMargin=0 text="#000000">
<table width="500" height="184" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
      
    <td width="737" height="184" align="left" valign="top"> <br>
      <%
	                     set rs=server.CreateObject("adodb.recordset")
						 id=request("id")
						 sql="select * from modeldb where id="&id&""
						 rs.open sql,conn,1,1
						 if rs("w_h")="heightH" then
						%>
<table width="485" border="0" cellspacing="0" cellpadding="0" height="486">
        <tr> 
    <td height="41" colspan="2"> <div align="center"><font size="3"> <img src="star/modelshowroom.jpg" width="138" height="13"> 
        </font></div>
      <br> <table width="500" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td width="500" background="star/hyh_2_5.jpg" height="5"></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="138" valign="top"> ��</td>
    <td width="347" valign="top"> ��</td>
  </tr>
  <tr> 
    <td width="138" valign="top"> <table width="113" border="0" cellspacing="0" cellpadding="0" height="486" bgcolor="#CCCCCC" align="center">
        <tr> 
          <td bgcolor="#FFFFFF"> <img src="manage/script/photo2/<%=rs("photo2")%>"> 
          </td>
        </tr>
      </table></td>
    <td width="347" valign="top"> <div align="center"><br>
        <b><font size="2">Model details</font></b><br>
        <br>
      </div>
      <table width="287" border="0" cellspacing="0" cellpadding="0" height="9">
        <tr> 
          <td width="14" height="20"> <div align="left"></div></td>
          <td width="273" height="20"> <table width="187" height="385" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="115" height="12">Type:&nbsp;&nbsp;</td>
                <td width="160" height="12"><font color="#828653"><%=rs("type")%></font></td>
              </tr>
              <tr> 
                <td height="16" width="115">Name:&nbsp;&nbsp;</td>
                <td height="16" width="160"><font color="#828653"><%=rs("name")%></font></td>
              </tr>
              <tr> 
                <td height="16" width="115">Country:&nbsp;&nbsp;</td>
                <td height="16" width="160"><font color="#828653"><%=rs("original")%></font></td>
              </tr>
              <tr> 
                <td height="16" width="115">Scale:&nbsp;&nbsp;</td>
                <td height="16" width="160"><font color="#828653"><%=rs("ratio")%></font></td>
              </tr>
              <tr> 
                <td height="16" width="115">Size:&nbsp;&nbsp;</td>
                <td height="16" width="160"><font color="#828653"><%=rs("size")%></font></td>
              </tr>
              <tr> 
                <td height="16" width="115">Material:</td>
                <td height="16" width="160"><font color="#828653"><%=rs("material")%></font></td>
              </tr>
              <tr> 
                <td height="16" width="115">Lead Time:</td>
                <td height="16" width="160"><font color="#828653"><%=rs("leadtime")%></font></td>
              </tr>
              <tr> 
                <td height="16" width="115">Packing:&nbsp;&nbsp;</td>
                <td height="16" width="160"><font color="#828653"><%=rs("packing")%></font></td>
              </tr>
              <tr> 
                <td height="16" width="115">Sample Available:</td>
                <td height="16" width="160"><font color="#828653"> 
                  <% sa=rs("sample")
                      if sa then
                      response.write "Yes"
                      else
                    response.write "No"
                    end if%>
                  </font></td>
              </tr>
              <tr> 
                <td height="16" width="115">Mini order quantity:&nbsp;&nbsp;<font color="#828653"><%=rs("Miniquantity")%></font></td>
                <td height="16" width="160"></td>
              </tr>
              <tr> 
                <td height="16" width="115"> Notes:&nbsp;&nbsp;<font color="#828653"><%=rs("Notes")%> 
                  </font></td>
                <td height="16" width="160"> </td>
              </tr>
              <tr> 
                <td height="16" width="115">&nbsp;</td>
                <td height="16" width="160"></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>
<p>
  <%else%>
</p>
      <table width="500" border="0" cellspacing="0" cellpadding="0" height="18" align="left">
        <tr> 
    <td height="41" > <div align="center"><font size="3"> <img src="star/modelshowroom.jpg" width="138" height="13"> 
        </font></div>
      <br> <table width="500" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td width="500" background="star/hyh_2_5.jpg" height="5"></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="500" height="10" align="center" valign="top"> <div align="center"> 
        <table width="113" border="0" cellspacing="0" cellpadding="0" height="486" bgcolor="#CCCCCC" align="center">
          <tr> 
            <td bgcolor="#FFFFFF"> <img src="manage/script/photo2/<%=rs("photo2")%>"> 
            </td>
          </tr>
        </table>
      </div></td>
  </tr>
  <tr> 
    <td width="500" height="10" valign="top"> <div align="center"> 
        <table width="400" height="9" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="20" height="20"> <div align="left">��</div></td>
            <td width="380" height="20"> <table width="377" height="366" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td height="12" width="138">Type:&nbsp;&nbsp;</td>
                  <td height="12" width="239"><font color="#828653"><%=rs("type")%></font></td>
                </tr>
                <tr> 
                  <td height="16" width="138">Name:&nbsp;&nbsp;</td>
                  <td height="16" width="239"><font color="#828653"><%=rs("name")%></font></td>
                </tr>
                <tr> 
                  <td height="16" width="138">Country:&nbsp;&nbsp;</td>
                  <td height="16" width="239"><font color="#828653"><%=rs("original")%></font></td>
                </tr>
                <tr> 
                  <td height="16" width="138">Scale:&nbsp;&nbsp;</td>
                  <td height="16" width="239"><font color="#828653"><%=rs("ratio")%></font></td>
                </tr>
                <tr> 
                  <td height="16" width="138">Size:&nbsp;&nbsp;</td>
                  <td height="16" width="239"><font color="#828653"><%=rs("size")%></font></td>
                </tr>
                <tr> 
                  <td height="16" width="138">Material:&nbsp;&nbsp;</td>
                  <td height="16" width="239"><font color="#828653"><%=rs("material")%></font></td>
                </tr>
                <tr> 
                  <td height="16" width="138">Lead Time:&nbsp;&nbsp;</td>
                  <td height="16" width="239"><font color="#828653"><%=rs("leadtime")%></font></td>
                </tr>
                <tr> 
                  <td height="16" width="138">Packing:&nbsp;&nbsp;</td>
                  <td height="16" width="239"><font color="#828653"><%=rs("packing")%></font></td>
                </tr>
                <tr> 
                  <td height="16" width="138">Sample Available:</td>
                  <td height="16" width="239"><font color="#828653"> 
                    <% sa=rs("sample")
                                          if sa then
                                          response.write "Yes"
                                          else
                                          response.write "No"
                                          end if%>
                    </font></td>
                </tr>
                <tr> 
                  <td height="16" width="138">Mini order quantity:&nbsp;&nbsp;<font color="#828653"><%=rs("Miniquantity")%></font></td>
                  <td height="16" width="239"></td>
                </tr>
                <tr> 
                  <td height="16" width="138"> Notes:&nbsp;&nbsp;<font color="#828653"><%=rs("Notes")%> 
                    </font></td>
                  <td height="16" width="239"> </td>
                </tr>
                <tr> 
                  <td height="16" width="138">&nbsp;</td>
                  <td height="16" width="239"></td>
                </tr>
              </table></td>
          </tr>
        </table>
      </div></td>
  </tr>
</table>
<p>
  <%end if%>
  </p>
    </td>
  </tr>
</table>
<%
rs.close
 set rs=nothing
 conn.close
 set conn=nothing
%>
</body>
</html>