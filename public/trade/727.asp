<div align=center>
  <table height=54 cellspacing=0 cellpadding=0 width=770 border=0>
    <tbody> 
    <tr> 
        <td width=285 height="54"> <img src="star/logo.jpg" width="210" height="54"></td>
      <td width=68></td>
      <td valign=middle> 
        <div align=center><b> 
            <%
		username=session("username")
if username="" then
		%>
            <a href="signin.asp"><font color="#000000">Sign in</font></a> - <a href="register.asp"><font color="#000000">Join 
            now</font></a> - Help&nbsp; Today: <%=date()%> 
    <%else%>
            Today: <%=date()%>&nbsp;&nbsp; Welcome:<%=username%>&nbsp;&nbsp;<a href="script/login.asp">logout </a> &nbsp;&nbsp;<a href="viewbasket.asp" target="_blank">View 
            Basket</a> 
           
            <%
		   set rs1=server.CreateObject("adodb.recordset")
		   set rs2=server.CreateObject("adodb.recordset")
		   sql="select id from back where ifopen=false and ask_member='"&username&"'"
		   sql2="select id from back where ask_member='"&username&"'"
		   rs2.open sql2,conn,1,1
		   ifexistback=rs2.recordcount
		   rs2.close
		   set rs2=nothing
		   rs1.open sql,conn,1,1
		   reds=rs1.recordcount
		%>
            <%if username<>"" then%>
            </b>
            <div align="center"> 
              <%if reds=0 then%>
              <img src="images/backno.gif"> <font color="#000066"><strong>You 
              have no new </strong></font><strong><font color="#000066">feedback</font></strong> &nbsp;&nbsp;
			  <%if ifexistback>0 then%>
			  <a href="view_back.asp">view old</a>
			  <%
			  
			  end if
			  %>
              <%else%>
              <img src="images/back.gif"> <font color="#000066"><strong>You have 
              <font color="#0000ff">
			  <%=reds%></font>new
			   
			   message/feedback
			   
			     
			   <a href="view_back.asp">view</a>
              <%end if%>
            </div>
            <%
		rs1.close
		end if
end if
		%>      <b> </b></div>
		</td>
    </tr>
    </tbody> 
  </table>
  <table height=47 cellspacing=0 cellpadding=0 width=770 border=0>
    <tbody>
      <tr> 
        <td width=359 height=47></td>
        <td width=10 height=47> <a href="index.asp" target="_blank"> <img src="star/home.jpg" border=0 width="86" height="28"> 
          </a></td>
        <td width=133 height=47> <a href="index.asp"> <img src="star/pro_c.jpg" border=0 width="133" height="31"> 
          </a></td>
        <td width=141 height=47> <a href="for_buyer.asp" targer="_blank"> <img src="star/for%20buyers.jpg" border=0 width="141" height="31"> 
          </a> </td>
        <td width=137 height=47> <a href="favorite.asp" targer="_blank"> <img src="star/favorite.jpg" border=0 width="134" height="31"> 
          </a> </td>
      </tr>
    </tbody>
  </table>
</div>
<div align=center> </div>
<div align=center> 
  <center>
    <table height=16 cellspacing=0 cellpadding=0 width=770 border=0>
  
      <tbody> 
      <tr> 
          <td width="100%" height="16">
		  <table width="770" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td background="images/bg_top.jpg">¡¡</td>
              </tr>
              <tr>
                <td background="images/bg_botton.jpg">
				<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="770" height="60">
                    <param name="movie" value="swf/top.swf">
                    <param name="quality" value="high">
                    <param name="wmode" value="transparent">
                    <embed src="swf/top.swf" width="769" height="60" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" wmode="transparent"></embed> 
                  </object>
				  </td>
              </tr>
            </table> </td>
      </tr>
      </tbody>
    </table>
  </center>
</div>
