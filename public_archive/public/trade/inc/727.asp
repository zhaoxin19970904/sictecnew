 
<script language="JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>

<body onLoad="MM_preloadImages('star/msh_product1.jpg','star/msh_company1.jpg','star/msh_contact1.jpg')">
<div align=center> 
  <div align=center>
    <table height=54 cellspacing=0 cellpadding=0 width=770 border=0>
      <tbody> 
      <tr> 
        <td width=285 height="54"> 
          <table width="234" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td height="31" width="18">&nbsp;</td>
              <td height="31" width="224"><img src="star/logo.jpg" width="210" height="54"></td>
            </tr>
          </table>
        </td>
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
            Today: <%=date()%>&nbsp;&nbsp; Welcome:<%=username%>&nbsp;&nbsp;<a href="script/login.asp">logout 
            </a> &nbsp;&nbsp;<a href="viewbasket.asp" target="_blank">View Basket</a> 
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
              have no new </strong></font><strong><font color="#000066">feedback</font></strong> 
              &nbsp;&nbsp; 
              <%if ifexistback>0 then%>
              <a href="view_back.asp">view old</a> 
              <%
			  
			  end if
			  %>
              <%else%>
              <img src="images/back.gif"> <font color="#000066"><strong>You have 
              <font color="#0000ff"> <%=reds%> &nbsp; </font>new message/feedback 
              <a href="view_back.asp">view</a> 
              <%end if%>
              </strong></font></div>
            <font color="#000066"><strong> 
            <%
		rs1.close
		end if
end if
		%>
            <b> </b></strong></font></div>
        </td>
      </tr>
      </tbody> 
    </table>
    <table width="770" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
          <table height=16 cellspacing=0 cellpadding=0 width=497 border=0 align="right">
            <tbody> 
            <tr> 
              <td width=99 height=21> 
                <div align="right"><a href="index.asp" target="_blank"> </a><a href="index.asp"><img src="star/msh_home.jpg" width="82" height="24" border="0"></a></div>
              </td>
              <td width=139 height=21> 
                <div align="right"><a href="index.asp"> </a><a href="index.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image12','','star/msh_product1.jpg',1)"><img name="Image12" border="0" src="star/msh_product.jpg" width="133" height="24"></a></div>
              </td>
              <td width=136 height=21> 
                <div align="right"><a href="for_buyer.asp" targer="_blank"> </a><a href="CompanyProfile.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image10','','star/msh_company1.jpg',1)"><img name="Image10" border="0" src="star/msh_company.jpg" width="128" height="24"></a></div>
              </td>
              <td width=120 height=21> 
                <div align="right"><a href="favorite.asp" targer="_blank"> </a><a href="contactus.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image11','','star/msh_contact1.jpg',1)"><img name="Image11" border="0" src="star/msh_contact.jpg" width="114" height="24"></a></div>
              </td>
            </tr>
            </tbody> 
          </table>
        </td>
      </tr>
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
                <td background="images/bg_top.jpg">&nbsp;</td>
              </tr>
              <tr> 
                <td background="images/bg_botton.jpg"> <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="770" height="60">
                    <param name="movie" value="swf/top.swf">
                    <param name="quality" value="high">
                    <param name="wmode" value="transparent">
                    <embed src="swf/top.swf" width="769" height="60" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" wmode="transparent">
                    </embed> 
                  </object> </td>
              </tr>
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
    </center>
  </div>
</div>
<div align=center></div>
