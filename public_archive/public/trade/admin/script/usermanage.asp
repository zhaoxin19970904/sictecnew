<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href=../css.css rel=STYLESHEET type=text/css>
</head>
<STYLE type=text/css>
.unnamed1 {
	FONT-SIZE: 9pt
}
</STYLE>
<script>
  function user(id) { window.open("viewuser.asp?user_id="+id,"","height=400,width=600,left=190,top=0,resizable=yes,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no");} 
</script>
<%
if session("admin_name")="" then response.end
set rs=server.createobject("adodb.recordset")
%>
<!--#include file="conn.asp"-->
<body>
<div align="center">
  <%
page=request.querystring("page")
if page="" then page=1
if not(isnumeric(page)) then page=1
if page<1 then page=1
page=int(page)
sql="select * from member order by id desc"
rs.open sql,conn,1,1
if rs.eof then
response.Write ("No user!")
response.end()
 else
rs.pagesize=30
totalrec=rs.recordcount
totalpage=rs.pagecount
if page>totalpage then page=totalpage
rs.absolutepage=page
rs.cachesize=rs.pagesize
i=0
dim id(),member_uid(),psd(),sex(),firstname(),lastname(),jobtitle(),email(),tel(),country(),city()
dim address(),post(),company(),main_biz_area(),main_products(),biznature(),com_tel(),fax()
dim productbuy(),productsell(),homepage(),reg_time(),reg_ip()
do while not rs.eof and (i<rs.pagesize)
i=i+1
redim preserve id(i),member_uid(i),psd(i),sex(i),firstname(i),lastname(i),jobtitle(i),email(i),tel(i),country(i),city(i)
redim preserve address(i),post(i),company(i),main_biz_area(i),main_products(i),biznature(i),com_tel(i),fax(i)
redim preserve productbuy(i),productsell(i),homepage(i),reg_time(i),reg_ip(i)

 id(i)=rs("id")
 member_uid(i)=rs("member_uid")
 psd(i)=rs("psd")
 sex(i)=rs("sex")
 firstname(i)=rs("firstname")
 lastname(i)=rs("lastname")
 jobtitle(i)=rs("jobtitle")
 email(i)=rs("email")
 tel(i)=rs("tel")
 country(i)=rs("country")
 city(i)=rs("city")
 address(i)=rs("address")
 post(i)=rs("post")
 company(i)=rs("company")
 main_biz_area(i)=rs("main_biz_area")
 main_products(i)=rs("main_products")
 biznature(i)=rs("biznature")
 com_tel(i)=rs("com_tel")
 fax(i)=rs("fax")
 productbuy(i)=rs("productbuy")
 productsell(i)=rs("productsell")
 homepage(i)=rs("homepage")
 reg_time(i)=rs("reg_time")
 reg_ip(i)=rs("reg_ip")
rs.movenext
loop
rs.close
%>
  <font color="#000066"><strong>注册用户<br>
  </strong></font></div>
<table width="2440" height="76" border="1" cellpadding="0" cellspacing="0" class="unnamed1">
  <tr> 
    <td width="95" height="16"><div align="center"><font color="#0000CC">member_uid</font> 
      </div></td>
    <td width="95"><div align="center"><font color="#0000CC">password</font></div></td>
    <td width="94"><div align="center"><font color="#0000CC">sex</font></div></td>
    <td width="95"><div align="center"><font color="#0000CC">firstname</font></div></td>
    <td width="95"><div align="center"><font color="#0000CC">lastname</font></div></td>
    <td width="95"><div align="center"><font color="#0000CC">jobtitle</font></div></td>
    <td width="94"><div align="center"><font color="#0000CC">email</font></div></td>
    <td width="95"><div align="center"><font color="#0000CC">telphone</font></div></td>
    <td width="95"> <div align="center"><font color="#0000CC">country</font> </div></td>
    <td width="94"><div align="center"><font color="#0000CC">city</font></div></td>
    <td width="94"><div align="center"><font color="#0000CC">address</font></div></td>
    <td width="82"><div align="center"><font color="#0000CC">post</font></div></td>
    <td width="106"><div align="center"><font color="#0000CC">company</font></div></td>
    <td width="108"><div align="center"><font color="#0000CC">main_biz_area</font></div></td>
    <td width="110"><div align="center"><font color="#0000CC">main_products</font></div></td>
    <td width="111"><div align="center"><font color="#0000CC">biznature</font></div></td>
    <td width="106"><div align="center"><font color="#0000CC">company telphone</font></div></td>
    <td width="66"><div align="center"><font color="#0000CC">fax</font></div></td>
    <td width="122"><div align="center"><font color="#0000CC">product buy</font></div></td>
    <td width="142"><div align="center"><font color="#0000CC">product sell</font></div></td>
    <td width="139"><div align="center"><font color="#0000CC">homepage</font></div></td>
    <td width="111"><div align="center"><font color="#0000CC">register time</font></div></td>
    <td width="74"><div align="center"><font color="#0000CC">register IP</font></div></td>
    <td width="72"> <div align="center"><font color="#0000CC">operate</font> </div></td>
  </tr>
  <%for i=1 to ubound(member_uid)%>
  <tr> 
    <td height="16"><div align="center"><%=member_uid(i)%></div></td>
    <td><div align="center"><%=psd(i)%></div></td>
    <td><div align="center"><%=sex(i)%></div></td>
    <td><div align="center"><%=firstname(i)%></div></td>
    <td><div align="center"><%=lastname(i)%></div></td>
    <td><div align="center"><%=jobtitle(i)%></div></td>
    <td><div align="center"><%=email(i)%></div></td>
    <td><div align="center"><%=tel(i)%></div></td>
    <td><div align="center"><%=country(i)%></div></td>
    <td><div align="center"><%=city(i)%></div></td>
    <td><div align="center"><%=address(i)%></div></td>
    <td><div align="center"><%=post(i)%></div></td>
    <td><div align="center"><%=company(i)%></div></td>
    <td><div align="center"><%=main_biz_area(i)%></div></td>
    <td><div align="center"><%=main_products(i)%></div></td>
    <td><div align="center"><%=biznature(i)%></div></td>
    <td><div align="center"><%=com_tel(i)%></div></td>
    <td><div align="center"><%=fax(i)%></div></td>
    <td><div align="center"><%=productbuy(i)%></div></td>
    <td><div align="center"><%=productsell(i)%></div></td>
    <td><div align="center"><%=homepage(i)%></div></td>
    <td><div align="center"><%=reg_time(i)%></div></td>
    <td><div align="center"><%=reg_ip(i)%></div></td>
    <td><div align="center"><a href="deluser.asp?id=<%=id(i)%>" onClick="return ifdel()">delete</a></div></td>
  </tr>
  <%next%>
  <tr> 
    <td height="28" colspan="24"><font color=666666>&nbsp; </font>Total:<font color=red><%=totalpage%></font> 
	&nbsp;&nbsp;&nbsp;Page No:<%=page%><font color=666666> 
      <%if page-1>0 then%>
      <a href="usermanage.asp?page=<%=page-1%>">Pre</a> 
      <%else%>
      <%end if%>
      　 
      <%if page+1<=totalpage then%>
      <a href="usermanage.asp?page=<%=page+1%>">Next</a> 
      <%else%>
      <%end if%>
      </font></td>
  </tr>
</table>
<%end if%>
<script language="JavaScript">
function ifdel()
  {
    var fig;
     fig=confirm("Are you sure!")
	 return fig
  }
</script>
</body>
</html>
