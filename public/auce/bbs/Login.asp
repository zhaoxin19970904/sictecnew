<!--#include file="CONN.ASP" -->
<!--#include file="md5.asp" -->
<%
if request.Form("username")<>"" then
call checklogin()
end if
sub checklogin()
dim user,upwd,rs,usercookies,gourl,userip
userip=request.ServerVariables("REMOTE_ADDR")
gourl=request.Form("thispagename")
user=trim(replace(request.Form("username"),"'",""))
upwd=trim(replace(request.Form("pwd"),"'",""))
usercookies=cint(request.form("cookies"))
if gourl="" or left(gourl,9)="login.asp" then
gourl="index.asp"
end if

set rs=conn.execute("select user,password,id from userinfo where user='"&user&"'")
if not rs.eof then 
     if rs("password")<>MD5(upwd) then
	 errmess=errmess&"<li>����û����������벻��ȷ!"
	exit sub
	end if
select case usercookies
case 0
	Response.Cookies("renwen")("usercookies") = usercookies
case 1
   	Response.Cookies("renwen").Expires=Date+1
	Response.Cookies("renwen")("usercookies") = usercookies
case 2
	Response.Cookies("renwen").Expires=Date+31
	Response.Cookies("renwen")("usercookies") = usercookies
case 3
	Response.Cookies("renwen").Expires=Date+365
	Response.Cookies("renwen")("usercookies") = usercookies
case other
	Response.Cookies("renwen").Expires=date+0.01
	Response.Cookies("renwen")("usercookies") = usercookies
end select
response.cookies("renwen")("user")=rs("user")
response.cookies("renwen")("passedok")="ofdkjduy"
conn.execute("update userinfo set userlog=userlog+1,userjf=userjf+1,lastlogin='"&now()&"',lastlogIP='"&userip&"' where user='"&user&"'")
	founderr=true
	errmess=errmess&"<li><b>���Ѿ��ɹ���¼.</b> �����Զ�������ҳ.<li><li><a href=index.asp>������ҳ</a>"
	errmess=errmess&"<script>window.tm = setInterval(""location.href='index.asp'"", 3000)</script>"
	rs.close
	else
	errmess=errmess&"<li>����û����������벻��ȷ!"
    rs.close
  end if
end sub
if request.QueryString("userlogout")="loginout" then
dim userip
userip=request.ServerVariables("REMOTE_ADDR")
conn.execute=("update userinfo set logout=logout+1,lastlogIP='"&userip&"' where user='"&request.cookies("renwen")("user")&"'")
Response.Cookies("renwen").Expires=Date+0.001
response.cookies("renwen")("user")=""
response.cookies("renwen")("passedok")=""
userislogin=false
	founderr=true
	errmess=errmess&"<li><b>���Ѿ��ɹ��˳���¼.</b> �����Զ�������ҳ.<li><li><a href=index.asp>������ҳ</a>"
	errmess=errmess&"<script>window.tm = setInterval(""location.href='index.asp'"", 3000)</script>"
end if
%>
<!--#include file="mymem.asp" -->    
<FORM  action=Login.asp  method=post name="Login">
<%
if founderr=true then
response.cookies("rewin")("errmess")=errmess
response.Redirect("error.asp?founderr=mess")
end if
%>
  <table cellpadding=3 cellspacing=1 class=fs height=219 width="898" align="center" bgcolor="#92b9fb">
    <tr><td bgcolor="#FFFFFF" colspan=2>
	<%if errmess<>"" then response.Write(errmess) end if%></td></tr>
	<tr bgcolor="#006699"> 
      <td height="27" colspan="2" background="backimg/bg1.gif"> 
        <div align="center"><font color="#FFFFFF"><strong>�û���¼</strong></font></div>
      </td>
    </tr>
    <tr bgcolor="#E0E4FE"> 
      <td width="162" height="24" bgcolor="#F2F2F2"> 
        <div align="center"><font color="#000033">�û���:</font></div>
      </td>
      <td width="384" height="24" bgcolor="#FFFFFF"> <font color="#000033"> 
        <input class=inputline name=username size=15>
        <font size="2"><a href="regedit.asp">û��ע��</a></font></font></td>
    </tr>
    <tr bgcolor="#E0E4FE"> 
      <td width="162" height="9" bgcolor="#F2F2F2"> 
        <div align="center"><font color="#000033">��&nbsp;&nbsp;��:</font></div>
      </td>
      <td width=384 height="9" bgcolor="#FFFFFF"> <font color="#000033"> 
        <input class=inputline name=pwd size=15 
            type=password>
        <font size="2"><a href="lostpassport.asp">��������</a>?</font></font></td>
    </tr>
    <tr bgcolor="#E0E4FE"> 
      <td width="162" height="96" bgcolor="#F2F2F2"><font color="#000033">��ѡ��cookies�ڱ���������б����ʱ�䣨�������������û�����Ĭ��ѡ�񣩡�</font></td>
      <td width="384" height="96" bgcolor="#FFFFFF"> <font color="#000033"> 
        <input type="radio" name="cookies" value="0" checked>
        ������ ���ر��������ʧЧ<br>
        <input type="radio" name="cookies" value="1">
        һ��<br>
        <input type="radio" name="cookies" value="2">
        һ����<br>
        <input type="radio" name="cookies" value="3">
        һ�� </font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="33" colspan=2> 
        <div align="left"></div>
        <table width="188" height="26" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td> 
           
                <input name=form1 type=submit value=" �� ¼ ">
              </td>
            <td> 
                <input class=main name=reset2 type=reset value=" �� �� ">
              </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</FORM >
<!--#include file="end.asp" -->
</BODY>
</HTML>
