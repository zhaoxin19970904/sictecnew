
<%
dim mythispage
mythispage=thispage_name(request.ServerVariables("SCRIPT_NAME"),"title")
%>
<html>
<head>
<title><%response.Write(mythispage)%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="JavaScript">
//***********Ĭ�����ö���.*********************
tPopWait=20;//ͣ��tWait�������ʾ��ʾ��
tPopShow=5000;//��ʾtShow�����ر���ʾ
showPopStep=20;
popOpacity=99;

//***************�ڲ���������*****************
sPop=null;
curShow=null;
tFadeOut=null;
tFadeIn=null;
tFadeWaiting=null;

document.write("<style type='text/css'id='defaultPopStyle'>");
document.write(".cPopText {  background-color: #FFFFFF;color:#000000; border: 1px #6699CC solid;font-color: font-size: 12px; padding-right: 4px; padding-left: 4px; height: 20px; padding-top: 2px; padding-bottom: 2px; filter: Alpha(Opacity=0)}");
document.write("</style>");
document.write("<div id='dypopLayer' style='position:absolute;z-index:1000;' class='cPopText'></div>");


function showPopupText(){
var o=event.srcElement;
	MouseX=event.x;
	MouseY=event.y;
	if(o.alt!=null && o.alt!=""){o.dypop=o.alt;o.alt=""};
        if(o.title!=null && o.title!=""){o.dypop=o.title;o.title=""};
	if(o.dypop!=sPop) {
			sPop=o.dypop;
			clearTimeout(curShow);
			clearTimeout(tFadeOut);
			clearTimeout(tFadeIn);
			clearTimeout(tFadeWaiting);	
			if(sPop==null || sPop=="") {
				dypopLayer.innerHTML="";
				dypopLayer.style.filter="Alpha()";
				dypopLayer.filters.Alpha.opacity=0;	
				}
			else {
				if(o.dyclass!=null) popStyle=o.dyclass 
					else popStyle="cPopText";
				curShow=setTimeout("showIt()",tPopWait);
			}
			
	}
}

function showIt(){
		dypopLayer.className=popStyle;
		dypopLayer.innerHTML=sPop;
		popWidth=dypopLayer.clientWidth;
		popHeight=dypopLayer.clientHeight;
		if(MouseX-12+popWidth>document.body.clientWidth) popLeftAdjust=-popWidth-24
			else popLeftAdjust=0;
		if(MouseY+1+popHeight>document.body.clientHeight) popTopAdjust=-popHeight-24
			else popTopAdjust=0;
		dypopLayer.style.left=MouseX-12+document.body.scrollLeft+popLeftAdjust;
		dypopLayer.style.top=MouseY+1+document.body.scrollTop+popTopAdjust;
		dypopLayer.style.filter="Alpha(Opacity=0)";
		fadeOut();
}

function fadeOut(){
	if(dypopLayer.filters.Alpha.opacity<popOpacity) {
		dypopLayer.filters.Alpha.opacity+=showPopStep;
		tFadeOut=setTimeout("fadeOut()",1);
		}
		else {
			dypopLayer.filters.Alpha.opacity=popOpacity;
			tFadeWaiting=setTimeout("fadeIn()",tPopShow);
			}
}

function fadeIn(){
	if(dypopLayer.filters.Alpha.opacity>0) {
		dypopLayer.filters.Alpha.opacity-=1;
		tFadeIn=setTimeout("fadeIn()",1);
		}
}
document.onmouseover=showPopupText;

</script>
<link rel="stylesheet" href="DEFAULT.css" type="text/css">
</head>

<body leftmargin="0" topmargin="0">

<!--#include file="top.asp" -->
<table width="898" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  <tr>
    <td bgcolor="#FFFFFF"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td> 
<%
'response.Write(loadcopyc("tophtm"))
%>          </td>
        </tr>
      </table> 
<%
call readltmess()
sub readltmess()
dim rs,sqltext,notread
Set rs = Server.CreateObject("ADODB.Recordset")
sqltext="select id,fd from mess where fd='"&loginuser&"' and click=0 order by date " 
rs.open sqltext,conn,1,1
if not rs.eof then
notread=rs.recordcount
else
notread=0
end if
if userislogin then %>
      <table width="100%" border=0 cellpadding=0 cellspacing=1>
        <tr> 
          <td height="23" background="backimg/tabs_m_tile.gif">&gt;&gt; ��ӭ����<%=request.cookies("renwen")("user")%>��<a href="Login.asp">�ص�¼</a> | <a href="usermanager.asp">�û����</a> 
            | <a href="#">����</a> | <a href="#">����</a> | <a href="#" onClick="location.reload();">ˢ��</a> 
            | <a href="#">����</a> | <a href="listalluser.asp">�û��б�</a> | <a href="look-ip.asp">�鿴IP��Դ</a> 
            <% if master then
			response.Write(" | <a href=admin_login.asp>����</a> ")
			end if %> 
            | <a href="Login.asp?userlogout=loginout"></a><a href="Login.asp?userlogout=loginout">�˳���̳</a>&nbsp; 
            &nbsp;&nbsp;<img src="backimg/msg_no_new_bar.gif" width="16" height="16"><a href="usermanager.asp?p=usersms">����Ϣ</a><font color="#FF0000"> 
            <%=notread%></font>(��)</td>
        </tr>
      </table>
      <%else%>
      <table width="100%" border=0 cellpadding=0 cellspacing=1>
        <tr> 
          <td height="23" background="backimg/tabs_m_tile.gif">&gt;&gt; ��ӭ����<b>����</b>�� 
            <a href="Login.asp">��¼</a> | <a title=ע���˲��ܷ�������Ŷ�� 
            href="regedit.asp">ע��</a> | <a href="lostpassport.asp">��������</a> | 
            <a href="#">����</a> | <a href="#">����</a> | <a href="#" onClick="location.reload();">ˢ��</a> 
            | <a href="listalluser.asp">�û��б�</a> |<a href="look-ip.asp"> �鿴IP��Դ</a> 
            | ����</td>
        </tr>
      </table>
      <%
end if
end sub
%>
    </td>
  </tr>
</table>
<table width="898" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr> 
    <td>
          </td>
  </tr>
  <tr> 
    <td height="13" valign="middle" >
	 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
        <tr> 
          <td height="28" bgcolor="#EAEAEA"><img src="backimg/Forum_nav.gif" width="9" height="9">&nbsp; 
            <%response.Write(thispage_name(request.ServerVariables("SCRIPT_NAME"),"url"))%> </td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td > 
      <%if not userislogin then%> 
        <table border="0" cellspacing="0">
		<form name="formlogin" method="post" action="Login.asp">
          <tr valign="bottom"> 
            <td width="57" align="right"> �û�����</td>
            <td width="75"> <input name="username" type="text" id="username" size="12"></td>
            <td width="37">����:</td>
            <td width="96"><font color="#000033"> 
              <input name="pwd" type="password" id="pws" size="12">
              </font> </td>
            <td width="77"><font color="#000033"> 
              <select name="cookies" id="cookies">
                <option value="0">������</option>
                <option value="1">һ��</option>
                <option value="2">һ����</option>
                <option value="3">һ��</option>
              </select>
              </font></td>
            <td width="106"> <input name="formlogin" type="submit" id="formlogin" value=" �� ¼ "> 
            </td>
          </tr>  </form>
        </table>
      <%end if%>
    </td>
  </tr>
</table>
<%
'��̳�����˵��ӳ���
sub listchangeLt()
dim rs,sqltext,i,tempout
set rs=server.createobject("adodb.recordset")
sqltext="select ltname,ifoff,id from ltname order by b_order,id desc"
rs.open sqltext,conn,1,1
tempout="<select onChange=""location.href=this.value"">"
tempout=tempout&"<option selected>��ת��̳��...</option>"
if not rs.eof then
for i=1 to rs.recordcount
tempout=tempout&"<option value='index.asp?atlb="&rs("id")&"'>��"&rs("ltname")&"</option>"&returnzlt(rs("id"))
rs.movenext
next
tempout=tempout&"</select>"
response.Write(tempout)
end if 
end sub
'Ϊ�����ӳ������ĺ���
function returnzlt(zlt)
dim rs,sqltext,tempout,i
set rs=server.createobject("adodb.recordset")
sqltext="select ltlb,id from ltlb where atlb='"&zlt&"' order by b_order desc,id desc"
rs.open sqltext,conn,1,1
if not rs.eof then
for i=1 to rs.recordcount
tempout=tempout&"<option value='list.asp?lb="&rs("id")&"'> ��"&rs("ltlb")&"</option>"
rs.movenext
next
end if
returnzlt=tempout
end function

'{�û�����ˢ�³���
call onlineUpdate()
function usersysinfo(info,getinfo)
if instr(info,";")>0 then
	dim usersys
	usersys=split(info,";")
	usersys(1)=replace(usersys(1),"MSIE","Internet Explorer")
	usersys(2)=replace(usersys(2),")","")
	usersys(2)=replace(usersys(2),"NT 5.1","XP")
	usersys(2)=replace(usersys(2),"NT 5.0","2000")
	usersys(2)=replace(usersys(2),"9x","Me")
	usersys(1)="� �� ����" & Trim(usersys(1))
	usersys(2)="����ϵͳ��" & Trim(usersys(2))
	if getinfo=1 then
		usersysinfo=usersys(1)
	else
		usersysinfo=usersys(2)
	end if
else
	if getinfo=1 then
		usersysinfo="δ֪"
	else
		usersysinfo="δ֪"
	end if
end if
end function

sub onlineUpdate()
dim ip,userip,userip2,lst,rs,sqltext,user,userjibie,usersys,userie,userisinonline,usercookieid
user=loginuser
usersys=usersysinfo(Request.ServerVariables("HTTP_USER_AGENT"),2)
userie=usersysinfo(Request.ServerVariables("HTTP_USER_AGENT"),1)
userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
userip2 = Request.ServerVariables("REMOTE_ADDR")
	if userip = ""  then
		ip=userip2
	else
		ip=userip
	end if

if not userislogin then
user="����"
userjibie="�ο�"
else
set rs=conn.execute("select jibie,user from userinfo where user='"&user&"'")
userjibie=rs("jibie")
    set rs=conn.execute("select id,user from online where user='"&user&"'")
    if (not rs.eof) and (not rs.bof) then
    response.Cookies("rewin")("userid")=rs("id")
    end if
end if

usercookieid=request.Cookies("rewin")("userid")
if usercookieid<>"" then 
set rs=conn.execute("select user,id from online where id="&usercookieid)
 if rs.eof or rs.bof then
 userisinonline=false
 else
 userisinonline=true
 end if
end if
if userisinonline=false then
conn.execute("insert into online([ip],[address],[where],[jibie],[user],[ie],[sys],[time],[hdtime]) values ('"&ip&"','"&ip&"','"&mythispage&"','"&userjibie&"','"&user&"','"&userie&"','"&usersys&"','"&now()&"','"&now()&"')")
set rs=conn.execute("select top 1 id from online order by id desc")
response.Cookies("rewin")("userid")=rs("id")
else
conn.execute("update online set [ip]='"&ip&"',[address]='"&ip&"',[where]='"&mythispage&"',[jibie]='"&userjibie&"',[user]='"&user&"',[ie]='"&userie&"',[sys]='"&usersys&"',[hdtime]=now() where id="&usercookieid)
end if
'ɾ����ʱ���û�
lst=5
conn.execute("delete from online where datediff('n',hdtime,Now())>"&lst)
end sub

sub onlineusersum(th)
dim sql,rs,rs1,honline,htime,djzx,rs2
set rs1=conn.execute("select honline,htime from admin_copyc")
honline=rs1("honline")
htime=rs1("htime")
set rs=server.CreateObject("adodb.recordset")
sql="select user from online"
rs.open sql,conn,1,1
djzx=rs.RecordCount
if honline<djzx then
conn.execute("update admin_copyc set honline="&djzx&",htime='"&now()&"' where id=1")
honline=djzx
htime=now()
end if
set rs2=server.CreateObject("adodb.recordset")
sql="select jibie from online where jibie='�ο�'"
rs2.open sql,conn,1,1
if th="all" then
response.Write("��ǰ�����û�:<font color=red><b>"&djzx&"</b></font>��,����ע���û�:<b>"&(djzx-rs2.recordcount)&"</b>��,����:<b>"&rs2.recordcount&"</b>��<br>")
response.Write("������߼�¼ <b>"&honline&" </b>��ͬʱ���ߣ�����ʱ��:"&htime)
else
response.Write("�����ߣ�<font color=red><b>"&djzx&"</b></font>��")
end if
end sub
%>