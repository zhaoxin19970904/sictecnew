<!--#include file="connect.asp" -->
<%
call checkIfOff()
sub checkIfOff() '������Ƿ�ر�
dim rs
set rs=conn.execute("select offlt,offtitle from admin_copyc")
if rs("offlt")=1 then
response.Write(rs("offtitle"))
response.End()
end if
end sub
'errmess ������Ϣ�ִ� userislogin �û��Ƿ��¼ master�û��Ƿ��ǹ���Ա userloginlock�û��ʺ��Ƿ����� loginuser ��¼�û���
dim founderr,errmess,userislogin,master,rsa,userloginlock,loginuser
dim bbs_seting '��̳����
userislogin=false

if request.cookies("renwen")("user")<>"" and request.cookies("renwen")("passedok")="ofdkjduy" then 
userislogin=true
loginuser=request.cookies("renwen")("user")
end if
if userislogin then
set rsa=conn.execute("select user,jibie,kill from userinfo where user='"&loginuser&"'")
if rsa.eof or rsa.bof then
userislogin=false
else
if rsa("jibie")<>"��ͨ�û�" then master=true end if
end if
userloginlock=rsa("kill")
end if

'��Ȩ��Ϣ�ӳ������̳�ϲ�HTML,�����ɵ�����̳����		 
		  function  loadcopyc(lbhtm)
		  dim rs,sql
		  Set rs = Server.CreateObject("ADODB.Recordset")
          sql="select id,copyc,ltname,tophtm,badwords,reguser from admin_copyc"
          rs.open sql,conn,1,1
		  select case(lbhtm)
		  case "copyc"
		  loadcopyc=("<div align=center><html>"&rs("copyc")&"</html>")
		  case "tophtm"
		  loadcopyc=("<html>"&rs("tophtm")&"</html>")
		  case "ltname"
		  loadcopyc=(rs("ltname"))
		  case "badwords"
		  loadcopyc=rs("badwords")
		  case "reguser"
		  loadcopyc=rs("reguser")
		  end select
		  end function 
'������,�����ֱ�����Ƿ���ʾ,���Ϸ��·�
sub tabletop(ifhidde,where)
if ifhidde=true then
select case(where)
case "top"
response.Write("<TABLE width=898 border=0 align=center cellPadding=0 cellSpacing=0 borderColor=#92b9fb>")
 response.Write(" <TR> ")
    response.Write("<TD width=28 height=28><IMG src=backimg/a1.gif></TD>")
    response.Write("<TD width=625 background=backimg/a2(1).gif height=28>&nbsp;</TD>")
   response.Write(" <TD width=19><IMG src=backimg/a3.gif></TD>")
   response.Write("<TD width=296><IMG src=backimg/a4.gif></TD>")
  response.Write("</TR>")
   response.Write(" </TABLE>")
case"down"
response.Write("<TABLE width=898 border=0 align=center cellPadding=0 cellSpacing=0 borderColor=#92b9fb >")
response.Write("<TR>") 
response.Write("<TD width=100><IMG src=backimg/a6.gif width=100 height=23></TD>")
response.Write("<TD width=768 background=backimg/a7.gif></TD>")
response.Write("<TD width=100><IMG src=backimg/a8.gif></TD>")
response.Write("</TR>")
response.Write("</TABLE>")
end select
end if
end sub

function indexurl()
dim rs
indexurl="<a href=index.asp>"&loadcopyc("ltname")&"</a>"
if request.QueryString("atlb")<>"" then 
set rs=conn.execute("select ltname,id from ltname where id="&cint(request.QueryString("atlb")))
 if rs.eof or rs.bof then
 errmess=errmess+"<li>�Բ���!û���ҵ������������̳,���ܴ���̳�Ѿ���ɾ��!������û�д���ȷ���ӽ���.<li><a href=index.asp>������ҳ</a>"
 call founderror(errmess)
 end if
indexurl=indexurl&" �� <a href=index.asp?atlb="&rs("id")&">"&rs("ltname")&"</a> �� ��̳�б�"
else
indexurl=indexurl&" �� ��̳��ҳ"
end if
end function

function ltlbname(lb,url)
dim rs,rs1
set rs=conn.execute("select ltlb,id,atlb,ifoff,ltconfig from ltlb where id="&cint(lb))
if rs.eof or rs.bof then
errmess=errmess+"<li>�Բ���!û���ҵ�����������̳,���ܴ���̳�Ѿ���ɾ��!������û�д���ȷ���ӽ���.<li><a href=index.asp>������ҳ</a>"
call founderror(errmess)
else
  if rs("ifoff")=1 and (not master) then
  errmess=errmess+"<li>�Բ���!����̳��ʱ�ر�...<li>ԭ��:���ڹ���Ա���ڶ���̳����ά��.<li>����ʱ����̳ <a href=index.asp>������ҳ</a>"
  response.cookies("rewin")("errmess")=errmess
  response.Redirect("error.asp?founderr=mess")
  end if
end if
bbs_seting=split(rs("ltconfig"),"|")  
set rs1=conn.execute("select ltname,id from ltname where id="&cint(rs("atlb")))
if url="url" then
ltlbname=loadcopyc("ltname")&"-"&rs1("ltname")&"-"&rs("ltlb")
else
ltlbname="<a href=index.asp>"&loadcopyc("ltname")&"</a> �� <a href=index.asp?atlb="&rs1("id")&">"&rs1("ltname")&"</a>"
ltlbname=ltlbname&" �� <a href=list.asp?lb="&rs("id")&">"&rs("ltlb")&"</a>"
end if
end function

function ltttname(lb)
dim rs
set rs=conn.execute("select ltname,id from ltname where id="&cint(lb))
ltttname=rs("ltname")
end function
function zlbltname(lb)
dim rs
set rs=conn.execute("select ltlb,id from ltlb where id="&cint(lb))
zlbltname=rs("ltlb")
end function

function type_name(id,url)
dim rs
set rs=conn.execute("select lb,name from borecorder where id="&cint(id))
if rs.eof or rs.bof then
errmess=errmess+"<li>�Բ���!û���ҵ����������ҳ��,���ܴ����Ѿ���ɾ��!������û�д���ȷ���ӽ���.<li><a href=index.asp>������ҳ</a>"
call founderror(errmess)
end if
if url="url" then
type_name=ltlbname(rs("lb"),"url")
else
type_name=ltlbname(rs("lb"),"nourl")&" �� "&ChkBadWords(rs("name"))
end if
end function

'ȷ����Ȼҳ�����ӳ���
function thispage_name(scriptname,url)
dim pagename,casename,pageurl,thispage,htmltitle,thetemp
htmltitle=loadcopyc("ltname")
pagename=split(scriptname,"/")
casename=pagename(ubound(pagename))
select case(casename)
  case "index.asp"
      thispage=indexurl()
      if request.QueryString("atlb")<>"" then
      htmltitle=htmltitle&"-��̳�б�"
      else
      htmltitle=htmltitle&"-��̳��ҳ"
      end if
case "list.asp"
 if request.QueryString("lb")="" then
 response.Redirect("error.asp?founderr=link")
 end if
thispage=ltlbname(request.QueryString("lb"),"nourl")&" �� �����б�"
htmltitle=ltlbname(request.QueryString("lb"),"url")&" �� �����б�"
case "type.asp"
 if request.QueryString("id")="" then
 response.Redirect("error.asp?founderr=link")
 end if
thispage=type_name(request.QueryString("id"),"nourl")
htmltitle=type_name(request.QueryString("id"),"url")&"-�������"
 if cint(bbs_seting(0))=1 and (not userislogin) then
 errmess="<li>��������!<li>ԭ��:����̳��Ҫע���Ա�ſ������,�����ڻ�����ע���Ա,<a href=regedit.asp>��ע��</a>"
 call founderror(errmess)
 end if
case "AddIn.asp"
 if request.QueryString("lb")="" or request.QueryString("czlb")="" then
 response.Redirect("error.asp?founderr=link")
 end if
thispage=ltlbname(request.QueryString("lb"),"nourl")&" �� "& addin_czlb(request.QueryString("czlb"))
htmltitle=ltlbname(request.QueryString("lb"),"url")&"-"&addin_czlb(request.QueryString("czlb"))
 if cint(bbs_seting(1))=1 then
 errmess="<li>��������!<li>��̳������Ա����,ֻ�����,���ɷ���"
 call founderror(errmess)
 end if
case "Login.asp"
thispage=indexurl()&" �� <a href=Login.asp>�û���¼</a>"
htmltitle=htmltitle&"-��¼ҳ��"
case "listalluser.asp"
thispage=indexurl()&" �� <a href=listalluser.asp>�û��б�</a>"
htmltitle=htmltitle&"-�û��б�"
case "regedit.asp"
thispage=indexurl()&" �� <a href=regedit.asp>�û�ע��</a>"
htmltitle=htmltitle&"-�û�ע��"
case "tpaddIn.asp"
 if request.QueryString("lb")="" then
 response.Redirect("error.asp?founderr=link")
 end if
thispage=ltlbname(request.QueryString("lb"),"nourl")&" �� ����ͶƱ"
htmltitle=ltlbname(request.QueryString("lb"),"url")&"-����ͶƱ"
 if cint(bbs_seting(1))=1 then
 errmess="<li>��������!<li>��̳������Ա����,ֻ�����,���ɷ���"
 call founderror(errmess)
 end if
case "error.asp"
if request.QueryString("lb")<>"" then
thispage=ltlbname(request.QueryString("lb"),"nourl")&" �� ��Ϣҳ��"
htmltitle=ltlbname(request.QueryString("lb"),"url")&"-��Ϣҳ��"
else
thispage="<a href=index.asp>"&loadcopyc("ltname")&"</a> �� ��Ϣҳ��"
htmltitle=htmltitle&"-��Ϣҳ��"
end if
case "usermanager.asp"
thetemp=usermanager_zclb(request.QueryString("p"))
thispage="<a href=index.asp>"&loadcopyc("ltname")&"</a> �� <a href=usermanager.asp>�������</a> �� "&thetemp
htmltitle=htmltitle&"- ������� - "&thetemp
case "changpsw.asp"
thispage="<a href=index.asp>"&loadcopyc("ltname")&"</a> �� <a href=usermanager.asp>�������</a> �� �޸�����"
htmltitle=htmltitle&"-�������-�޸�����"
case "egedit.asp"
thispage="<a href=index.asp>"&loadcopyc("ltname")&"</a> �� <a href=usermanager.asp>�������</a> �� �����޸�"
htmltitle=htmltitle&"-�������-�����޸�"
case "usertype.asp"
thispage="<a href=index.asp>"&loadcopyc("ltname")&"</a> ��  �û���Ϣ  �� ������Ϣ"
htmltitle=htmltitle&"-�û���Ϣ-��������"
case "caozuo.asp"
if request.QueryString("lb")="" then
thispage="<a href=index.asp>"&loadcopyc("ltname")&"</a> �� �������"
htmltitle=loadcopyc("ltname")&"-�������"
else
thispage=ltlbname(request.QueryString("lb"),"nourl")&" �� �������"
htmltitle=ltlbname(request.QueryString("lb"),"url")&"-�������"
end if
case "look-ip.asp"
 if request.QueryString("id")="" then
 thispage="<a href=index.asp>"&loadcopyc("ltname")&"</a>  �� �鿴IP��Դ"
 htmltitle=loadcopyc("ltname")&"-�鿴IP��Դ"
 else
 thispage=type_name(request.QueryString("id"),"nourl")&" �� �鿴IP��Դ"
 htmltitle=type_name(request.QueryString("id"),"url")&"-�鿴IP��Դ"
 end if
case "bbs_log.asp"
 thispage="<a href=index.asp>"&loadcopyc("ltname")&"</a>  �� ��̳��־"
 htmltitle=loadcopyc("ltname")&"-��̳��־"
case "helpubb.asp"
 thispage="<a href=index.asp>"&loadcopyc("ltname")&"</a>  �� UBB����"
 htmltitle=loadcopyc("ltname")&"-UBB����"
case "lostpassport.asp"
 thispage="<a href=index.asp>"&loadcopyc("ltname")&"</a>  �� �һ�����"
 htmltitle=loadcopyc("ltname")&"-�һ�����"
end select
if url="url" then
thispage_name=thispage
else
thispage_name=htmltitle
end if
end function

function usermanager_zclb(czlb)
dim usermanager_zclbx
select case czlb
	 case "userftz"  
usermanager_zclbx="���������"
	 case "usersms"
usermanager_zclbx="���ŷ���"	 
	case "addfriend"
usermanager_zclbx="��Ӻ���"
    case "editfriend"
usermanager_zclbx="�༭����"	 
   case "userf"
usermanager_zclbx="�û��ղ�"
   case else
usermanager_zclbx="�����ҳ"
end select
usermanager_zclb=usermanager_zclbx
end function

function addin_czlb(czlb)
dim  addin_czlbx
 select case czlb
 case "addtz"
 addin_czlbx="��������"
 case "edittz"
 addin_czlbx="�޸�����"
 case "retz"
 addin_czlbx="�ظ�����"
 case "editretz"
 addin_czlbx="�ظ��޸�"
 case "vote"
 addin_czlbx="��������"
 case "editvote"
 addin_czlbx="�����޸�"
 case "revote"
 addin_czlbx="���ûظ�"
 case "ok"
 addin_czlbx="�����ɹ�"
 end select
 addin_czlb=addin_czlbx
end function

sub updateftuser(ifzt) '�û������ɹ�,��������,ifzt����Ϊ ���ַ�����,�ظ�
dim rs,tzxlb
tzxlb=request.Form("lb")
if ifzt then
conn.execute("update userinfo set ftz=ftz+1,userjf=userjf+10 where user='"&loginuser&"'")
set rs=conn.execute("select todayft,today from ltlb where id="&tzxlb)
  if rs("today")<>date() then
  conn.execute("update ltlb set yestdayft=todayft,today=date() ,todayft=0")
  end if
conn.execute("update ltlb set todayft=todayft+1 where id="&tzxlb)
conn.execute("update ltlb set tzsum=tzsum+1 where id="&tzxlb)
else
conn.execute("update borecorder set rely=rely+1,retime='"&now()&"',lastly='"&loginuser&"' where id="&request.Form("id"))
conn.execute("update ltlb set retzsum=retzsum+1 where id="&tzxlb)
end if
end sub

sub founderror(themess) '��̳��Ϣ��ʾ�ӳ���
response.cookies("rewin")("errmess")=themess
response.Redirect("error.asp?founderr=mess&lb="&request.QueryString("lb"))
end sub
'��̳��ʾ��Ϣ,BadWords��̳�����ַ�
dim error1,error2,BadWords
error1="<li>�Բ���!���Ƿ��������ڻ�û��¼!�����޷������Ĳ�������!��<a href=Login.asp>��¼</a>"
error2="<li>�Բ���!���Ƿ������ID�����!<li>ԭ��:������Υ���йع涨<li>�����޷������Ĳ�������!��<a href=index.asp>������ҳ</a>"
BadWords=loadcopyc("badwords")
%>