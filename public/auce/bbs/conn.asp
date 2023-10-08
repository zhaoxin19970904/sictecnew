<!--#include file="connect.asp" -->
<%
call checkIfOff()
sub checkIfOff() '检查论是否关闭
dim rs
set rs=conn.execute("select offlt,offtitle from admin_copyc")
if rs("offlt")=1 then
response.Write(rs("offtitle"))
response.End()
end if
end sub
'errmess 错误信息字串 userislogin 用户是否登录 master用户是否是管理员 userloginlock用户帐号是否被锁定 loginuser 登录用户名
dim founderr,errmess,userislogin,master,rsa,userloginlock,loginuser
dim bbs_seting '论坛设置
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
if rsa("jibie")<>"普通用户" then master=true end if
end if
userloginlock=rsa("kill")
end if

'版权信息子程序和论坛上部HTML,参数可调出论坛名称		 
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
'表格外框,参数分别控制是否显示,和上方下方
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
 errmess=errmess+"<li>对不起!没有找到你所请求的论坛,可能此论坛已经被删除!或者你没有从正确链接进入.<li><a href=index.asp>返回首页</a>"
 call founderror(errmess)
 end if
indexurl=indexurl&" → <a href=index.asp?atlb="&rs("id")&">"&rs("ltname")&"</a> → 论坛列表"
else
indexurl=indexurl&" → 论坛首页"
end if
end function

function ltlbname(lb,url)
dim rs,rs1
set rs=conn.execute("select ltlb,id,atlb,ifoff,ltconfig from ltlb where id="&cint(lb))
if rs.eof or rs.bof then
errmess=errmess+"<li>对不起!没有找到你所请求论坛,可能此论坛已经被删除!或者你没有从正确链接进入.<li><a href=index.asp>返回首页</a>"
call founderror(errmess)
else
  if rs("ifoff")=1 and (not master) then
  errmess=errmess+"<li>对不起!此论坛暂时关闭...<li>原因:由于管理员正在对论坛进行维护.<li>请访问别的论坛 <a href=index.asp>返回首页</a>"
  response.cookies("rewin")("errmess")=errmess
  response.Redirect("error.asp?founderr=mess")
  end if
end if
bbs_seting=split(rs("ltconfig"),"|")  
set rs1=conn.execute("select ltname,id from ltname where id="&cint(rs("atlb")))
if url="url" then
ltlbname=loadcopyc("ltname")&"-"&rs1("ltname")&"-"&rs("ltlb")
else
ltlbname="<a href=index.asp>"&loadcopyc("ltname")&"</a> → <a href=index.asp?atlb="&rs1("id")&">"&rs1("ltname")&"</a>"
ltlbname=ltlbname&" → <a href=list.asp?lb="&rs("id")&">"&rs("ltlb")&"</a>"
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
errmess=errmess+"<li>对不起!没有找到你所请求的页面,可能此帖已经被删除!或者你没有从正确链接进入.<li><a href=index.asp>返回首页</a>"
call founderror(errmess)
end if
if url="url" then
type_name=ltlbname(rs("lb"),"url")
else
type_name=ltlbname(rs("lb"),"nourl")&" → "&ChkBadWords(rs("name"))
end if
end function

'确定当然页名称子程序
function thispage_name(scriptname,url)
dim pagename,casename,pageurl,thispage,htmltitle,thetemp
htmltitle=loadcopyc("ltname")
pagename=split(scriptname,"/")
casename=pagename(ubound(pagename))
select case(casename)
  case "index.asp"
      thispage=indexurl()
      if request.QueryString("atlb")<>"" then
      htmltitle=htmltitle&"-论坛列表"
      else
      htmltitle=htmltitle&"-论坛首页"
      end if
case "list.asp"
 if request.QueryString("lb")="" then
 response.Redirect("error.asp?founderr=link")
 end if
thispage=ltlbname(request.QueryString("lb"),"nourl")&" → 帖子列表"
htmltitle=ltlbname(request.QueryString("lb"),"url")&" → 帖子列表"
case "type.asp"
 if request.QueryString("id")="" then
 response.Redirect("error.asp?founderr=link")
 end if
thispage=type_name(request.QueryString("id"),"nourl")
htmltitle=type_name(request.QueryString("id"),"url")&"-浏览帖子"
 if cint(bbs_seting(0))=1 and (not userislogin) then
 errmess="<li>产生错误!<li>原因:此论坛需要注册会员才可能浏览,你现在还不是注册会员,<a href=regedit.asp>请注册</a>"
 call founderror(errmess)
 end if
case "AddIn.asp"
 if request.QueryString("lb")="" or request.QueryString("czlb")="" then
 response.Redirect("error.asp?founderr=link")
 end if
thispage=ltlbname(request.QueryString("lb"),"nourl")&" → "& addin_czlb(request.QueryString("czlb"))
htmltitle=ltlbname(request.QueryString("lb"),"url")&"-"&addin_czlb(request.QueryString("czlb"))
 if cint(bbs_seting(1))=1 then
 errmess="<li>产生错误!<li>论坛被管理员锁定,只可浏览,不可发帖"
 call founderror(errmess)
 end if
case "Login.asp"
thispage=indexurl()&" → <a href=Login.asp>用户登录</a>"
htmltitle=htmltitle&"-登录页面"
case "listalluser.asp"
thispage=indexurl()&" → <a href=listalluser.asp>用户列表</a>"
htmltitle=htmltitle&"-用户列表"
case "regedit.asp"
thispage=indexurl()&" → <a href=regedit.asp>用户注册</a>"
htmltitle=htmltitle&"-用户注册"
case "tpaddIn.asp"
 if request.QueryString("lb")="" then
 response.Redirect("error.asp?founderr=link")
 end if
thispage=ltlbname(request.QueryString("lb"),"nourl")&" → 发起投票"
htmltitle=ltlbname(request.QueryString("lb"),"url")&"-发起投票"
 if cint(bbs_seting(1))=1 then
 errmess="<li>产生错误!<li>论坛被管理员锁定,只可浏览,不可发帖"
 call founderror(errmess)
 end if
case "error.asp"
if request.QueryString("lb")<>"" then
thispage=ltlbname(request.QueryString("lb"),"nourl")&" → 信息页面"
htmltitle=ltlbname(request.QueryString("lb"),"url")&"-信息页面"
else
thispage="<a href=index.asp>"&loadcopyc("ltname")&"</a> → 信息页面"
htmltitle=htmltitle&"-信息页面"
end if
case "usermanager.asp"
thetemp=usermanager_zclb(request.QueryString("p"))
thispage="<a href=index.asp>"&loadcopyc("ltname")&"</a> → <a href=usermanager.asp>控制面板</a> → "&thetemp
htmltitle=htmltitle&"- 控制面板 - "&thetemp
case "changpsw.asp"
thispage="<a href=index.asp>"&loadcopyc("ltname")&"</a> → <a href=usermanager.asp>控制面板</a> → 修改密码"
htmltitle=htmltitle&"-控制面板-修改密码"
case "egedit.asp"
thispage="<a href=index.asp>"&loadcopyc("ltname")&"</a> → <a href=usermanager.asp>控制面板</a> → 资料修改"
htmltitle=htmltitle&"-控制面板-资料修改"
case "usertype.asp"
thispage="<a href=index.asp>"&loadcopyc("ltname")&"</a> →  用户信息  → 基本信息"
htmltitle=htmltitle&"-用户信息-基本资料"
case "caozuo.asp"
if request.QueryString("lb")="" then
thispage="<a href=index.asp>"&loadcopyc("ltname")&"</a> → 管理操作"
htmltitle=loadcopyc("ltname")&"-管理操作"
else
thispage=ltlbname(request.QueryString("lb"),"nourl")&" → 管理操作"
htmltitle=ltlbname(request.QueryString("lb"),"url")&"-管理操作"
end if
case "look-ip.asp"
 if request.QueryString("id")="" then
 thispage="<a href=index.asp>"&loadcopyc("ltname")&"</a>  → 查看IP来源"
 htmltitle=loadcopyc("ltname")&"-查看IP来源"
 else
 thispage=type_name(request.QueryString("id"),"nourl")&" → 查看IP来源"
 htmltitle=type_name(request.QueryString("id"),"url")&"-查看IP来源"
 end if
case "bbs_log.asp"
 thispage="<a href=index.asp>"&loadcopyc("ltname")&"</a>  → 论坛日志"
 htmltitle=loadcopyc("ltname")&"-论坛日志"
case "helpubb.asp"
 thispage="<a href=index.asp>"&loadcopyc("ltname")&"</a>  → UBB帮助"
 htmltitle=loadcopyc("ltname")&"-UBB帮助"
case "lostpassport.asp"
 thispage="<a href=index.asp>"&loadcopyc("ltname")&"</a>  → 找回密码"
 htmltitle=loadcopyc("ltname")&"-找回密码"
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
usermanager_zclbx="发表过主题"
	 case "usersms"
usermanager_zclbx="短信服务"	 
	case "addfriend"
usermanager_zclbx="添加好友"
    case "editfriend"
usermanager_zclbx="编辑好友"	 
   case "userf"
usermanager_zclbx="用户收藏"
   case else
usermanager_zclbx="面板首页"
end select
usermanager_zclb=usermanager_zclbx
end function

function addin_czlb(czlb)
dim  addin_czlbx
 select case czlb
 case "addtz"
 addin_czlbx="发新帖子"
 case "edittz"
 addin_czlbx="修改帖子"
 case "retz"
 addin_czlbx="回复帖子"
 case "editretz"
 addin_czlbx="回复修改"
 case "vote"
 addin_czlbx="引用帖子"
 case "editvote"
 addin_czlbx="引用修改"
 case "revote"
 addin_czlbx="引用回复"
 case "ok"
 addin_czlbx="发帖成功"
 end select
 addin_czlb=addin_czlbx
end function

sub updateftuser(ifzt) '用户发帖成功,更新数据,ifzt参数为 区分发主题,回复
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

sub founderror(themess) '论坛信息提示子程序
response.cookies("rewin")("errmess")=themess
response.Redirect("error.asp?founderr=mess&lb="&request.QueryString("lb"))
end sub
'论坛提示信息,BadWords论坛过虑字符
dim error1,error2,BadWords
error1="<li>对不起!我们发现你现在还没登录!所以无法完成你的操作请求!请<a href=Login.asp>登录</a>"
error2="<li>对不起!我们发现你的ID被封闭!<li>原因:可能你违反有关规定<li>所以无法完成你的操作请求!请<a href=index.asp>返回首页</a>"
BadWords=loadcopyc("badwords")
%>