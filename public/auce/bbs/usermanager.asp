<!--#include file="CONN.ASP" -->
<!--#include file="ubbcode.asp" -->
<%
if not userislogin then
errmess=errmess&error1
call founderror(errmess)
end if
%>
<!--#include file="mymem.asp" -->
<!--#include file="mymeumu.asp" -->
<table width="898" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td colspan="2">
<%call tabletop(true,"top")%></td>
  </tr>
  <tr> 
    <td height="27" colspan="2" background="backimg/bg1.gif"><strong><font color="#FFFFFF"> 
      &nbsp;&nbsp;<%=right(mythispage,5)%></font></strong></td>
  </tr>
  <tr> 
    <td colspan="2" valign="top"> 
      <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#92b9fb">
	  <%
	  select case (request.QueryString("p"))
	 case "userftz"  
	      response.Write("<tr bgcolor=""#FFFFFF""><td>")
          call listzjftzx(100,"fz")
          response.Write("</tr></td>")
	 case "usersms"
	      response.Write("<tr bgcolor=""#FFFFFF""><td>")
          call listusershormess(100) 
          response.Write("</tr></td>")	 
	case "addfriend"
		   response.Write("<tr bgcolor=""#FFFFFF""><td>")
          call addfriend()
          response.Write("</tr></td>")	 
    case "editfriend"
		 response.Write("<tr bgcolor=""#FFFFFF""><td>")
          call userfriendlist(10)
          response.Write("</tr></td>")	 
   case "userf"
   call listzjftzx(100,"userf")
   case "userpic"
   call userpiclist()	
   case else
	  %>
        <tr bgcolor="#FFFFFF"> 
          <td> <table width="100%" border="0">
              <tr> 
                <td width="25%" height="198" valign="top">
	<%
		  call listuserjbmess()
		  sub listuserjbmess()
		  dim rs
		  set rs=conn.execute("select * from userinfo where user='"&loginuser&"'")
		  %> <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#92b9fb">
                    <tr bgcolor="#FFFFFF"> 
                      <td colspan="2">用户头像</td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td colspan="2"><img src="../bbs/<%=rs("myface")%>" width="<%=rs("width")%>" height="<%=rs("height")%>" id=idface></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="42%">级别</td>
                      <td width="58%"><%=rs("jibie")%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="42%">用户名</td>
                      <td><%=rs("user")%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="42%">性 别</td>
                      <td><%=rs("xb")%>&nbsp;</td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="42%">出 生</td>
                      <td><%=rs("birthday")%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="42%">E-mail</td>
                      <td> <%if rs("setuser1")=0 and rs("user")<>loginuser then response.write "用户未开放" else response.write rs("email") end if%> </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="42%">oicq</td>
                      <td><%=rs("oicq")%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="42%">帖子总数</td>
                      <td><%=rs("ftz")%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="42%">被删主题</td>
                      <td><%=rs("deltz")%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="42%">登陆次数</td>
                      <td><%=rs("userlog")%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="42%">正常退出:</td>
                      <td><%=rs("logout")%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="42%">现在积分</td>
                      <td><%=rs("userjf")%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="42%">联系方式</td>
                      <td> <%if rs("setuser1")=0 and rs("user")<>loginuser then response.write "用户未开放" else response.write rs("conn") end if%> </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="42%">在线时长</td>
                      <td> <%=rs("atime")%> 秒</td>
                    </tr>
                  </table>
                  <%end sub%> <br>
<%
call userfriendlist(10)
sub userfriendlist(listh)
response.Write("<table width=""100%"" cellpadding=""4"" cellspacing=""1"" bgcolor=""#92b9fb"">")
response.Write("<tr><td width=""60%"" height=""24"" background=""backimg/bg1.gif"">好友名字</td>") 
response.Write("<td width=""34%"" background=""backimg/bg1.gif""> 加入时间</td><td background=""backimg/bg1.gif""> </td></tr>")
dim rs,i,rs2
set rs=conn.execute("select id,fd,me,date from fd where me='"&loginuser&"' order by date")
 i=1
 if( not rs.eof )and (not rs.bof) then
  do while not rs.eof 
  response.Write("<tr><td bgcolor=""#FFFFFF""><img src=""PIC/msg_new_bar.gif"" width=""16"" height=""16""> <a href=usertype.asp?user="&rs("fd")&"> "&rs("fd")&"</a>")
  set rs2=conn.execute("select user from online where user='"&rs("fd")&"'")
  if not rs2.eof then
  response.Write(" [在线]")
  else
  response.Write(" [离线]")
  end if
  response.Write("</td><td bgcolor=""#FFFFFF"">"&rs("date")&"</td><td bgcolor=""#FFFFFF""><a href=foruser.asp?czlb=delfriend&id="&rs("id")&" title=删除当前好友 onclick=""{if(confirm('删除好友，确定删除吗?')){return true;}return false;}"">删</a></td></tr>")
   rs.movenext
   if i>10 then exit do
   i=i+1
   loop
else 
response.Write("<tr><td bgcolor=""#FFFFFF"" align=""right"" colspan=3>你还没好友</td></tr></table>")
end if
response.Write("<tr><td bgcolor=""#FFFFFF"" align=""right"" colspan=3><input type=""button"" name=""Submit"" onClick=""location.href='usermanager.asp?p=addfriend'"" value=""添加好友""></td></tr></table>")
end sub
%></td>
                <td width="75%" valign="top">
 <%
		 call listusershormess(10) 
		 sub listusershormess(s)
 %> <form name="form1" method="post" action="foruser.asp?czlb=delmess">
                  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#92b9fb">
                    <tr align="center"> 
                      <td colspan="5" bgcolor="#FFFFFF"><img src="PIC/m_inbox.gif" width="40" height="40"><img src="PIC/M_address.gif" width="40" height="40"> 
                        <img src="PIC/M_issend.gif" width="40" height="40"> <img src="PIC/M_outbox.gif" width="40" height="40"><img src="PIC/M_recycle.gif" width="40" height="40"><img src="PIC/m_write.gif" width="40" height="40"></td>
                    </tr>
                    <tr align="center" background="backimg/bg1.gif"> 
                      <td width="40" height="27" background="backimg/bg1.gif"> 
                        状态 </td>
                      <td width="274" height="27" background="backimg/bg1.gif"> 
                        短信标题</td>
                      <td width="77" height="27" background="backimg/bg1.gif"> 
                        发信人</td>
                      <td height="27" colspan="2" background="backimg/bg1.gif"> 
                        发信时间</td>
                    </tr>
                    <%
dim rs,i,nr,allmess
Set rs = conn.execute("select * from mess where fd='"&loginuser&"' order by id desc " )
i=0
nr=0
if (not rs.eof) and (not rs.bof )then
  do while not rs.eof 
%>
                    <tr bgcolor="#FFFFFF"> 
                      <td> <%
		 if rs("click")=0 then 
		 response.Write("<img src=""pic/m_news.gif"" width=""21"" height=""14"" alt=""未读"">" )
         nr=nr+1
		 else
		 response.Write("<img src=""pic/m_olds.gif"" width=""21"" height=""14"" alt=""已读"">")
		 end if
%> </td>
                      <td><a href="#" onClick="MM_openBrWindow('readmess.asp?rmid=<%=rs("id")%>','','scrollbars=yes,resizable=yes,width=500,height=450')"> 
                        <%if rs("click")<1 then response.write "<b>"&rs("name")&"</b>" else response.write rs("name") end if %>
                        </a></td>
                      <td><a href="usertype.asp?user=<%=rs("me")%>"><%=rs("me")%></a></td>
                      <td width="123"><%=rs("time")%></td>
                      <td width="16"><input name="pdel" type="checkbox" title="批量删除选项" value="<%=rs("id")%>" ></td>
                    </tr>
                    <%
   rs.movenext
   if s=10 then 
   if i>9 then exit do
   end if
   i=i+1
   loop
else
response.Write("<tr><td bgcolor=""#FFFFFF"" colspan=""5"" >你没短信息</td></tr>")
end if
%>
                    <tr align="right"> 
                        <td colspan="5" bgcolor="#FFFFFF"> 你一共有 <b><%=i%></b> 条短信,其中 <b><%=nr%></b> 条未读 
<script language="JavaScript">
function todel(){
var id=""
var x=<%=i%>
for (i=0;i<x;i++){
if ( document.form1.mplay[i].checked ==true){
id+=document.form1.mplay[i].value+','
}
}
openwindow(id)
}
function selectall()
{
var x=<%=i%>
for (i=0;i<x;i++)
{
document.form1.pdel[i].checked =true;
}
}
function delall(){
var id
    id=document.form1.pdel.value
window.location.href="del.asp?czlb=delp&id="+id;
}
</script>
<input type="button" name="Submit3" value="全选" onClick="selectall();" > 
                          <input type="submit" name="Submit2" onclick="{if(confirm('删除你所选定的短信息，确定删除吗?')){return true;}return false;}" value="删除所选">
                        </td>
                    </tr>
                  </table></form>
                  <%end sub%> <br> <br>
<%
call listzjftzx(10,"cy")
sub listzjftzx(totles,userf)
dim rs,i,rsf,userfid,lastcz
%> 

                  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#92b9fb">
                    <tr height="27" background="backimg/bg1.gif"> 
                      <td width="43%" height="27" background="backimg/bg1.gif"> 
                        <strong>主题</strong> </td>
                      <td width="12%" background="backimg/bg1.gif">作者</td>
                      <td width="12%" background="backimg/bg1.gif">人气/回复</td>
                      <td width="11%" background="backimg/bg1.gif">最后回复</td>
                      <td width="22%" background="backimg/bg1.gif">发帖时间</td>
		                </tr>
             <%
if userf="userf" then
set rsf=conn.execute("select tzid,user from userf where user='"&loginuser&"'")
do while not rsf.eof
userfid=rsf("tzid")&","&userfid
rsf.movenext
loop
if userfid="" then
 response.Write("<tr><td bgcolor=#FFFFFF colspan=""5"">你还没收藏过任何主题.</td></tr></table>") 
exit sub
end if
set rs=conn.execute("select id,name,time,user,retime,lastly,click,rely,lb from borecorder where id in("&userfid&") order by id desc",300,1)
else
set rs=conn.execute("select id,name,time,user,retime,lastly,click,rely,lb from borecorder where user='"&request.cookies("renwen")("user")&"' order by id desc",300,1)
end if
			 if rs.eof or rs.bof then 
			  response.Write("<tr><td bgcolor=#FFFFFF colspan=""5"">你还没有参与过任何主题</td></tr></table>") 
			 exit sub
			 end if
			  if totles=10 then 
			  do while not rs.eof
			  response.Write("<tr bgcolor=#FFFFFF><td>□ <a href=type.asp?id="&rs("id")&">"&dvHTMLEncode(rs("name"))&"</a></td><td align=center>"&rs("user")&"</td><td align=center>"&rs("click")&"/"&rs("rely")&"</td><td align=center>"&rs("lastly")&"</td><td align=center>"&rs("time")&"</td></tr>")
			  rs.movenext
			  i=i+1
			  if i>10 then exit do
			  loop
			  else
			  do while not rs.eof
			  response.Write("<tr bgcolor=#FFFFFF><td>□ <a href=type.asp?id="&rs("id")&">"&dvHTMLEncode(rs("name"))&"</a>")
        '当回复达到多页时显示
		dim crely,x,lb
		lb=rs("lb")
		crely=rs("rely")
		if crely>10 then
		dim disppagesd
		response.Write("[<b><img src=backimg/multipage.gif>")
		if CInt(crely/10+0.5)>8 then 
		disppagesd=8
		else
		disppagesd=CInt(crely/10+0.5)
		end if
        for x=1 to disppagesd
        Response.write(" <a href='type.asp?page="&x&"&lb="&lb&"&id="&rs("id")&"'><font color=red>"&x&"</font></a>")
        next
		if disppagesd=8 then 
		response.Write(" ... <a href='type.asp?page="&CInt(crely/10+0.5)&"&lb="&lb&"&id="&rs("id")&"'><font color=red>"&CInt(crely/10+0.5)&"</font></a>")
		end if
		response.Write("</b>]")
        end if
		 if userf="userf" then
		 lastcz="<td><a href=foruser.asp?czlb=deluserf&id="&rs("id")&" title=把这个帖子移出我的收藏>删</a></td>"
		 end if			  
 			  response.Write("</td><td align=center>"&rs("user")&"</td><td align=center>"&rs("click")&"/"&rs("rely")&"</td><td align=center>"&rs("lastly")&"</td><td align=center>"&rs("time")&"</td>"&lastcz&"</tr>")  
			  rs.movenext
			  i=i+1
			  if i>300 then exit do
			  loop
			 end if
%>
                  </table>
<%end sub%>
<%sub addfriend()%>
<form name="form2" method="post" action="foruser.asp?czlb=addfriend">
                  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#92b9fb">
                    <tr> 
                        <td height="27" background="backimg/bg1.gif">添加好友</td>
                    </tr>
                    <tr> 
                      <td height="30" bgcolor="#FFFFFF"><input name="friend" type="text" id="friend" value="<%=request.QueryString("fd")%>" size="30">
                        使用逗号（,）分开，最多5位用户 </td>
                    </tr>
                    <tr>
                        <td height="30"  bgcolor="#FFFFFF"> 
                          <input type="submit" name="Submit" value=" 添 加 ">
                        </td>
                      </tr></table></form>
			  <%end sub%>
                 <%
				 sub userpiclist()
				 dim rs
				 set rs=conn.execute("select * from upfile where user='"&loginuser&"' order by lb desc" )
				 %>
                  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#92b9fb">
                    <tr> 
                      <td width="10%" height="27" background="backimg/bg1.gif">文件类型</td>
                      <td width="35%" height="27" background="backimg/bg1.gif">文件名</td>
                      <td width="26%" height="27" background="backimg/bg1.gif">上传时间</td>
                      <td width="9%" height="27" background="backimg/bg1.gif">大小</td>
                      <td width="12%" background="backimg/bg1.gif">图片位置</td>
                      <td width="8%" background="backimg/bg1.gif">操作</td>
                    </tr>
                    <%
					do while not rs.eof
					response.Write("<tr bgcolor=#FFFFFF>")
                	response.Write("<td height=19>"&rs("F_type")&"</td><td><a href="&rs("pic")&">"&rs("pic")&"</a></td><td>"&rs("time")&"</td><td>"&rs("size")&"</td>")
					if rs("lb")="face" then
					response.Write("<td>头像</td>")
					else
					response.Write("<td>"&zlbltname(cint(rs("lb")))&"</td>")
					end if
					response.Write("<td><a href=foruser.asp?czlb=delfile&id="&rs("id")&" onclick=""{if(confirm('确实要删除此图吗，确定删除吗?')){return true;}return false;}"">删除</a></td></tr>")
					rs.movenext
					loop
					%>
                  </table>
				 <%end sub%></td>
              </tr>
            </table></td>
        </tr>
<%end select%>
      </table>
    </td>
  </tr>
  <tr>
    <td colspan="2">
<%call tabletop(true,"down")%></td>
  </tr>
</table>
<!--#include file="end.asp" -->