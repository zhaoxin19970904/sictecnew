<!--#include file="CONN.ASP" -->
<!--#include file="ubbcode.asp"-->
<%
if request("form1")<>"" then
call addshortmess()
end if
sub addshortmess()
dim rs,sqltext
set rs = Server.CreateObject("ADODB.Recordset")
sqltext = "select *From mess "
rs.open sqltext,conn,1,3
rs.addnew

rs("me")=request.cookies("renwen")("user")
rs("fd")=request.form("fd")
rs("ly")=replace(request.form("ly"),vbCrlf,"<br>")
rs("name")=request.form("name")
rs("date") = date
rs("time") = time
rs("ip")=address(Request.ServerVariables("REMOTE_ADDR"))
rs.Update
Rs.Close
Conn.Close
Set Rs=Nothing
Set Conn=Nothing
response.write "<script language=JavaScript>" & chr(13) & "alert('短信发送成功！谢谢你使用本程序');" & "window.close();" & "</script>"
end sub
%>
<html>
<head>
<title>查看短消息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../images/style.css" type="text/css">
<SCRIPT language=javascript id=clientEventHandlersJS>
<!--

function form1_onsubmit() 
{
  if(document.FORM1.name.value.length<1)
 {
   alert("你忘了写短信标题了");
   document.FORM1.name.focus();
   return false;
 }
   if(document.FORM1.ly.value.length<1)
 {
   alert("请填写好你要发给好友的短信");
   document.FORM1.ly.focus();
   return false;
  }

}

//-->
</SCRIPT>
<script language="JavaScript">
<!--
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_showHideLayers() { //v3.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v='hide')?'hidden':v; }
    obj.visibility=v; }
}
//-->
</script>
<link href="DEFAULT.css" rel="stylesheet" type="text/css">
</head>
<%
dim id,rs,sqltext
id=request("rmid")
conn.execute("update mess set click=click+1 where id="&id)
set rs=server.createobject("adodb.recordset")
sqltext="select * from mess where id="&id
rs.open sqltext,conn,1,1
%>
<body bgcolor="#E0E4FE" background="backimg/03.gif" leftmargin="0" topmargin="0">
<div align="center"></div>
<table width="466" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="466"> 
      <table cellspacing=0 cellpadding=0 width="466" border=0 align="center">
        <tbody> 
        <tr> 
            <td height="24" background="backimg/bg1.gif"> 
              <div align="center"><b>欢迎使用短消息接收</b>，<%=request.cookies("renwen")("user")%></div>
          </td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
  <tr>
    <td> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="100%" bgcolor="#92b9fb"> 
            <table width="100%" cellpadding="4" cellspacing="1" bordercolor="#6666CC">
              <tr> 
                <td width="469" height="14" bgcolor="#FFFFFF"><b><%=rs("me")%> 给您发送的消息！</b></td>
              </tr>
              <tr > 
                <td height="64" valign="middle" bgcolor="#FFFFFF"> 
                  <div align="center"> </div>
                  标题:<b> <%=rs("name")%></b><br>
                  <hr>
                  内容:&nbsp; <%=rs("ly")%></td>
              </tr>
              <tr bgcolor="#d6baff"> 
                <td valign="bottom" bgcolor="#FFFFFF"> 
                  <div align="right"><font size="2">短信来自:<%=rs("ip")&" "%>短信时间:<%=rs("date")&" "&rs("time")%> 
                    </font></div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<form language=javascript name=FORM1   onSubmit="return form1_onsubmit()"
      action=readmess.asp method=post>
  <table width="466" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td bgcolor="#92b9fb"><table width="100%" align=center cellpadding=3 cellspacing=1 class=tableborder1>
          <tbody>
            <tr align="center"> 
              <td height="26" colspan=3 background="backimg/bg1.gif"> 
                <input type="hidden" name="fd" value="<%=rs("me")%>"> 
                <b><font color="#000000">回复 <%=rs("me")%> 发送短消息（请输入完整信息）</font></b> </td>

            <tr> 
              <td valign=top bgcolor="#E0E4FE" class=tablebody1><b>标题：</b></td>
              <td valign=center bgcolor="#FFFFFF" class=tablebody1><font color="#0000FF"> 
                <input type="text" name="name" size="40" >
                </font> </td>
            </tr>
            <tr> 
              <td valign=top bgcolor="#E0E4FE" class=tablebody1><b>内容：</b></td>
              <td valign=center bgcolor="#FFFFFF" class=tablebody1> 
                <textarea rows="8" name="ly" cols="50" ></textarea> 
              </td>
            </tr>
            <tr> 
              <td colspan=2 bgcolor="#E0E4FE" class=tablebody1> 
                标题最多<b>50</b>个字符，内容最多<b>300</b>个字符 </td>
          </tr>
            <tr> 
              <td colspan=2 align=middle valign=center bgcolor="#E0E4FE" class=tablebody2> 
                <div align="center"> 
                <font color="#0000FF"> 
                <input name=FORM1 type=submit  value="回复短信">&nbsp;
                  <input type="reset" name="Reset" value="重 填">
                  &nbsp;&nbsp;
                <input type="button" name="Submit2" value="关闭窗口"onClick="javascript:window.close()">
                </font></div></td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>

  </body>
</html>
