<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<HTML>
<HEAD>
<title>商务校友录</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="同学、同学录、校友、老师、学校、班级、交流">
<STYLE type=text/css>
</STYLE>
<LINK href="../scss.css" rel=stylesheet>
</HEAD>
<BODY BGCOLOR=#FFFFFF leftmargin="0" topmargin="0"><!--#include file=../top1.htm-->

<table width="550" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center" valign="top"> <br>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <tr>
           <td height="20" align="center" valign="top" class="cn">
<%
'on error resume next
Server.ScriptTimeout = 1000
UserIP = Request.ServerVariables("REMOTE_ADDR")
BadIP=""
BadIP=BadIP&"202.100.57.28"
ClassID=Trim(Session("ClassID"))
UserID=Trim(Session("UserID"))
'ClassID="GoodLuck"
'UserID="GoodLuck"
'or UserID="" or ClassID=""

if not InStr(BadIP,UserIP)=0 or UserID="" or ClassID=""  then
   writeln "请您先登陆"
else
 			set conn = server.createobject("ADODB.connection")
			conn.open DBParams
 			sql="select max(PicID) from upload where ClassID='"&ClassID&"'"
 			set rs=conn.Execute(sql)
 			if not rs.eof then
 			   curnum=rs(0)
 			end if
 			if curnum<1 or IsNull(curnum) then curnum=1
 			rs.close
			set rs=nothing
			
 set FileUp=server.createobject("ChinaASP.UpLoad")
 dim iCount
 iCount=1
 for each f in fileup.Files 
	if f.isempty=false then
		if f.filesize<>0 then
		   
		   '下面的语句用于将文件内容保存到sql server 7内。

 			if iCount=1 then
               sql = "Insert Into upload (ClassID,PicID,UserIP,FileName,FileSize,UserID,Title,PicType,Comment) Values('"&ClassID&"','"&iCount+curnum&"','"&UserIP&"','"&ClassID&iCount+curnum&"."&getFileExtName(f.filename)&"','"&f.filesize&"','"&UserID&"','"&replace(fileup.Form(2).value,"'","‘")&"','"&fileup.Form(3).value&"','"&replace(fileup.Form(4).value,"'","‘")&"')"
            else
               sql = "Insert Into upload (ClassID,PicID,UserIP,FileName,FileSize,UserID,Title,PicType,Comment) Values('"&ClassID&"','"&iCount+curnum&"','"&UserIP&"','"&ClassID&iCount+curnum&"."&getFileExtName(f.filename)&"','"&f.filesize&"','"&UserID&"','"&replace(fileup.Form(2).value,"'","‘")&iCount&"','"&fileup.Form(3).value&"','"&replace(fileup.Form(4).value,"'","‘")&"')"
			end if 			
			conn.Execute(sql) 			
                      
		   'writeln  "服务器上文件路径是  <font color=red>UserPhoto"&cstr(iCount+curnum)&"."&getFileExtName(f.fileName) & "</font>"
		   'writeln  "<br><font color=red>注意：</font>请您将上面红色的部分复制到‘图片’框中<br><br><a href=javascript:window.close()>CLOSE</a>"
		   writeln iCount&"文件上传完毕，你的照片已经成功添加在班级相册中！"
		else
		   writeln iCount&"文件为空上传失败"
		end if
		f.saveas server.MapPath("../Pic/") & "\"&ClassID&iCount+curnum&"."&getFileExtName(f.fileName)

	else
		writeln iCount&"文件没有选择,上传失败"
	end if
	iCount=iCount+1
next
set conn=nothing
set fileup=nothing

sub WriteLn(str)
	Response.Write str & "<br>"
end sub

'获得文件的后缀名
function getFileExtName(fileName)
	dim pos
	pos=instrrev(filename,".")
	if pos>0 then 
		getFileExtName=mid(fileName,pos+1)
	else
		getFileExtName=""
	end if
end function 


end if
%> 
</td>
        </tr>
      </table>
      <br>
      <br>
      <br>
      <table width="510" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="20" align="center" valign="top" class="cn"><a href="javascript:history.go(-1)">
          <font color="#000000">返回</font></a><font color="#000000"> 
            | </font> <a href="javascript:window.close()"><font color="#000000">关闭窗口</font></td>
        </tr>
      </table>
      <br>
    </td>
  </tr>
</table>
</BODY>
</HTML>