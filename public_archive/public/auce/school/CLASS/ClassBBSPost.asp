<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
   ClassID=Trim(Request("ClassID"))
   UserID=Request("UserID")
   PassWord=Trim(Request("PassWord"))
   UserIP=Request.ServerVariables("REMOTE_ADDR")
   ToUserID=Trim(Request("ToUserID"))
   Title = Trim(CStr(request("Title")))	
      if Title="" then response.redirect "../Error.asp?Info=文章没有标题!"
      if len(Title)>50 then response.redirect "../Error.asp?Info=文章标题太长!"
   Comment = CStr(Request("Comment"))
   
 set rs = server.createobject("ADODB.recordset")
  
 '验证用户口令
   ok=true
   if UserID=Session("UserID") then
      ok=true
      '为管理加入特殊用户
      'if Session("UserID")="GoodLuck" and not PassWord="" then UserID=PassWord
   else
      if PasswordOK(UserID,PassWord) then Session("UserID")=UserID:ok=true
   end if 
  
 SQL = "select Max(MsgID) as m from message where ClassID='"&ClassID&"'"
 rs.Open SQL,DBParams,1,1
 
 if IsNull(rs("m")) then 
 	MsgID=1
 else
 	MsgID=rs("m")+1
 end if   
 rs.Close

 '禁止 keyWord
 cando=true

 'SQL = "select * from KeyWord"
 'rs.Open SQL,DBParams
 'flag = true
 'do while not rs.EOF and flag
 '   if InStr(1,Comment,trim(rs("KeyWord"))) <> 0 or InStr(1,Title,trim(rs("KeyWord"))) <> 0 then
 '      flag = false
 '      cando = false
 '      cannotdoMsg = "此文章不宜在此发表"
 '   end if
 '   rs.MoveNext
 'loop
 'rs.Close

'开始填写帖子
   
   if (ok and cando) then 
	   Set Conn2 = Server.CreateObject("ADODB.Connection")
	   Conn2.Open DBParams	   

	   Comment=replace(Comment,"<","《")
	   Comment=replace(Comment,">","》")
	   PostBody=replace(Comment,Chr(13)&Chr(10),"<br>")
	   
	   'curiturl = Request("iturl")
	   'if curiturl<>"http://" then
	   '	URLName=Request("urlname")
	   '	if URLName="" then URLName=curiturl
	   '	PostBody=PostBody&"<br><br><a href='"&curiturl&"'>"&URLName&"</a>"
	   'end if
	   
	   'curimg = Request("img")	
	   'if curimg<>"" then
	   '	PostBody=PostBody&"<br><img src='"&curimg&"'>"
	   'end if	   	
	  
	   '禁止html--->tag	   
 	   NoHtmlTitle=Title
 	   NoHtmlTitle=replace(NoHtmlTitle,"<","《")
 	   NoHtmlTitle=replace(NoHtmlTitle,">","》")
 	   
	   SQL = "select top 1 * from message where  ClassID='"&ClassID&"' order by MsgID DESC"
	   rs.open SQL,DBParams
	   if not rs.eof then
	      if rs("UserID")=UserID and rs("Title")=NoHtmlTitle and rs("Comment")=PostBody then
	   	RegTime=rs("RegTime")
	   	     rs.close
		     response.redirect "../Error.asp?Info=您发了重复的贴子!"
		     response.end		
	     end if	   	   
	   end if
	   rs.close

	  

	   rs.Open "Select top 1 * from message where 1=2",DBParams,1,3
	   rs.AddNew
	    rs("MsgID") =MsgID
	    rs("UserID")=UserID
	    rs("ClassID") = ClassID
	    rs("RegTime") = Now
	    rs("Title") =NoHtmlTitle 'SubJect->Title
	    rs("Comment") = PostBody
	    'rs("MsgSize") = len(PostBody)
	    'rs("Hits") = 0
	    rs("HeadPic") = CStr(Request("HeadPic"))
	    'rs("UserIP")=UserIP
	    if not ToUserID="" then 
	       if ToUserID="0" then response.redirect "../Error.asp?Info=您没有选择接受人名称！"
	       rs("ToUserID")=ToUserID
	    end if
	    rs.Update
	    rs.Close 
          conn2.close
           	
   
    else
	if not ok then 
	      response.redirect "../Error.asp?Info=密码与用户名称不符"
       end if
       if not cando then
          response.redirect "../Error.asp?Info="&cannotdoMsg
       end if
    
    end if

set conn2 = nothing
set rs = nothing

select case Trim(LCase(Request.ServerVariables("http_referer")))
   case "popwindow_perliuyan.asp"
      response.redirect "popwindow_Perliuyan.asp"
   case "classbbslist.asp"
      response.redirect "ClassBBSList.asp"
   case else
      response.redirect "../Success.asp?Info=提交完成，谢谢！"  '&LCase(Request.ServerVariables("http_referer"))
end select

%>