<%@ Language=VBScript %>
<!--#include file=globals.asp -->
<%
   ClassID=Trim(Request("ClassID"))
   PicID=Trim(Request("PicID"))
   UserID=Request("UserID")
   PassWord=Trim(Request("PassWord"))
   UserIP=Request.ServerVariables("REMOTE_ADDR")
   ToUserID=Trim(Request("ToUserID"))
   Title = Trim(CStr(request("Title")))	:if Title="" then response.redirect "../Error.asp?Info=����û�б���!"
   Comment = CStr(Request("Comment"))
   
 set rs = server.createobject("ADODB.recordset")
  
 '��֤�û�����
   ok=true
   if UserID=Session("UserID") then
      ok=true
      'Ϊ������������û�
      'if Session("UserID")="GoodLuck" and not PassWord="" then UserID=PassWord
   else
      if PasswordOK(UserID,PassWord) then Session("UserID")=UserID:ok=true
   end if 
  
 SQL = "select Max(MsgID) as m from UploadReply where ClassID='"&ClassID&"' and PicID='"&PicID&"'"
 rs.Open SQL,DBParams,1,1
 
 if IsNull(rs("m")) then 
 	MsgID=1
 else
 	MsgID=rs("m")+1
 end if   
 rs.Close

 '��ֹ keyWord
 cando=true

 'SQL = "select * from KeyWord"
 'rs.Open SQL,DBParams
 'flag = true
 'do while not rs.EOF and flag
 '   if InStr(1,Comment,trim(rs("KeyWord"))) <> 0 or InStr(1,Title,trim(rs("KeyWord"))) <> 0 then
 '      flag = false
 '      cando = false
 '      cannotdoMsg = "�����²����ڴ˷���"
 '   end if
 '   rs.MoveNext
 'loop
 'rs.Close

'��ʼ��д����
   
   if (ok and cando) then 
	   Set Conn2 = Server.CreateObject("ADODB.Connection")
	   Conn2.Open DBParams	   

	   Comment=replace(Comment,"<","��")
	   Comment=replace(Comment,">","��")
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
	  
	   '��ֹhtml--->tag	   
 	   NoHtmlTitle=Title
 	   NoHtmlTitle=replace(NoHtmlTitle,"<","��")
 	   NoHtmlTitle=replace(NoHtmlTitle,">","��")
 	   
	   SQL = "select top 1 * from UploadReply where  ClassID='"&ClassID&"' and PicID='"&PicID&"' order by MsgID DESC"
	   rs.open SQL,DBParams
	   if not rs.eof then
	   if rs("UserID")=UserID and rs("Title")=NoHtmlTitle and rs("Comment")=PostBody then
	   	RegTime=rs("RegTime")
	   	     rs.close
		     response.redirect "../Error.asp?Info=�������ظ�������!"
		     response.end
		
	   end if
	   if rs("UserID")=UserID then
	   	RegTime=rs("RegTime")
	   	i=DateDiff("s",RegTime,now)
	   	if i<15 then 
		     rs.close
		     response.redirect "../Error.asp?Info=�����˹�ˮ����!"
		     response.end
		end if
	   end if
	   	   
	  end if
	  rs.close

	  

	   rs.Open "Select top 1 * from UploadReply where 1=2",DBParams,1,3
	   rs.AddNew
	    rs("MsgID") =MsgID
	    rs("UserID")=UserID
	    rs("ClassID") = ClassID
	    rs("PicID") = PicID
	    rs("Title") =NoHtmlTitle 'SubJect->Title
	    rs("Comment") = PostBody
	    'rs("MsgSize") = len(PostBody)
	    rs("Hits") = 0
	    rs("HeadPic") = CStr(Request("HeadPic"))
	    rs("UserIP")=UserIP
	    'if not ToUserID="" then rs("ToUserID")=ToUserID
	    rs.Update
	    rs.Close 
          conn2.close
           	
   
    else
	if not ok then 
	      response.redirect "../Error.asp?Info=�������û����Ʋ���"
       end if
       if not cando then
          response.redirect "../Error.asp?Info="&cannotdoMsg
       end if
    
    end if

set conn2 = nothing
set rs = nothing
response.redirect "ClassPhotoView.asp?ClassID="&ClassID&"&PicID="&PicID
%>
