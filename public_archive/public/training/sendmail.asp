<style type="text/css">
<!--
body,td,p,th{font-size:14px;line-height:180%;}
input{font-size:12px;}
-->
</style>


<%@LANGUAGE="vb-script" CODEPAGE="936"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>mail</title>
</head>
<body>
<% 
Dim JMail,SendMail 
 Set JMail=Server.CreateObject("JMail.Message") 
 JMail.silent=true     
 JMail.Logging=True    
 JMail.Charset="gb2312"   
 JMail.From="sictec.bj@sictec.com"   
 JMail.fromname="li"            
 JMail.mailserverusername="123456"    
 JMail.MailserverPassword="miao"     
 JMail.addRecipient "sictec.bj@sictec.com"    
 JMail.Subject="charter"    
 JMail.Body="charter 2"          
 JMail.send ("sictec.bj#210.82.98.165")   
 JMail.close
%>
</body>
</html>
