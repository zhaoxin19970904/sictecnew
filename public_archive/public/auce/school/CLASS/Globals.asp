
<%
  '全局数据库参数.
   set schooldb = server.createobject("ADODB.connection")
   schooldb.open  "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("../schoolmate.asa")

  DBParams=schooldb
  '生成HTML脚本时美化生成结果时，自动带上回车换行符.
  CRLF = Chr(13)&Chr(10)

  Function PassWordOK(Byval UserID,Password) 'PasswordOK(UserID,PassWord)
    passwordok=true
    Set PWrs = Server.CreateObject("ADODB.Recordset")
    SQL = "select * from UserInfo where UserID='"&UserID&"'"
    PWrs.Open SQL,DBParams
    if not PWrs.eof then
        if PWrs("userpassword")=password then
            passwordok=true
        else
            passwordok=false
        end if
    else
        passwordok=false
    end if
    if len(password)>8 or not instr(password,"or")=0 then
        passwordok=false
    end if
    PWrs.Close
    set PWrs = nothing
  End Function
%>