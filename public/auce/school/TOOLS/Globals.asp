
<%
  'ȫ�����ݿ����.
   set schooldb = server.createobject("ADODB.connection")
   schooldb.open  "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("../schoolmate.asa")

  DBParams=schooldb
  '����HTML�ű�ʱ�������ɽ��ʱ���Զ����ϻس����з�.
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