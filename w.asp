<!--#include file="include/head.asp"-->
<!--#include file="include/md5.asp"-->
<%act=Request.QueryString("act")%>

<%If act="logins" Then '��¼%>
<html>
<head>
<meta http-equiv="Content-Type" Content="text/html; charset=gb2312">
<title>�û���¼</title>
<link href="img/<%=strStyle%>" rel="stylesheet" type="text/css" />
</head>
<body>
<%If User_ID="" Then%>
<div id="dltop"><p>�û���¼</p></div>
<div id="denglu">
<form action="w.asp?act=login&id=1" method="post" name="form2">
<p>�û�����<input type="text" name="username" size="14" class="input" /></p>
<p>��&nbsp;&nbsp;&nbsp;&nbsp;�룺<input type="password" name="userpass" size="14" class="input" /></p>
<p><input type="checkbox" value="1" name="CookieTime" checked="checked">��ס����</p>
<p><input type="submit" value="��¼" name="B1" tabindex="3" class="dlbutton" />&nbsp;&nbsp;<input type="button" value="ȡ��" onclick="javascript:window.close();" class="dlbutton" /></p>
</div>
<%Else%>
<script language="JavaScript" type="text/JavaScript">window.opener.location.reload(); window.close();</script>
<%End If%>
</body>
</html>
<%End If%>

<%If act="member" Then '��¼%>
<%
'��ʾ�û���Ϣ
Dim Form_UserName
Form_UserName=Request.QueryString("username")
If Form_UserName="" Then Form_UserName=User_ID
Sql="select * from Users where UserName='"&SqlShow(Form_UserName)&"'"
OpenRs(Sql)
If Rs.RecordCount=0 Then Call ShowError("�޴��û�",0)
%>
<html>
<head>
<meta http-equiv="Content-Type" Content="text/html; charset=gb2312">
<title><%=HtmlEncode(Form_UserName) & "�ĸ�������"%></title>
<link rel="stylesheet" type="text/css" href="img/<%=strStyle%>" />
<script language="javascript" >
function askdelUsers()
{
	if(confirm("��ȷ��Ҫɾ�����û���"))
		return(true);
	else
		return(false);
}
function askinitpass()
{
	if(confirm("��ȷ�����ø��û�������"))
		return(true);
	else
		return(false);
}
</script>
</head>
<body>
<div class="zhucebg">
<div class="zhuce_1">��<%=HtmlEncode(Form_UserName)%>���ĸ������� </div>
<div class="zhuce_2"><span>��¼������</span>&nbsp;&nbsp;<%=Rs("LogonCount")%></div>
<div class="zhuce_2"><span>ע��ʱ�䣺</span>&nbsp;&nbsp;<%=Rs("RegTime")%></div>
<div class="zhuce_2"><span>����¼��</span>&nbsp;&nbsp;<%=Rs("LastLogonTime")%></div>
<div class="zhuce_2"><span>�����ʼ���</span>&nbsp;&nbsp;<%=Rs("Email")%></div>
<div class="zhuce_2"><span>������ҳ��</span>&nbsp;&nbsp;<a target=_blank href="<%=Rs("Userweb")%>"><%=Rs("Userweb")%></a></div>
</div>
<div class="zhuce_4">
<%If User_ID="admin" Then Response.write("<a class=""hrefbutton"" href='admin.asp?act=delUsers&id="&Rs("id")&"' onclick='javascript:return askdelUsers();'>ɾ���û�</a><a class=""hrefbutton"" href='admin.asp?act=initpass&id="&Rs("id")&"' onclick='javascript:return askinitpass();'>��������</a>")%>
<input onClick="javascript:window.close();" type="button" value="�ر�" name="B1" class="zhuce_button">
</div>
</body>
</html>
<%End If%>

<%If act="reg" Then 'ע���Ա%>
<%
Dim Msg,Form_Password,Form_Email,Form_Userweb
Act=Request.QueryString("regto")
If Act<>"new" And Act<>"edit" Then
	If User_ID="" Then
		 Act="new"
	Else
		 Act="edit"
		 Call LoadUser
	End If
Else
	If Act="new" Then Call SaveUser Else Call EditUser
End If

Sub LoadUser
	Sql="select * from Users where UserName='"&SqlShow(User_ID)&"'"
	OpenRs(Sql)
	Form_UserName=User_ID
	Form_Password="v8KWs5Uj"
	Form_Email=Rs("Email")
	Form_Userweb=Rs("Userweb")
	Rs.Close
End Sub

Sub EditUser
	Form_Password=Request.Form("userpass")
	Form_Email=Request.Form("email")
	Form_Userweb=Request.Form("userweb")

	Sql="select * from Users where UserName='" & SqlShow(User_ID) & "'"
	OpenRs(Sql)
	If Form_Password<>"v8KWs5Uj" Then
	Rs("UserPass")=MD5(Form_Password)
	Response.cookies(systemkey)("userpassword")=MD5(Form_Password)
	End If
	Rs("Email")=Form_Email
	Rs("Userweb")=Form_Userweb
	Rs.Update
	Rs.Close
	Act="Save"
	Call ShowError("�޸ĳɹ���",0)
End Sub

Sub SaveUser
	Form_UserName=Request.Form("username")
	Form_Password=Request.Form("userpass")
	Form_Email=Request.Form("email")
	Form_Userweb=Request.Form("userweb")
	If IsSpecial(Form_UserName)=True or CCEmpty(Form_UserName)<>Form_UserName or Form_UserName="" Then
		 Msg="�û����в����԰��������ַ�!"
		 Exit Sub
	End If
	Sql="select * from Users where UserName='" & SqlShow(Form_UserName) & "'"
	OpenRs(Sql)
	If Rs.RecordCount>0 Then
		 Rs.Close
		 Msg="���û����Ѿ����ڣ�������ѡ���û���"
		 Exit Sub
	End If
	If strlen(Form_UserName)<3 Then
	        Rs.Close
		Msg="�û�������С��3���ַ�"
		Exit Sub
	End If
	Rs.AddNew
	Rs("UserName")=Form_UserName
	Rs("UserPass")=MD5(Form_Password)
	Rs("Email")=Form_Email
	Rs("Userweb")=Form_Userweb
	Rs("RegTime")=Date

	Rs.Update
	Rs.Close
	If Enable_Cookies = 1 Then
			Response.cookies(systemkey)("userid")=Form_UserName
			Response.cookies(systemkey)("userpassword")=MD5(Form_Password)
       Else
		Session(SystemKey & "User_ID")=Form_UserName
    End If
	Msg="�ѳɹ�ע�ᣬ�����Ѿ���ע���û���ݵ�¼"
	Act="Save"
End Sub
%><html>
<head>
<meta http-equiv="Content-Type" Content="text/html; charset=gb2312">
<title><%If Act="new" Then%>ע���û�<%Else%>�޸�����<%End If%></title>
<link href="img/<%=strStyle%>" rel="stylesheet" type="text/css" />
<script src="include/js.js" type="text/javascript"></script>
<script language="javascript">
function checkform()
{
	if(form1.username.value==""){
		form1.username.focus();
		return(false);
	}
	if(form1.userpass.value==""){
		form1.userpass.focus();
		return(false);
	}
	if(form1.passcheck.value==""){
		form1.passcheck.focus();
		return(false);
	}
	if(form1.passcheck.value!=form1.userpass.value){
		alert("������������벻һ��!");
		form1.userpass.focus();
		return(false);
	}
}
<%
If Msg<>"" Then
Printl "alert(""" & Msg & """);"
End If
If Act="Save" Then
	Act="edit"
%>
window.opener.document.location.reload();
window.close();
<%
End If
%>
</script>
<body >
<form name="form1" action="w.asp?act=reg&regto=<%=Act%>" method="post" onSubmit="javascript:return checkform();">
<div class="zhucebg">
<div class="zhuce_1">�û�����</div>
<div class="zhuce_2"><span class="zhuce_color">������</span><input <%if act="edit" then%>readonly <%end if%>maxlength="9" type="text" name="username" size="28" class="input" value="<%=HtmlEncode(Form_UserName)%>"></div>
<div class="zhuce_2"><span class="zhuce_color">���룺</span><input maxlength="20" type="password" name="userpass" size="28" class="input" value="<%=HtmlEncode(Form_Password)%>"></div>
<div class="zhuce_2"><span class="zhuce_color">ȷ�����룺</span><input maxlength="20" type="password" name="passcheck" size="28" class="input" value="<%=HtmlEncode(Form_Password)%>"></div>
<div class="zhuce_2"><span> Email��</span><input maxlength="48" type="text" name="email" size="28" class="input" value="<%=HtmlEncode(Form_Email)%>"></div>
<div class="zhuce_2"><span>������ҳ��</span><input name="userweb" type="text" class="input" id="userweb" value="<%=HtmlEncode(Form_Userweb)%>" size="28" maxlength="48"></div>
<div class="zhuce_4"><input type="submit" value=" ȷ �� " name="B1" onMouseOut="javascript:this.style.backgroundColor='';" class="zhuce_button">&nbsp;&nbsp;<input type="reset" value=" �� �� " name="B2" onMouseOut="javascript:this.style.backgroundColor='';" class="zhuce_button"></div>
</div>
</form>
</body>
</html>
<%End If%>

<%
If act="login" Then '��¼��֤
Dim Action
Action=Request("id")
Select Case Action
	Case "1"
		Form_UserName=Request("username")
		Form_Password=Request("userpass")
		Sql="select * from Users where UserName='" & SqlShow(Form_UserName) & "' and UserPass='" & MD5(Form_Password) & "'"
		OpenRs(Sql)
		If Rs.RecordCount>0 Then
			If Enable_Cookies = 1 Then
			Response.cookies(systemkey)("userid")=rs("username")
			Response.cookies(systemkey)("userpassword")=rs("userpass")
            Select Case Request.Form("CookieTime")
				Case 1
					Response.Cookies(SystemKey).Expires=Date+31
			End Select
			Else
			Session(SystemKey & "User_ID")=Form_UserName
			End If
			Rs("LogonCount")=Rs("LogonCount")+1
			Rs("LastLogonTime")=Now
			Rs.update
			Action="OK"
		else
		Call ShowError("�û������������",1)
		End If
	Case ""
		If Enable_Cookies = 1 Then
		Response.cookies(systemkey)("userid")=""
		Response.cookies(systemkey)("userpassword")=""
		Session(SystemKey & "User_ID")=""
		Else
		Session(SystemKey & "User_ID")=""
		End If
		User_ID=""
	Case ""
		TurnTo "/"
End Select
TurnTo Request.ServerVariables("HTTP_REFERER")
CloseAll
End If
%>

<%
'����
If act="TextReply" Then
If EnableReply="2" And User_ID="" Then Call ShowError("�ظ�ǰ�����ȵ�¼��",1)
If EnableReply="3" And User_ID="" Then User_ID="�ο�"
Dim Form_Context,Form_ID
Form_Context=Request.Form("Context")
Form_ID=Request.Form("artid")

If Form_Context="" Then Call ShowError("���ݲ���Ϊ�գ�",1)
Sql="select * from WeiBo where ID=" &Form_ID
OpenRs(Sql)
'��������΢���Ļظ���
Rs("ReplyNum")=Rs("ReplyNum")+1
Rs.Update
'�����ظ�����
Rs.AddNew
Rs("Parent")=Form_ID
Rs("Content")=Form_Context
Rs("Poster")=User_ID
Rs("PostTime")=Now
Rs("Ip")=RemoteIP()
Rs.Update
Rs.Close
If EnableReply<>"3" Then
'�����û��ظ�����
Sql="select * from Users where UserName='"&SqlShow(User_ID)&"'"
OpenRs Sql
Rs("ReplyNum")=Rs("ReplyNum")+1
Rs.Update
Rs.Close
End If
Response.write("<script language='javascript'>alert('���۷���ɹ���');window.location='text.asp?id="&Form_ID&"';</script>")
End If
%>

<%
'����΢��
If act="WeiBopost" Then
If EnablePost="2" And User_ID="" Then Call ShowError("�ݲ������οͷ���",1)
If EnablePost="3" Then User_ID="�ο�"
Form_Context=Request.Form("Context")
Form_ID=Request.Form("artid")

If Form_Context="" Then Call ShowError("���ݲ���Ϊ�գ�",1)
Sql="select * from WeiBo"
OpenRs(Sql)
Rs.AddNew
Rs("Content")=Request.Form("Context")
Rs("Poster")=User_ID
Rs("PostTime")=Now
Rs("Ip")=RemoteIP()
Rs("Parent")=0
Rs.Update
Rs.Close
Response.write("<script language='javascript'>alert('����ɹ���');window.location='index.asp';</script>")
End If
%>

<%Call CloseAll()%>
