<!--#include file="include/head.asp"-->
<!--#include file="include/md5.asp"-->
<%act=Request.QueryString("act")%>

<%If act="logins" Then '登录%>
<html>
<head>
<meta http-equiv="Content-Type" Content="text/html; charset=gb2312">
<title>用户登录</title>
<link href="img/<%=strStyle%>" rel="stylesheet" type="text/css" />
</head>
<body>
<%If User_ID="" Then%>
<div id="dltop"><p>用户登录</p></div>
<div id="denglu">
<form action="w.asp?act=login&id=1" method="post" name="form2">
<p>用户名：<input type="text" name="username" size="14" class="input" /></p>
<p>密&nbsp;&nbsp;&nbsp;&nbsp;码：<input type="password" name="userpass" size="14" class="input" /></p>
<p><input type="checkbox" value="1" name="CookieTime" checked="checked">记住密码</p>
<p><input type="submit" value="登录" name="B1" tabindex="3" class="dlbutton" />&nbsp;&nbsp;<input type="button" value="取消" onclick="javascript:window.close();" class="dlbutton" /></p>
</div>
<%Else%>
<script language="JavaScript" type="text/JavaScript">window.opener.location.reload(); window.close();</script>
<%End If%>
</body>
</html>
<%End If%>

<%If act="member" Then '登录%>
<%
'显示用户信息
Dim Form_UserName
Form_UserName=Request.QueryString("username")
If Form_UserName="" Then Form_UserName=User_ID
Sql="select * from Users where UserName='"&SqlShow(Form_UserName)&"'"
OpenRs(Sql)
If Rs.RecordCount=0 Then Call ShowError("无此用户",0)
%>
<html>
<head>
<meta http-equiv="Content-Type" Content="text/html; charset=gb2312">
<title><%=HtmlEncode(Form_UserName) & "的个人资料"%></title>
<link rel="stylesheet" type="text/css" href="img/<%=strStyle%>" />
<script language="javascript" >
function askdelUsers()
{
	if(confirm("你确定要删除该用户吗？"))
		return(true);
	else
		return(false);
}
function askinitpass()
{
	if(confirm("你确定重置该用户密码吗？"))
		return(true);
	else
		return(false);
}
</script>
</head>
<body>
<div class="zhucebg">
<div class="zhuce_1">【<%=HtmlEncode(Form_UserName)%>】的个人资料 </div>
<div class="zhuce_2"><span>登录次数：</span>&nbsp;&nbsp;<%=Rs("LogonCount")%></div>
<div class="zhuce_2"><span>注册时间：</span>&nbsp;&nbsp;<%=Rs("RegTime")%></div>
<div class="zhuce_2"><span>最后登录：</span>&nbsp;&nbsp;<%=Rs("LastLogonTime")%></div>
<div class="zhuce_2"><span>电子邮件：</span>&nbsp;&nbsp;<%=Rs("Email")%></div>
<div class="zhuce_2"><span>个人主页：</span>&nbsp;&nbsp;<a target=_blank href="<%=Rs("Userweb")%>"><%=Rs("Userweb")%></a></div>
</div>
<div class="zhuce_4">
<%If User_ID="admin" Then Response.write("<a class=""hrefbutton"" href='admin.asp?act=delUsers&id="&Rs("id")&"' onclick='javascript:return askdelUsers();'>删除用户</a><a class=""hrefbutton"" href='admin.asp?act=initpass&id="&Rs("id")&"' onclick='javascript:return askinitpass();'>重设密码</a>")%>
<input onClick="javascript:window.close();" type="button" value="关闭" name="B1" class="zhuce_button">
</div>
</body>
</html>
<%End If%>

<%If act="reg" Then '注册会员%>
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
	Call ShowError("修改成功！",0)
End Sub

Sub SaveUser
	Form_UserName=Request.Form("username")
	Form_Password=Request.Form("userpass")
	Form_Email=Request.Form("email")
	Form_Userweb=Request.Form("userweb")
	If IsSpecial(Form_UserName)=True or CCEmpty(Form_UserName)<>Form_UserName or Form_UserName="" Then
		 Msg="用户名中不可以包含特殊字符!"
		 Exit Sub
	End If
	Sql="select * from Users where UserName='" & SqlShow(Form_UserName) & "'"
	OpenRs(Sql)
	If Rs.RecordCount>0 Then
		 Rs.Close
		 Msg="该用户名已经存在，请另外选择用户名"
		 Exit Sub
	End If
	If strlen(Form_UserName)<3 Then
	        Rs.Close
		Msg="用户名不能小于3个字符"
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
	Msg="已成功注册，现在已经以注册用户身份登录"
	Act="Save"
End Sub
%><html>
<head>
<meta http-equiv="Content-Type" Content="text/html; charset=gb2312">
<title><%If Act="new" Then%>注册用户<%Else%>修改资料<%End If%></title>
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
		alert("两次输入的密码不一致!");
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
<div class="zhuce_1">用户资料</div>
<div class="zhuce_2"><span class="zhuce_color">姓名：</span><input <%if act="edit" then%>readonly <%end if%>maxlength="9" type="text" name="username" size="28" class="input" value="<%=HtmlEncode(Form_UserName)%>"></div>
<div class="zhuce_2"><span class="zhuce_color">密码：</span><input maxlength="20" type="password" name="userpass" size="28" class="input" value="<%=HtmlEncode(Form_Password)%>"></div>
<div class="zhuce_2"><span class="zhuce_color">确认密码：</span><input maxlength="20" type="password" name="passcheck" size="28" class="input" value="<%=HtmlEncode(Form_Password)%>"></div>
<div class="zhuce_2"><span> Email：</span><input maxlength="48" type="text" name="email" size="28" class="input" value="<%=HtmlEncode(Form_Email)%>"></div>
<div class="zhuce_2"><span>个人主页：</span><input name="userweb" type="text" class="input" id="userweb" value="<%=HtmlEncode(Form_Userweb)%>" size="28" maxlength="48"></div>
<div class="zhuce_4"><input type="submit" value=" 确 定 " name="B1" onMouseOut="javascript:this.style.backgroundColor='';" class="zhuce_button">&nbsp;&nbsp;<input type="reset" value=" 清 空 " name="B2" onMouseOut="javascript:this.style.backgroundColor='';" class="zhuce_button"></div>
</div>
</form>
</body>
</html>
<%End If%>

<%
If act="login" Then '登录验证
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
		Call ShowError("用户名或密码错误！",1)
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
'评论
If act="TextReply" Then
If EnableReply="2" And User_ID="" Then Call ShowError("回复前，请先登录！",1)
If EnableReply="3" And User_ID="" Then User_ID="游客"
Dim Form_Context,Form_ID
Form_Context=Request.Form("Context")
Form_ID=Request.Form("artid")

If Form_Context="" Then Call ShowError("内容不能为空！",1)
Sql="select * from WeiBo where ID=" &Form_ID
OpenRs(Sql)
'更新主题微博的回复数
Rs("ReplyNum")=Rs("ReplyNum")+1
Rs.Update
'创建回复内容
Rs.AddNew
Rs("Parent")=Form_ID
Rs("Content")=Form_Context
Rs("Poster")=User_ID
Rs("PostTime")=Now
Rs("Ip")=RemoteIP()
Rs.Update
Rs.Close
If EnableReply<>"3" Then
'更新用户回复数量
Sql="select * from Users where UserName='"&SqlShow(User_ID)&"'"
OpenRs Sql
Rs("ReplyNum")=Rs("ReplyNum")+1
Rs.Update
Rs.Close
End If
Response.write("<script language='javascript'>alert('评论发表成功！');window.location='text.asp?id="&Form_ID&"';</script>")
End If
%>

<%
'发表微博
If act="WeiBopost" Then
If EnablePost="2" And User_ID="" Then Call ShowError("暂不允许游客发表！",1)
If EnablePost="3" Then User_ID="游客"
Form_Context=Request.Form("Context")
Form_ID=Request.Form("artid")

If Form_Context="" Then Call ShowError("内容不能为空！",1)
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
Response.write("<script language='javascript'>alert('发表成功！');window.location='index.asp';</script>")
End If
%>

<%Call CloseAll()%>
