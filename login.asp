<!--#include file="config.asp"-->
<!--#include file="about.asp"-->
<!--#include file="md5.asp"-->
<%
Select Case Request.ServerVariables("REQUEST_METHOD")
Case "POST"
x_user=request("LouHB_GongHao")  '接受管理员用户名
x_password=request("LouHB_MiMa")  '接受管理员密码
If Request("LouHB_YanZhengM") = Session("L_YanZhengM") Then '3

if x_user<>"" and x_password<>"" then
set Lours=Louconn.execute("select * from xy_admin where x_user='"& x_user & "'")	 '判断是否存在该用户
if not (Lours.bof and Lours.eof) then '1
If Lours("x_password")=md5(x_password) Then '0
Session("x_user")=Lours("x_user")
Response.Redirect("index.asp")
Else '0
Cuo=1
End if '0
Else '1
Cuo=2
End if '1
End if
Else '3
Cuo=3
End If '3
End Select
%>
<!--#include file="yanz.asp"-->
<!--#include file="head.asp"-->
<Center>

<br>
<table width="768" border="0" align="center" cellpadding="0" cellspacing="0" class="table4">
<tr>
<td height="30" align="center" class=td2>管理员登录</td>
</tr>
</table>
<br><br>
<table width="300" border="0">
<form action="login.asp" method="post" name="form2" id="form2">
<tr>
<td height="30" colspan="2" align="Center">
<%
if Cuo=1 then 
Response.Write "<p class=p2>密码错误！</p>"
elseif Cuo=2 then
Response.Write "<p class=p2>登录失败！</p>"
elseif Cuo=3 then
Response.Write "<p class=p2>验证码错误！</p>"
end if
%>
<a href="index.asp"><img src="images/Start.gif" alt="返回首页" width="46" height="16" border="0"></a></td>
</tr>
<tr>
<td width="197" height="30" align="right">账 号：</td>
<td width="205" height="30">
<input name="LouHB_GongHao" type="text" class="button1" size="20" maxlength="20" value="" onFocus="this.select()" onMouseOver="this.focus();"></td>
</tr>
<tr>
<td height="30" align="right">密  码：</td>
<td height="30"><input name="LouHB_MiMa" type="password" class="button1" size="21" maxlength="20" value="" onFocus="this.select()" onMouseOver="this.focus();"></td>
</tr>
<td align="right">验证码：</td>
<td><input name="LouHB_YanZhengM" type="text" class="button1" size="6" maxlength="20" onFocus="this.select()" onMouseOver="this.focus();"><Font Color="Red"><%=gen_key(4)%></Font></td>
<td height="30"></td>
<tr align="center">
<td height="30" colspan="2">
<input type="image" name="Reg2" value="登录" src="images/login.gif" /></td>
</tr></form> 
</table>
</center>
<br>
<%Call Copy%>
</Center>