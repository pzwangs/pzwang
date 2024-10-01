<!-- #include file="Config.asp"-->
<!-- #include file="Check_Session.asp" -->
<!-- #include file="md5.asp"-->
<!-- #include file="about.asp" -->
<!--#include file="head.asp"-->
<!-- #include file="top.asp"-->
<br>
<%
L_Reg=Request("LouHB_Reg")
L_GongHao=Request("LouHB_GongHao")
L_MiMa=Request("LouHB_MiMa")
L_MiMa2=Request("LouHB_MiMa2")
L_BeiZhu=Request("LouHB_BeiZhu")
ID=Request("LouHB_ID")
TianJ_XiuG=Request("LouHB_TianJ_XiuG")

Select Case L_Reg
Case "添　加"

If L_MiMa<>L_MiMa2 Then L_CuoWu("两次输入的新密码不一致，请重新输入！")
if Len(L_GongHao) < 1 then L_CuoWu("请输入用户名！")
if Len(L_MiMa) < 1 then L_CuoWu("请输入密码！")

Sub L_CuoWu(LouHB_CuoWu)
if LouHB_CuoWu <> "" then Response.Write("<script>alert('" & LouHB_CuoWu & "');history.back();</script>")
Response.End
End Sub

Set Lours=server.createobject("adodb.recordset")
Check="Select x_user From xy_admin where x_user='"&L_GongHao&"'"
Lours.Open Check,Louconn,1,1
If Not (Lours.EOF And Lours.BOF) Then L_CuoWu("该管理账号重名，请不要重复添加！")
Lours.Close
Set Lours=Nothing

L_MiMa=md5(L_MiMa,32)
SaveU="insert into xy_admin(x_user,x_password,x_mz) values('"&L_GongHao&"','"&L_MiMa&"','"&L_BeiZhu&"')"
Louconn.Execute(SaveU)
set SaveU=nothing
Call Copy()
response.write "<script language=JavaScript>" & chr(13) & "alert('恭喜，管理员添加成功！');"&"window.location.href='Add_manage.asp'"&" </script>"
Response.End

Case "修　改"
if Len(L_GongHao) < 1 then L_CuoWu("管理员账号必须填写！")
Set Lours=server.createobject("adodb.recordset")
up="Select * From xy_admin where id="&ID
Lours.Open up,Louconn,1,3
Lours("x_mz")=L_BeiZhu
If L_MiMa <> "" Then
Lours("x_password")=md5(L_MiMa)
End If
Lours.update
Lours.Close
Set Lours=Nothing
LouConn.Close
Set LouConn=Nothing
Call Copy()
response.write "<script language=JavaScript>" & chr(13) & "alert('恭喜，管理员修改成功！');"&"window.location.href='Add_manage.asp'"&" </script>"
Response.End
End Select

Select Case TianJ_XiuG
Case 2
Set Lours=server.createobject("adodb.recordset")
CX="Select * From xy_admin where id="&ID
Lours.Open CX,Louconn,1,1
IF Not(Lours.EOF And Lours.BOF) Then
L_GongHao=Lours("x_user")
L_MiMa=Lours("x_password")
L_BeiZhu=Lours("x_mz")
End IF
Lours.CLose
Set Lours=Nothing
End Select
%>
<table width="41%"  border="0" cellspacing="0" cellpadding="0" align="center" class="table4">
<form name="form1" method="post" action="?LouHB_TianJ_XiuG=<%=TianJ_XiuG%>&LouHB_ID=<%=ID%>">
<tr align="center">
<td height="30" colspan="2" background="images/bg.gif">
<%
If TianJ_XiuG=1 Then
response.write"添加管理员"
ElseIF TianJ_XiuG=2 Then
response.write"修改管理员"
End If
%></td>
</tr>
<tr>
<td width="22%" align="center">管理员账号</td>
<td width="78%">
<%
response.write"<input name='LouHB_GongHao' type='text' class='button1' size='30' "
If TianJ_XiuG=2 Then 
response.write"value='"&L_GongHao&"' " 
End If
response.write">"
%></td>
</tr>
<tr>
<td align="center">&nbsp;</td>
<td>&nbsp;</td>
</tr>
<tr>
<td align="center">密　码</td>
<td><input name="LouHB_MiMa" type="password" class="button1" size="30" value=""></td>
</tr>
<tr>
<td align="center">&nbsp;</td>
<td>&nbsp;</td>
</tr>
<tr>
<td align="center">密码确认</td>
<td><input name="LouHB_MiMa2" type="password" class="button1" size="30" value=""></td>
</tr>
<tr>
<td align="center">&nbsp;</td>
<td>&nbsp;</td>
</tr>

<tr>
<td align="center">备　注</td>
<td>
<input name="LouHB_BeiZhu" type="text" class="button1" size="30" value="<%=L_BeiZhu%>"></td>
</tr>
<tr>
<td align="center">&nbsp;</td>
<td>&nbsp;</td>
</tr>

<tr align="center">
<td colspan="2">
<%
If TianJ_XiuG=2 Then
response.write"<input name=""LouHB_Reg"" type=""submit"" class=""button1"" value=""修　改"">"
Else
response.write"<input name=""LouHB_Reg"" type=""submit"" class=""button1"" value=""添　加"">"
End If
%></td>
</tr>
<tr>
<td colspan="2">&nbsp;</td>
</tr>
</form>
</table>
<br>
<%
Set Lours=server.createobject("adodb.recordset")
List="Select * From xy_admin"
Lours.Open List,LouConn,1,1
If Lours.EOF And Lours.BOF Then
response.write"管理员未添加……"
Else
%>
<table width="768" border="0" cellspacing="0" cellpadding="0" align="center" class="table4">
<tr>
<td width="167" height="30" align="center" background="images/btg.gif" class="td3">管理员账号</td>
<td width="187" align="center" background="images/btg.gif">备注</td>
<td width="412" align="center" background="images/btg.gif">操　　作</td>
</tr>
<%Do While Not Lours.EOF%>
<tr>
<td height="30" align="center" class="table11"><%=Lours("x_user")%></td>
<td align="center" class="table11">
<%
If Lours("x_mz")="" Then
response.write"&nbsp;"
Else
response.write""&Lours("x_mz")&""
End If
%>
</td>
<td align="center" class="table11">
<a href="Add_Manage.asp?LouHB_ID=<%=Lours("ID")%>&LouHB_TianJ_XiuG=2">修　　改</a>　　　
<%
If Lours("x_user") <> Session("x_user") Then
response.write"<a href='Delete.asp?LouHB_Table=xy_admin&LouHB_ID="&Lours("ID")&"' onClick=""return confirm('操作提示：您确实要删除该管理员吗？');"">"
End If
response.write"删　　除"
If Lours("x_user") <> Session("x_user") Then
response.write"</a>"
End If
%>
</td>
</tr>
<%
Lours.Movenext
Loop
%>
</table>
<br>
<%
End If
Lours.CLose
Set Lours=Nothing
LouConn.Close
Set LouConn=Nothing
Call Copy()
%>