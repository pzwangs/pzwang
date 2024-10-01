<%
web_x_mz=""
set rs=LouConn.Execute("select * from xy_admin where x_user='" & Trim(Session("x_user")) & "'")
if not(rs.eof or rs.bof) then
	web_x_mz=Rs("x_mz")
end if
rs.close:set rs=nothing
%>
<table width="768" height="35" align="center" cellpadding="0" cellspacing="0" background="images/DHT.gif">
<tr>
<td align="center"><a href="index.asp">返回首页</a></td>
<%If Session("x_user") = "" Then%>
<td align="center"><a href="Login.asp">登　　录</a></td>
<%
End If
If Session("x_user") <> "" Then
%>
<td align="center"><a href="Add_User.asp?doActs=Add_New">添加用户</a></td>
<td align="center"><a href="Add_Manage.asp">添加管理员</a></td>
<td align="center">你好，<%=web_x_mz%> <a href="logout.asp"> --退　出--</a></td>
<%End If%>
</tr>
</table>