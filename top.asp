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
<td align="center"><a href="index.asp">������ҳ</a></td>
<%If Session("x_user") = "" Then%>
<td align="center"><a href="Login.asp">�ǡ���¼</a></td>
<%
End If
If Session("x_user") <> "" Then
%>
<td align="center"><a href="Add_User.asp?doActs=Add_New">����û�</a></td>
<td align="center"><a href="Add_Manage.asp">��ӹ���Ա</a></td>
<td align="center">��ã�<%=web_x_mz%> <a href="logout.asp"> --�ˡ���--</a></td>
<%End If%>
</tr>
</table>