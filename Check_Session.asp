<%
Dim x_user,web_x_mz
x_user=Trim(Session("x_user"))
if x_user = "" then Response.Redirect("login.asp")
set rs=LouConn.Execute("select * from xy_admin where x_user='" & x_user & "'")
if not(rs.eof or rs.bof) then
	web_x_mz=Rs("x_mz")
else
	Response.Redirect("login.asp")
end if
rs.close:set rs=nothing

%>