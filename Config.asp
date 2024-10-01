<%
on error resume next

dim LouConn,ConnStr,db
db="xing.mdb" 
Set LouConn = Server.CreateObject("ADODB.Connection")
ConnStr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
LouConn.Open ConnStr

If Err then
Err.Clear
Set Louconn = Nothing
Response.Write "系统调整中......请稍候再试！！"
Response.End
End If


dim Logo_alt : Logo_alt = "思翔 www.sxang.cn"
dim Page_Size : Page_Size = 20 '每页显示的人数



'出错提示
Sub ld_error(msg, url)
	Select Case Trim(url)
	Case "0"
		url = "javascript:history.back()"
	Case "1"
		url = Request.ServerVariables("HTTP_REFERER") 
	End Select
	Response.Write "<script Language=""Javascript"">alert('" & msg & "');location.href = '" & url & "';</script> "
	Response.End()
End Sub
%>