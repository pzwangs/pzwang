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
Response.Write "ϵͳ������......���Ժ����ԣ���"
Response.End
End If


dim Logo_alt : Logo_alt = "˼�� www.sxang.cn"
dim Page_Size : Page_Size = 20 'ÿҳ��ʾ������



'������ʾ
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