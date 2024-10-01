<!--#include file="config.asp"-->
<!-- #include file="Check_Session.asp" -->
<!--#include file="about.asp"-->
<!--#include file="head.asp"-->
<!--#include file="top.asp"-->
<br>
<%
ID=Request("LouHB_ID")
L_Table=Request("LouHB_Table")

Del="delete from "&L_Table&" where ID="&ID
Louconn.Execute Del
Set LouConn=Nothing

Call Copy()
response.write "<script language=JavaScript>" & chr(13) & "alert('¹§Ï²£¬É¾³ý³É¹¦£¡');"&"window.location.href = 'index.asp'"&" </script>"
response.end
%>