<!--#include file="include/head.asp"-->
<%
Dim ContextID,formID,i2,title,Page,Pages,ShowCount,PastCount,listMethod,listpage,j,showOrder,Poster,TextFanYe,ZhengWen,TextLieBiao,i3,plk
ContextID=Cnum(Request.QueryString("id"))
Page=Max(Cnum(sqlshow(Request("page"))),1)
listpage=sqlshow(Request.QueryString("listpage"))
HtmlFiles="text"

Sql="select * from WeiBo where id="&ContextID&" or Parent="&ContextID&" order by Parent,PostTime"
OpenRs(Sql)

'显示内容
If EnablePost="1" Then
If User_ID="admin" Then
zhengwen="<p>"&Ubbcode(HTMLEncode(Rs("Content")))&"</p><p class=""TextTime"">"&Rs("PostTime")&" | <a>回复("&Rs("ReplyNum")&")</a> <a href=""admin.asp?act=WeiBoedit&id="&Rs("id")&""">编辑</a> - <a href=""admin.asp?act=WeiBodel&id="&Rs("id")&"&Parent=0"">删除</a></p>"
Else
zhengwen="<p>"&Ubbcode(HTMLEncode(Rs("Content")))&"</p><p class=""TextTime"">"&Rs("PostTime")&" | <a>回复("&Rs("ReplyNum")&")</a></p>"
End If
Else
If User_ID="admin" Then
zhengwen="<p>"&Ubbcode(HTMLEncode(Rs("Content")))&"</p><p class=""TextTime"">分享者:<a title=""显示["&Rs("poster")&"]的基本信息"" href=""javascript:memberinfo('"&Rs("Poster")&"')"">"&Rs("Poster")&"</a> | "&Rs("PostTime")&" | <a>回复("&Rs("ReplyNum")&")</a> <a href=""admin.asp?act=WeiBoedit&id="&Rs("id")&""">编辑</a> - <a href=""admin.asp?act=WeiBodel&id="&Rs("id")&"&Parent=0"">删除</a></p>"
Else
zhengwen="<p>"&Ubbcode(HTMLEncode(Rs("Content")))&"</p><p class=""TextTime"">分享者:<a title=""显示["&Rs("poster")&"]的基本信息"" href=""javascript:memberinfo('"&Rs("Poster")&"')"">"&Rs("Poster")&"</a> | "&Rs("PostTime")&" | <a>回复("&Rs("ReplyNum")&")</a></p>"
End If
End If
Rs.Close

Sql="select * from WeiBo where Parent="&ContextID&" order by id"
OpenRs Sql

'cookie记忆分页模式
listMethod=Request.cookies(systemkey)("listMethod")
AllCount=Rs.RecordCount
If listMethod = "" or listMethod= "page" Then
	Pages=Fix(AllCount/textPageSize)
	If Pages*textPageSize<AllCount Then
		Pages=Pages+1	
	End If
	PastCount=(Page-1)*textPageSize
	If PastCount>=AllCount Then
		Page=1
		PastCount=0
	End If
	If AllCount>PastCount Then
		Rs.Move PastCount
	End If
	ShowCount=Min(AllCount-PastCount,textPageSize)
Else
	ShowCount=AllCount
End If
'分页结束

dim Cnumber
Cnumber=textPageSize*(Page-1)+textPageSize
If Pages=1 Then Cnumber=ShowCount
If Page=Pages and Page>1 Then Cnumber=Allcount
'===================================评论内容
If Allcount>0 Then
For i=1 To ShowCount
If ShowOrder=1 Then
i2=textPageSize*(Page-1)+i
Else
i2= Allcount-(textPageSize*(Page-1)+i-1)
End If
If User_ID="admin" Then
i3="<p>"&Ubbcode(HTMLEncode(Rs("Content")))&"</p><p class=""TextTime""><a title=""显示["&Rs("poster")&"]的基本信息"" href=""javascript:memberinfo('"&Rs("Poster")&"')"">"&Rs("Poster")&"</a> | "&Rs("PostTime")&"&nbsp;|&nbsp;<a href=""admin.asp?act=WeiBoedit&id="&Rs("id")&""">编辑</a> - <a href=""admin.asp?act=WeiBodel&id="&Rs("id")&"&Parent="&Rs("Parent")&""">删除</a></p>"
Else
i3="<p>"&Ubbcode(HTMLEncode(Rs("Content")))&"</p><p class=""TextTime""><a title=""显示["&Rs("poster")&"]的基本信息"" href=""javascript:memberinfo('"&Rs("Poster")&"')"">"&Rs("Poster")&"</a> | "&Rs("PostTime")&"</p>"
End If
TextLieBiao=TextLieBiao+i3
Rs.MoveNext
Next
End If
Rs.Close

'评论翻页
TextFanYe=ShowPages(Pages,Page,"text.asp?id="&contextID&"")

'登录
If User_ID="" Then Logon="<a href=""javascript:logins();"">登录</a> - <a href=""javascript:openreg();"">注册</a>"
If User_ID<>"" Then Logon="<span>【"&User_ID&"】<a href=""javascript:openreg();"">修改资料</a> - <a href=""w.asp?act=login&id="">退出</a></span>"
If User_ID="admin" Then Logon="<span>【"&User_ID&"】<a href=""admin.asp"">管理</a> - <a href=""w.asp?act=login&id="">退出</a></span>"

'评论输入框
If EnableReply="2" And User_ID<>"" Then
plk="<form id=""form1"" name=""form1"" method=""POST"" action=""w.asp?act=TextReply""><input type=""hidden"" name=""artid"" value="""&ContextID&"""><textarea id=""context"" name=""context"" cols=""78"" rows=""6""></textarea><p><input type=""submit"" value=""确认发表"" class=""button""></p>"
Else
plk="<form id=""form1""><textarea id=""context"" cols=""78"" rows=""6""></textarea><p style=""color:#ff0000;"">发表前请先 <a href=""javascript:openreg();"">注册</a> <a href=""javascript:logins();"">登录</a></p>"
End If
If EnableReply="3" Then plk="<form id=""form1"" name=""form1"" method=""POST"" action=""w.asp?act=TextReply""><input type=""hidden"" name=""artid"" value="""&ContextID&"""><textarea id=""context"" name=""context"" cols=""78"" rows=""6""></textarea><p><input type=""submit"" value=""确认发表"" class=""button""></p>"
If EnableReply="1" Then plk=""

OpenTemplate '打开模板文件
Template=replace(Template,"{微博名称}",strSiteName)
Template=replace(Template,"{博主名称}",BoZhuName)
Template=replace(Template,"{微博网址}",strSiteWeiBoUrl)
Template=replace(Template,"{博主简介}",strBanner)
Template=replace(Template,"{主页网址}",strSiteUrl)
Template=replace(Template,"{发表数}",tallcount)
Template=replace(Template,"{评论数}",allreply)
Template=replace(Template,"{会员人数}",membercount)
Template=replace(Template,"{登录}",Logon)
Template=replace(Template,"{微博正文}",ZhengWen)
Template=replace(Template,"{评论列表}",TextLieBiao)
Template=replace(Template,"{评论翻页}",TextFanYe)
Template=replace(Template,"{评论输入框}",plk)
Template=replace(Template,"{友情链接}",friendsite)
Template=replace(Template,"{最新回复}",NewReply)
Template=replace(Template,"{css样式}",strStyle)

Response.write Template

Call CloseAll
%>