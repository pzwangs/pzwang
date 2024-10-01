<!--#include file="include/head.asp"-->
<%
Dim Page,Pages,ShowCount,PastCount,ii,fbk
Page=Max(Cnum(Request("page")),1)
HtmlFiles="index"

'列表
Sql="select * from WeiBo where parent=0 order by PostTime desc"
OpenRs(Sql)
AllCount=Rs.RecordCount
Pages=Fix(AllCount/indexPageSize)
If Pages*indexPageSize<AllCount Then Pages=Pages+1	
PastCount=(Page-1)*indexPageSize
If PastCount>=AllCount Then
Page=1
PastCount=0
End If
If AllCount>PastCount Then Rs.Move PastCount
ShowCount=Min(AllCount-PastCount,indexPageSize)
For i=1 To ShowCount
If EnablePost="1" Then
ii="<p>"&Ubbcode(HTMLEncode(Rs("Content")))&"</p><p class=""TextTime"">"&formatdate(Rs("PostTime"),5)&" | <a href=""text.asp?id="&Rs("id")&""">回复("&Rs("ReplyNum")&")</a></p>"
Else
ii="<p>"&Ubbcode(HTMLEncode(Rs("Content")))&"</p><p class=""TextTime"">分享者:<a title=""显示["&Rs("poster")&"]的基本信息"" href=""javascript:memberinfo('"&Rs("Poster")&"')"">"&Rs("Poster")&"</a> | "&formatdate(Rs("PostTime"),5)&" | <a href=""text.asp?id="&Rs("id")&""">回复("&Rs("ReplyNum")&")</a></p>"
End If
LieBiao=LieBiao+ii
Rs.MoveNext
Next
Rs.Close

'翻页
FanYe=ShowPages(Pages,Page,"")

'登录
If User_ID="" Then Logon="<a href=""javascript:logins();"">登录</a> - <a href=""javascript:openreg();"">注册</a>"
If User_ID<>"" Then Logon="<span>【"&User_ID&"】<a href=""##post"">发表微博</a> - <a href=""javascript:openreg();"">修改资料</a> - <a href=""w.asp?act=login&id="">退出</a></span>"
If User_ID="admin" Then Logon="<span>【"&User_ID&"】<a href=""admin.asp"">管理</a> - <a href=""w.asp?act=login&id="">退出</a></span>"

'发表框
If EnablePost="2" And User_ID<>"" Then
fbk="<form id=""form1"" name=""form1"" method=""POST"" action=""w.asp?act=WeiBopost""><textarea id=""context"" name=""context"" cols=""78"" rows=""6""></textarea><p><input type=""submit"" value=""确认发表"" class=""button""></p>"
Else
fbk="<form id=""form1""><textarea id=""context"" cols=""78"" rows=""6""></textarea><p style=""color:#ff0000;"">发表前请先 <a href=""javascript:openreg();"">注册</a> <a href=""javascript:logins();"">登录</a></p>"
End If
If EnablePost="3" Then fbk="<form id=""form1"" name=""form1"" method=""POST"" action=""w.asp?act=WeiBopost""><textarea id=""context"" name=""context"" cols=""78"" rows=""6""></textarea><p><input type=""submit"" value=""确认发表"" class=""button""></p>"
If EnablePost="1" Then fbk=""

OpenTemplate '打开模板文件
Template=replace(Template,"{微博名称}",strSiteName)
Template=replace(Template,"{主页网址}",strSiteUrl)
Template=replace(Template,"{博主名称}",BoZhuName)
Template=replace(Template,"{微博网址}",strSiteWeiBoUrl)
Template=replace(Template,"{博主简介}",strBanner)
Template=replace(Template,"{发表数}",allcount)
Template=replace(Template,"{评论数}",allreply)
Template=replace(Template,"{会员人数}",membercount)
Template=replace(Template,"{登录}",Logon)
Template=replace(Template,"{列表}",LieBiao)
Template=replace(Template,"{翻页}",FanYe)
Template=replace(Template,"{最新回复}",NewReply)
Template=replace(Template,"{友情链接}",friendsite)
Template=replace(Template,"{css样式}",strStyle)
Template=replace(Template,"{发表输入框}",fbk)

Response.write Template

Call CloseAll
%>