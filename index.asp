<!--#include file="include/head.asp"-->
<%
Dim Page,Pages,ShowCount,PastCount,ii,fbk
Page=Max(Cnum(Request("page")),1)
HtmlFiles="index"

'�б�
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
ii="<p>"&Ubbcode(HTMLEncode(Rs("Content")))&"</p><p class=""TextTime"">"&formatdate(Rs("PostTime"),5)&" | <a href=""text.asp?id="&Rs("id")&""">�ظ�("&Rs("ReplyNum")&")</a></p>"
Else
ii="<p>"&Ubbcode(HTMLEncode(Rs("Content")))&"</p><p class=""TextTime"">������:<a title=""��ʾ["&Rs("poster")&"]�Ļ�����Ϣ"" href=""javascript:memberinfo('"&Rs("Poster")&"')"">"&Rs("Poster")&"</a> | "&formatdate(Rs("PostTime"),5)&" | <a href=""text.asp?id="&Rs("id")&""">�ظ�("&Rs("ReplyNum")&")</a></p>"
End If
LieBiao=LieBiao+ii
Rs.MoveNext
Next
Rs.Close

'��ҳ
FanYe=ShowPages(Pages,Page,"")

'��¼
If User_ID="" Then Logon="<a href=""javascript:logins();"">��¼</a> - <a href=""javascript:openreg();"">ע��</a>"
If User_ID<>"" Then Logon="<span>��"&User_ID&"��<a href=""##post"">����΢��</a> - <a href=""javascript:openreg();"">�޸�����</a> - <a href=""w.asp?act=login&id="">�˳�</a></span>"
If User_ID="admin" Then Logon="<span>��"&User_ID&"��<a href=""admin.asp"">����</a> - <a href=""w.asp?act=login&id="">�˳�</a></span>"

'�����
If EnablePost="2" And User_ID<>"" Then
fbk="<form id=""form1"" name=""form1"" method=""POST"" action=""w.asp?act=WeiBopost""><textarea id=""context"" name=""context"" cols=""78"" rows=""6""></textarea><p><input type=""submit"" value=""ȷ�Ϸ���"" class=""button""></p>"
Else
fbk="<form id=""form1""><textarea id=""context"" cols=""78"" rows=""6""></textarea><p style=""color:#ff0000;"">����ǰ���� <a href=""javascript:openreg();"">ע��</a> <a href=""javascript:logins();"">��¼</a></p>"
End If
If EnablePost="3" Then fbk="<form id=""form1"" name=""form1"" method=""POST"" action=""w.asp?act=WeiBopost""><textarea id=""context"" name=""context"" cols=""78"" rows=""6""></textarea><p><input type=""submit"" value=""ȷ�Ϸ���"" class=""button""></p>"
If EnablePost="1" Then fbk=""

OpenTemplate '��ģ���ļ�
Template=replace(Template,"{΢������}",strSiteName)
Template=replace(Template,"{��ҳ��ַ}",strSiteUrl)
Template=replace(Template,"{��������}",BoZhuName)
Template=replace(Template,"{΢����ַ}",strSiteWeiBoUrl)
Template=replace(Template,"{�������}",strBanner)
Template=replace(Template,"{������}",allcount)
Template=replace(Template,"{������}",allreply)
Template=replace(Template,"{��Ա����}",membercount)
Template=replace(Template,"{��¼}",Logon)
Template=replace(Template,"{�б�}",LieBiao)
Template=replace(Template,"{��ҳ}",FanYe)
Template=replace(Template,"{���»ظ�}",NewReply)
Template=replace(Template,"{��������}",friendsite)
Template=replace(Template,"{css��ʽ}",strStyle)
Template=replace(Template,"{���������}",fbk)

Response.write Template

Call CloseAll
%>