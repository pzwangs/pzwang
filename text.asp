<!--#include file="include/head.asp"-->
<%
Dim ContextID,formID,i2,title,Page,Pages,ShowCount,PastCount,listMethod,listpage,j,showOrder,Poster,TextFanYe,ZhengWen,TextLieBiao,i3,plk
ContextID=Cnum(Request.QueryString("id"))
Page=Max(Cnum(sqlshow(Request("page"))),1)
listpage=sqlshow(Request.QueryString("listpage"))
HtmlFiles="text"

Sql="select * from WeiBo where id="&ContextID&" or Parent="&ContextID&" order by Parent,PostTime"
OpenRs(Sql)

'��ʾ����
If EnablePost="1" Then
If User_ID="admin" Then
zhengwen="<p>"&Ubbcode(HTMLEncode(Rs("Content")))&"</p><p class=""TextTime"">"&Rs("PostTime")&" | <a>�ظ�("&Rs("ReplyNum")&")</a> <a href=""admin.asp?act=WeiBoedit&id="&Rs("id")&""">�༭</a> - <a href=""admin.asp?act=WeiBodel&id="&Rs("id")&"&Parent=0"">ɾ��</a></p>"
Else
zhengwen="<p>"&Ubbcode(HTMLEncode(Rs("Content")))&"</p><p class=""TextTime"">"&Rs("PostTime")&" | <a>�ظ�("&Rs("ReplyNum")&")</a></p>"
End If
Else
If User_ID="admin" Then
zhengwen="<p>"&Ubbcode(HTMLEncode(Rs("Content")))&"</p><p class=""TextTime"">������:<a title=""��ʾ["&Rs("poster")&"]�Ļ�����Ϣ"" href=""javascript:memberinfo('"&Rs("Poster")&"')"">"&Rs("Poster")&"</a> | "&Rs("PostTime")&" | <a>�ظ�("&Rs("ReplyNum")&")</a> <a href=""admin.asp?act=WeiBoedit&id="&Rs("id")&""">�༭</a> - <a href=""admin.asp?act=WeiBodel&id="&Rs("id")&"&Parent=0"">ɾ��</a></p>"
Else
zhengwen="<p>"&Ubbcode(HTMLEncode(Rs("Content")))&"</p><p class=""TextTime"">������:<a title=""��ʾ["&Rs("poster")&"]�Ļ�����Ϣ"" href=""javascript:memberinfo('"&Rs("Poster")&"')"">"&Rs("Poster")&"</a> | "&Rs("PostTime")&" | <a>�ظ�("&Rs("ReplyNum")&")</a></p>"
End If
End If
Rs.Close

Sql="select * from WeiBo where Parent="&ContextID&" order by id"
OpenRs Sql

'cookie�����ҳģʽ
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
'��ҳ����

dim Cnumber
Cnumber=textPageSize*(Page-1)+textPageSize
If Pages=1 Then Cnumber=ShowCount
If Page=Pages and Page>1 Then Cnumber=Allcount
'===================================��������
If Allcount>0 Then
For i=1 To ShowCount
If ShowOrder=1 Then
i2=textPageSize*(Page-1)+i
Else
i2= Allcount-(textPageSize*(Page-1)+i-1)
End If
If User_ID="admin" Then
i3="<p>"&Ubbcode(HTMLEncode(Rs("Content")))&"</p><p class=""TextTime""><a title=""��ʾ["&Rs("poster")&"]�Ļ�����Ϣ"" href=""javascript:memberinfo('"&Rs("Poster")&"')"">"&Rs("Poster")&"</a> | "&Rs("PostTime")&"&nbsp;|&nbsp;<a href=""admin.asp?act=WeiBoedit&id="&Rs("id")&""">�༭</a> - <a href=""admin.asp?act=WeiBodel&id="&Rs("id")&"&Parent="&Rs("Parent")&""">ɾ��</a></p>"
Else
i3="<p>"&Ubbcode(HTMLEncode(Rs("Content")))&"</p><p class=""TextTime""><a title=""��ʾ["&Rs("poster")&"]�Ļ�����Ϣ"" href=""javascript:memberinfo('"&Rs("Poster")&"')"">"&Rs("Poster")&"</a> | "&Rs("PostTime")&"</p>"
End If
TextLieBiao=TextLieBiao+i3
Rs.MoveNext
Next
End If
Rs.Close

'���۷�ҳ
TextFanYe=ShowPages(Pages,Page,"text.asp?id="&contextID&"")

'��¼
If User_ID="" Then Logon="<a href=""javascript:logins();"">��¼</a> - <a href=""javascript:openreg();"">ע��</a>"
If User_ID<>"" Then Logon="<span>��"&User_ID&"��<a href=""javascript:openreg();"">�޸�����</a> - <a href=""w.asp?act=login&id="">�˳�</a></span>"
If User_ID="admin" Then Logon="<span>��"&User_ID&"��<a href=""admin.asp"">����</a> - <a href=""w.asp?act=login&id="">�˳�</a></span>"

'���������
If EnableReply="2" And User_ID<>"" Then
plk="<form id=""form1"" name=""form1"" method=""POST"" action=""w.asp?act=TextReply""><input type=""hidden"" name=""artid"" value="""&ContextID&"""><textarea id=""context"" name=""context"" cols=""78"" rows=""6""></textarea><p><input type=""submit"" value=""ȷ�Ϸ���"" class=""button""></p>"
Else
plk="<form id=""form1""><textarea id=""context"" cols=""78"" rows=""6""></textarea><p style=""color:#ff0000;"">����ǰ���� <a href=""javascript:openreg();"">ע��</a> <a href=""javascript:logins();"">��¼</a></p>"
End If
If EnableReply="3" Then plk="<form id=""form1"" name=""form1"" method=""POST"" action=""w.asp?act=TextReply""><input type=""hidden"" name=""artid"" value="""&ContextID&"""><textarea id=""context"" name=""context"" cols=""78"" rows=""6""></textarea><p><input type=""submit"" value=""ȷ�Ϸ���"" class=""button""></p>"
If EnableReply="1" Then plk=""

OpenTemplate '��ģ���ļ�
Template=replace(Template,"{΢������}",strSiteName)
Template=replace(Template,"{��������}",BoZhuName)
Template=replace(Template,"{΢����ַ}",strSiteWeiBoUrl)
Template=replace(Template,"{�������}",strBanner)
Template=replace(Template,"{��ҳ��ַ}",strSiteUrl)
Template=replace(Template,"{������}",tallcount)
Template=replace(Template,"{������}",allreply)
Template=replace(Template,"{��Ա����}",membercount)
Template=replace(Template,"{��¼}",Logon)
Template=replace(Template,"{΢������}",ZhengWen)
Template=replace(Template,"{�����б�}",TextLieBiao)
Template=replace(Template,"{���۷�ҳ}",TextFanYe)
Template=replace(Template,"{���������}",plk)
Template=replace(Template,"{��������}",friendsite)
Template=replace(Template,"{���»ظ�}",NewReply)
Template=replace(Template,"{css��ʽ}",strStyle)

Response.write Template

Call CloseAll
%>