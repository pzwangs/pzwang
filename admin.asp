<!--#include file="include/head.asp"-->
<!-- #include file=include/config.asp -->
<%
Dim id,filename,Parent
If User_ID<>"admin" Then Call ShowError("��δ��¼��",1)
act=Request.QueryString("act")
If act="" Then act="post"
%>
<html>
<head>
<meta name="Robots" Content="noindex,nofollow"> 
<meta http-equiv="Content-Type" Content="text/html; charset=gb2312" />
<title><%=strSiteName%></title>
<link href="img/<%=strStyle%>" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="img/jquery.js"></script> 
<script type="text/javascript" src="img/xheditor.js"></script>
<script type="text/javascript">$(pageInit);function pageInit(){$('#context').xheditor({tools:'mini',upLinkUrl:"upload.asp",upLinkExt:"zip,rar,txt",upImgUrl:"upload.asp",upImgExt:"jpg,jpeg,gif,png",upFlashUrl:"upload.asp",upFlashExt:"swf",upMediaUrl:"upload.asp",upMediaExt:"avi",beforeSetSource:ubb2html,beforeGetSource:html2ubb,emotMark:true,cleanPaste:3});}function submitForm(){$('#form1').submit();}</script>
<body>

<div id="Topbg"><h1><a href="<%=strSiteWeiBoUrl%>" id="logo"><%=strSiteName%></a></h1><p><%=strSiteWeiBoUrl%></p></div>

<div id="Wrapper">
    <div id="HouTaiZuo">
	     <%If act="post" Then Call FaBiao%>
	     <%If act="Config" Then Call WangZhanSheZhi%>
	     <%If act="friendsite" Then Call YouQingLianJie%>
	     <%If act="WeiBoedit" Then Call BianJiWeiBo%>
	     <%If act="WeiBodel" Then Call ShanChuWeiBo%>
	     <%If act="YongHulist" Then Call YongHuGuanLi%>
    </div>
	<div id="side">
	     <div id="HouTaiYou">
		      <div id="HouTaiYouT"><%Response.write("<a href=""index.asp"">������ҳ</a> - <a href=""javascript:openreg();"">�޸�����</a> - <a href=""w.asp?act=login&id="">�˳�</a>")%></div>
		      <p><a href="admin.asp?act=post">����΢��</a></p>
		      <p><a href="admin.asp?act=Config">��������</a></p>
		      <p><a href="admin.asp?act=YongHulist">�û�����</a></p>
		      <p><a href="admin.asp?act=friendsite">��������</a></p>
		 </div>
    </div>
</div>

<div id="Copyright">Copyright &copy; 2011 <a href="index.asp"><%=strSiteName%></a></div>
<div class="gotop"><button onclick="$('html,body').animate({scrollTop:0},300);return false;" title="���ض���"><span>���ض���</span></button></div>
</body>
</html>

<%Function YouQingLianJie%>
<div id="YouQingLianJie">
<form name="friendsitenew" method="POST" action="admin.asp?act=friendsitenew">
<%
sql="select * from friendsite order by siteorder"
openrs(sql)
For i=1 To Rs.RecordCount
%>
<p><%=rs("siteorder")%>&nbsp;<%=rs("sitename")%><a href="<%=rs("siteurl")%>" target="_blank"><%=rs("siteurl")%></a>��<a href="admin.asp?act=friendsitedelete&id=<%=rs("siteorder")%>">ɾ��</a>��</p>
<%
rs.movenext
Next
%>
<p><input type="text" name="siteorder" size="3" maxlength="4" class="input" ><input type="text" name="sitename" size="15" maxlength="30" class="input"><input type="text" name="siteurl" size="38" maxlength="100" class="input" value="http://"><input type="submit" value="ȷ�����" name="addurl" class="button"></p>
</form>
</div>
<%End Function%>

<%Sub FaBiao%>
<div id="FaBiaoWeiBo">
<form id="form1" name="form1" method="POST" action="admin.asp?act=WeiBopost">
<textarea id="context" name="context" cols="78" rows="10"></textarea>
<p><input type="submit" value="ȷ�Ϸ���" class="button"></p>
</div>
<%End Sub%>

<%
'�༭
Sub BianJiWeiBo
Dim ContextID
ContextID=Cnum(Request.QueryString("id"))
Sql="select * from WeiBo where id="&ContextID&" order by id"
OpenRs Sql
%>
<div id="FaBiaoWeiBo">
<form id="form1" name="form1" method="POST" action="admin.asp?act=WeiBoUpdate">
<input type="hidden" name="id" value="<%=rs("id")%>">
<textarea rows="10" name="Context" cols="78" id="context"><%=Rs("Content")%></textarea>
<p><input type="submit" value="ȷ�Ϸ���" class="button"></p>
</div>
<%End Sub%>

<%
'�û�����
Sub YongHuGuanLi
Response.write("<div id=""YongHuGuanLi""><ul>")
Sql="select * from Users order by id desc"
OpenRs Sql
For i=1 To Rs.RecordCount
Response.write("<li><a title=""�鿴�༭["&Rs("UserName")&"]�Ļ�����Ϣ"" href=""javascript:memberinfo('"&Rs("UserName")&"')"">"&rs("UserName")&"</a></li>")
rs.movenext
Next
Response.write("</ul></div>")
End Sub
%>

<%
'����΢��
If act="WeiBopost" Then
Sql="select * from WeiBo"
OpenRs(Sql)
Rs.AddNew
Rs("Content")=Request.Form("Context")
Rs("Poster")=User_ID
Rs("PostTime")=Now
Rs("Ip")=RemoteIP()
Rs("Parent")=0
Rs.Update
Rs.Close
Response.write("<script language='javascript'>alert('����ɹ���');window.location='index.asp';</script>")
End If
%>

<%
'�༭΢��
If act="WeiBoUpdate" Then
sql="select * from [WeiBo] where id=" &request("id")
Set rs=Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,3,3
rs("Content") =Request.Form("Context")
rs.update
rs.close
Response.write("<script language='javascript'>alert('�༭�ɹ���');window.location='index.asp';</script>")
End If
%>

<%
'�½���������
If act="friendsitenew" Then
dim siteorder,sitename,siteurl
siteorder=trim(request.form("siteorder"))
sitename=trim(request.form("sitename"))
siteurl=trim(request.form("siteurl"))
sql="select * from friendsite where siteorder="&siteorder&""
OpenRs(sql)
If Rs.recordcount>0 Then Call ShowError("ID����������ID�ظ���",1)
rs.close
'����������
Sql="select * from friendsite where siteorder=null"
openrs sql
rs.addnew
rs("siteorder")=siteorder
rs("sitename")=sitename
rs("siteurl")=siteurl
rs.update
Turnto "admin.asp?act=friendsite"
End If
%>

<%
'ɾ����������
If act="friendsitedelete" Then
id=request.querystring("id")
Sql="delete * from friendsite where siteorder="&id&""
OpenRs(Sql)
Turnto "admin.asp?act=friendsite"
End If
%>

<%Function WangZhanSheZhi%>
<div id="WangZhanSheZhi">
<form name="form1" method="post" action="admin.asp?act=ConfigEdit">
<p>΢�����ƣ�<input name="strSiteName" type="text" id="strSiteName" value="<%=strSiteName%>" size="16"></p>
<p>��ҳ��ַ��<input name="strSiteUrl" type="text" id="strSiteUrl" value="<%=strSiteUrl%>" size="38"></p>
<p>΢����ַ��<input name="strSiteWeiBoUrl" type="text" id="strSiteWeiBoUrl" value="<%=strSiteWeiBoUrl%>" size="38"></p>
<p>�������ƣ�<input name="strBoZhuName" type="text" id="strBoZhuName" value="<%=bozhuname%>" size="16"></p>
<p>������飺<input name="strBanner" type="text" id="strBanner" value="<%=strBanner%>" size="68"></p>
<p>������ã�<input name="strStyle" type="text" id="strStyle" value="<%=strStyle%>" size="16"></p>
<p>��ҳ��ҳ����<input name="indexPageSize" type="text" id="indexPageSize" value="<%=indexPageSize%>" size="1"></p>
<p>���۷�ҳ����<input name="textPageSize" type="text" id="textPageSize" value="<%=textPageSize%>" size="1"></p>
<p>���ݿ�·����<input name="Data_Path" type="text" id="Data_Path" value="<%=Data_Path%>" size="35"></p>
<p>����Ȩ�ޣ�<input name="EnablePost" type="text" id="EnablePost" value="<%=EnablePost%>" size="1"> ˵����1������2��Ա��3�ο�</p>
<p>�ظ�Ȩ�ޣ�<input name="EnableReply" type="text" id="EnableReply" value="<%=EnableReply%>" size="1"> ˵����1������2��Ա��3�ο�</p>
<p><input type="submit" value="ȷ���޸�" class="button"></p>
</form>
</div>
<%End Function%>

<%If act="ConfigEdit" Then
Dim urllist,fso,fout
'��ȡ�ύ����
strSiteName=request.form("strSiteName")
strSiteUrl=request.form("strSiteUrl")
strSiteWeiBoUrl=request.form("strSiteWeiBoUrl")
BoZhuName=request.form("strBoZhuName")
strBanner=request.form("strBanner")
strStyle=request.form("strStyle")
strStyle=Replace(strStyle,".css","")
strStyle=strStyle+".css"
indexPageSize=request.form("indexPageSize")
textPageSize=request.form("textPageSize")
Data_Path=request.form("Data_Path")
EnablePost=request.form("EnablePost")
EnableReply=request.form("EnableReply")
'���������б� 
urllist=urllist & chr(60) & "%" & VbCrLf
urllist=urllist & "strSiteName=" & chr(34) & ""&strSiteName&"" & chr(34) &VbCrLf &VbCrLf
urllist=urllist & "strSiteUrl=" & chr(34) & ""&strSiteUrl&"" & chr(34) & VbCrLf &VbCrLf
urllist=urllist & "strSiteWeiBoUrl=" & chr(34) & ""&strSiteWeiBoUrl&"" & chr(34) & VbCrLf &VbCrLf
urllist=urllist & "BoZhuName=" & chr(34) & ""&BoZhuName&"" & chr(34) & VbCrLf &VbCrLf
urllist=urllist & "strBanner=" & chr(34) & ""&strBanner&"" & chr(34) & VbCrLf &VbCrLf
urllist=urllist & "strStyle=" & chr(34) & ""&strStyle&"" & chr(34) & VbCrLf &VbCrLf
urllist=urllist & "indexPageSize=" &indexPageSize& VbCrLf &VbCrLf
urllist=urllist & "textPageSize=" &textPageSize& VbCrLf &VbCrLf
urllist=urllist & "Data_Path=" & chr(34) & ""&Data_Path&"" & chr(34) & VbCrLf &VbCrLf
urllist=urllist & "EnablePost=" &EnablePost& VbCrLf &VbCrLf
urllist=urllist & "EnableReply=" &EnableReply& VbCrLf &VbCrLf
urllist=urllist & "%" & chr(62) & VbCrLf

filename="include/config.asp"

Set fso = Server.CreateObject("Scripting.FileSystemObject")
Set fout = fso.Createtextfile(server.mappath(filename),true)
fout.writeline urllist
fout.close 
Response.Write"<script language=JavaScript>alert(""��ϲ.�޸ĳɹ�"");window.location='admin.asp?act=Config'</script>"
End If
%>

<%
'ɾ��΢��
Sub ShanChuWeiBo
id=Request.QueryString("id")
Parent=Request.QueryString("Parent")
%>
<div id="postbg">
<form name="form1" method="POST" action="admin.asp?act=delWeiBo">
<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="Parent" value="<%=Parent%>">
<p><input name="yesno" checked="checked" value="yes" type="radio"><label for="header_article">YES</label>&nbsp;&nbsp;<input name="yesno" value="no" type="radio"><label for="header_photo">NO</label> <input type="submit" value="ȷ��ɾ��" class="button"></p>
</div>
<%End Sub%>

<%
If act="delWeiBo" Then
If Request.Form("yesno")="no" Then Call ShowError("�㶼ѡNO�ˣ���Ҫɾ����",1)
If Request.Form("Parent")<>"0"  Then
Sql="select * from WeiBo where id=" &request("Parent")
OpenRs(sql)
'��������΢���Ļظ���
Rs("ReplyNum")=Rs("ReplyNum")-1
Rs.Update
Rs.Close
End If
sql="delete * from WeiBo where id=" &request("id")&" Or Parent=" &request("id")
OpenRs(sql)
Response.write("<script language='javascript'>alert('ɾ���ɹ���');window.location='index.asp';</script>")
End If
%>

<%
'ɾ���û�
If act="delUsers" Then 
sql="delete * from Users where id=" &request("id")
OpenRs(sql)
Response.write("<script language='javascript'>alert('ɾ���ɹ���');window.close();</script>")
End If
%>

<%
'�����û�����
If act="initpass" Then
Sql="select * from Users where id=" &request("id")
OpenRs(sql)
'��������΢���Ļظ���
Rs("UserPass")="49ba59abbe56e057"
Rs.Update
Rs.Close
Response.write("<script language='javascript'>alert('�������óɹ�����ʼ����Ϊ123456��');window.close();</script>")
End If
%>

<%Call CloseAll%>