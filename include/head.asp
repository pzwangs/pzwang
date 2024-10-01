<%@ CODEPAGE = "936" %>
<%
Option Explicit							'变量定义后可用
Response.CodePage=936
Response.Charset="gb2312" 
Session.CodePage=936

'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 

Dim ConnStr,Conn,Rs,PageStartTime,Data_Path,User_ID,sql,SystemKey,i,AllCount,Act
'定义模板
Dim HtmlFiles,TemplateHtml,Template,Logon,FanYe,LieBiao,friendsite
Dim strSiteName,strSiteUrl,strSiteWeiBoUrl,BoZhuName,strBanner,strStyle,indexPageSize,strBadWords,strSpamIp,strKeyword,Enable_Cookies,Enable_VLCode,textPageSize,EnablePost,EnableReply
PageStartTime=Timer()
SystemKey="RT4H2-8WYHG" '系统标识号，可任意设定
Enable_Cookies=1 '设置允许cookie方式保存登录信息,0关闭,1开启
strBadWords=""        '不良词语过滤，用|号分开
%>
<!--#include file="config.asp"-->
<%
strStyle=Replace(strStyle,".css","")
strStyle=strStyle+".css"

'开始数据库连接
If IsFile(Data_Path)=False Then Data_Path="../" & Data_Path
ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(Data_Path)

'文件说明：通用函数
Function SqlShow(Str)	'去除查询漏洞
	SqlShow=Replace(Str,"'","''")
End Function
'========================================================================
Sub OpenData()			'打开数据库连接
	Set Conn=Server.CreateObject("Adodb.Connection")
	Set Rs=Server.CreateObject("Adodb.Recordset")
	Conn.Open ConnStr
    If Err Then
       err.Clear
       Set conn = Nothing
        response.write("数据库连接出错，请修正文件")
        Response.End
	End If
End Sub
'========================================================================
Sub CloseData()			'关闭数据库连接
	Set Rs=Nothing
	Set Conn=Nothing
End Sub
'========================================================================
Sub CloseAll()			'关闭数据库及数据集
	Set Rs=Nothing
	Conn.Close
	Set Conn=Nothing
End Sub
'========================================================================
Sub OpenRs(ByVal SqlStr)'打开数据集
	If Left(LCase(SqlStr),6)="select" Then
		Rs.Open SqlStr,Conn,1,3
	Else
		Conn.Execute SqlStr
	End If
End Sub
'========================================================================
Sub CheckUser()			'检查用户登录状态
If Enable_Cookies = 1 Then
  If Session(SystemKey & "User_ID")="" Then
	dim userid,userpassword
	userid=Request.cookies(systemkey)("userid")
	userpassword=Request.cookies(systemkey)("userpassword")
	If userid<>"" Then
		sql="select * from users where username='"&userid&"' and userpass='"&userpassword&"' "
		openrs(sql)
		if  rs.EOF and rs.BOF Then
        Response.cookies(systemkey)("userid")=""
		Response.cookies(systemkey)("userpassword")=""
		Session(SystemKey & "User_ID")=""
		else
		Session(SystemKey & "User_ID")=rs("username")
		end if
		Rs.Close
	Else
		Response.cookies(systemkey)("userid")=""
		Response.cookies(systemkey)("userpassword")=""
		Session(SystemKey & "User_ID")=""
	End If
  End If
	User_ID=Session(SystemKey & "User_ID")
Else
	User_ID=Session(SystemKey & "User_ID")
End If
End Sub
'========================================================================
Function Cnum(Num)			'把一个字符变成一个数  
	If IsNumeric(Num) Then
		Cnum=Clng(Num)
	Else
		Cnum=0
	End If
End Function
'========================================================================
Function Max(Num1,Num2)		'求两数大者				
	If Num1>Num2 Then
		Max=Num1
	Else
		Max=Num2
	End If
End Function
'========================================================================
Function Min(Num1,Num2)		'求两数小者			
	If Num1>Num2 Then
		Min=Num2
	Else
		Min=Num1
	End If
End Function
'========================================================================
Function IsFile(tPath)		'判断文件是否存在
	Dim Fso,Path
	Set Fso=CreateObject("Scripting.FileSystemObject")
	If Mid(tPath,2,1)=":" Then
		Path=tPath
	Else
		Path=Server.MapPath(tPath)
	End If
	IsFile=Fso.FileExists(Path)
	Set Fso=Nothing
End Function
'========================================================================
Function ShowPages(Pages,Page,Url)		'分面显示
Dim i,Str,FrontStr,BackStr,ShowStr,StartNum,EndNum
Str=Url
If Replace(Str,"?","")<>Str Then
	Str=Str & "&page="
Else
	Str=Str & "?page="
End If
FrontStr="<a href=""" & Str & 1 & """ title=""第一页""><</a>"
BackStr="<a href=""" & Str & Pages & """ title=""最后一页"">></a>"
If Pages<=1 Then
	ShowPages=""
	Exit Function
End If
If Pages<=8 Then
	For i=1 To Pages
		If i<>Page Then
			ShowPages=ShowPages & "<a href=""" & Str & i & """>" & i & "</a> "
		Else
			ShowPages=ShowPages & "<a class=""tebo"">" & i & "</a>"
		End If
	Next
	ShowPages=FrontStr & " " &  ShowPages & " " & BackStr
	Exit Function
End If
If Pages>8 Then
	StartNum=Page-5
	EndNum=StartNum+7
	If StartNum<=0 Then
		StartNum=1
		EndNum=StartNum+7
	End If
	If EndNum>Pages Then
		EndNum=Pages
		StartNum=EndNum-7
	End If
	For i=StartNum To EndNum
		If i<>Page Then
			If i=Pages Then
				ShowPages=ShowPages & "<a href=""" & Str & Pages & """>" & Pages & "</a>"
				ShowPages=ShowPages & "<a href=""" & Str & Pages & """ title=""最后一页"">></a>"
			Else				
				ShowPages=ShowPages & "<a href=""" & Str & i & """>" & i & "</a> "
			End If
		Else
			If i=Pages Then
				ShowPages=ShowPages & "<a class=""tebo"" title=""最后一页"">" & Pages & "</a>"
			Else
				ShowPages=ShowPages & "<a class=""tebo"">" & i & "</a> "
			End If
		End If
	Next
	ShowPages=FrontStr & " " & ShowPages
	If EndNum<Pages Then
		ShowPages=ShowPages & "...&nbsp;<a href=""" & Str & Pages & """ title=""最后一页"">" & Pages & "</a><a href=""" & Str & Pages & """ title=""最后一页"">></a>"
	End If
End If
End Function
'=========================================================================
'文章中的纯文字
Function delHtml(strHtml) '做了一个函数名叫delhtml
Dim objRegExp, strOutput
Set objRegExp = New Regexp ' 建立正则表达式
objRegExp.IgnoreCase = True ' 设置是否区分大小写
objRegExp.Global = True '是匹配所有字符串还是只是第一个
objRegExp.Pattern = "([[a-zA-Z].*?])|([[\/][a-zA-Z].*?])" ' 设置模式引号中的是正则表达式，用来找出html标签
strOutput = objRegExp.Replace(strHtml, "") '将html标签去掉
strOutput = Replace(strOutput, "[", "[") '防止非html标签不显示
strOutput = Replace(strOutput, "]", "]") 
delHtml = strOutput
Set objRegExp = Nothing
End Function
'========================================================================
'该函数作用：按指定参数格式化显示时间。
'numformat=1:将时间转化为yyyy-mm-dd hh:nn格式。
'numformat=2:将时间转化为yyyy-mm-dd格式。
'numformat=3:将时间转化为hh:nn格式。
'numformat=4:将时间转化为yyyy年mm月dd日 hh时nn分格式。
'numformat=5:将时间转化为yyyy年mm月dd日格式。
'numformat=6:将时间转化为hh时nn分格式。
'numformat=7:将时间转化为yyyy年mm月dd日 星期×格式。
'numformat=8:将时间转化为yymmdd格式。
'numformat=9:将时间转化为mmdd格式。

Function formatdate(shijian,numformat)
Dim ystr,mstr,dstr,hstr,nstr '变量含义分别为年字符串，月字符串，日字符串，时字符串，分字符串

If isnull(shijian) Then
    numformat=0
Else
    ystr=datepart("yyyy",shijian)   
    
    If datepart("m",shijian)>9 Then 
      mstr=datepart("m",shijian)
    Else
      mstr="0"&datepart("m",shijian)  
    End If
  
    If datepart("d",shijian)>9 Then
      dstr=datepart("d",shijian)
    Else
      dstr="0"&datepart("d",shijian)  
    End If
  
    If datepart("h",shijian)>9 Then
      hstr=datepart("h",shijian)
    Else
      hstr="0"&datepart("h",shijian)  
    End If
  
    If datepart("n",shijian)>9 Then
      nstr=datepart("n",shijian)
    Else
      nstr="0"&datepart("n",shijian)  
    End If
End If  
  
Select Case numformat
    Case 0
    formatdate=""
    Case 1
    formatdate=ystr&"-"&mstr&"-"&dstr&" "&hstr&":"&nstr 
    Case 2
    formatdate=ystr&"-"&mstr&"-"&dstr 
    Case 3
    formatdate=hstr&":"&nstr
    Case 4
    formatdate=ystr&"年"&mstr&"月"&dstr&"日 "&hstr&"时"&nstr&"分"
    Case 5
    formatdate=ystr&"年"&mstr&"月"&dstr&"日" 
    Case 6
    formatdate=hstr&"时"&nstr&"分"
    Case 7
    formatdate=ystr&"年"&mstr&"月"&dstr&"日 "&weekdayname(weekday(shijian))
    Case 8
    formatdate=right(ystr,2)&mstr&dstr
    Case 9
    formatdate=mstr&dstr
End Select
End Function
'========================================================================
Function RemoteIp()				'取对方IP地址
	If Request.ServerVariables("HTTP_X_FORWARDED_FOR")<>"" then 
		RemoteIp=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	Else
		RemoteIp=Request.ServerVariables("REMOTE_ADDR")
	End If
End Function
'========================================================================
Function TurnTo(ByVal URl)		'页重定向
	On Error Resume Next
	Rs.Close
	CloseAll
	Response.Clear
	Response.Redirect(URL)
End Function
'=========================================================================
function strlen(str)			'字符串长度，支持中英文混用
dim p_len,xx 
p_len=0 
strlen=0 
if trim(str)<>"" then 
p_len=len(trim(str)) 
for xx=1 to p_len 
if asc(mid(str,xx,1))<0 then 
strlen=int(strlen) + 2 
else 
strlen=int(strlen) + 1 
end if 
next 
end if 
end function 
'================================================================
function strvalue(str,lennum)	'截取字符串，支持中英文混用
dim p_num,x 
dim i 
if strlen(str)<=lennum then 
strvalue=str 
else 
p_num=0 
x=0 
do while not p_num > lennum-2 
x=x+1 
if asc(mid(str,x,1))<0 then 
p_num=int(p_num) + 2 
else 
p_num=int(p_num) + 1 
end if 
strvalue=left(trim(str),x)&"…" 
loop 
end if 
end function 
'=========================================================================
Function CutStr(byVal Str,byVal StrLen) '截取字符串，支持中英文混用
     Dim l,t,c,i 
     l=Len(str) 
     t=0 
     For i=1 To l 
           c=AscW(Mid(str,i,1)) 
           If c<0 Or c>255 Then t=t+2 Else t=t+1 
           IF t>=StrLen Then 
                 CutStr=left(Str,i)&"..." 
                 Exit For 
           Else 
                 CutStr=Str 
           End If 
     Next 
End Function 
'=========================================================================
Function HTMLEncode(reString)			'转换HTML代码
	Dim Str:Str=reString
	If Not IsNull(Str) Then
		Str = UnCheckStr(Str)
		Str = Replace(Str, "&", "&amp;")
		Str = Replace(Str, ">", "&gt;")
		Str = Replace(Str, "<", "&lt;")
		Str = Replace(Str, CHR(34),"&quot;")
		Str = Replace(Str, CHR(39),"&#39;")
		Str = Replace(Str, CHR(13), "")
		Str = Replace(Str, "       ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 0)
		Str = Replace(Str, "      ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 0)
		Str = Replace(Str, "     ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 0)
		Str = Replace(Str, "    ", "&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 0)
		Str = Replace(Str, "   ", "&nbsp;&nbsp;&nbsp;", 1, -1, 0)
		Str = Replace(Str, "  ", "&nbsp;&nbsp;", 1, -1, 0)
		Str = Replace(Str, CHR(10), "<br>")
		Str	= chkBadWords(Str)
		HTMLEncode = Str
	End If
End Function
'==========================================================================
Function UnCheckStr(Str)
		Str = Replace(Str, "sel&#101;ct", "select")
		Str = Replace(Str, "jo&#105;n", "join")
		Str = Replace(Str, "un&#105;on", "union")
		Str = Replace(Str, "wh&#101;re", "where")
		Str = Replace(Str, "ins&#101;rt", "insert")
		Str = Replace(Str, "del&#101;te", "delete")
		Str = Replace(Str, "up&#100;ate", "update")
		Str = Replace(Str, "lik&#101;", "like")
		Str = Replace(Str, "dro&#112;", "drop")
		Str = Replace(Str, "cr&#101;ate", "create")
		Str = Replace(Str, "mod&#105;fy", "modify")
		Str = Replace(Str, "ren&#097;me", "rename")
		Str = Replace(Str, "alt&#101;r", "alter")
		Str = Replace(Str, "ca&#115;t", "cast")
		UnCheckStr=Str
End Function
'==================================================================
Function formatIP(IpAddress)
	dim IpArray,IpAdd,x
	IpArray=split(IpAddress,".",-1,1)
	For x = Lbound(IpArray) to Ubound(IpArray)-1 
		IpAdd=IpAdd+IpArray(x)+"."
	Next 
	formatIP=IpAdd+"*"
end Function
'==================================================================
Function formatIP2(IpAddress)
	dim IpArray,IpAdd,x
	IpArray=split(IpAddress,".",-1,1)
	For x = Lbound(IpArray) to Ubound(IpArray)-2 
		IpAdd=IpAdd+IpArray(x)+"."
	Next 
	formatIP2=IpAdd+"*.*"
end Function
'========================================================================
Sub ShowError(msg,mode)			'显示错误信息
Response.write "<script language='javascript'>alert('"&msg&"');"
If mode=0 Then 
Response.write("window.close();</script>") 
Response.end
Else 
Response.write("history.go(-"&mode&");</script>")
Response.end
End If
End Sub
'========================================================================
Function CCEmpty(StrData)			'置空
	Dim Str
	Str=StrData
	Str=Replace(Str," ","")
	Str=Replace(Str,"　","")
	CCEmpty=Str
End Function
'========================================================================
'判断是否包含特殊字符
Function IsSpecial(StrData)			
	Dim Str,Fy_In,Fy_Inf,Fy_Xh
	Str=StrData
	Fy_In = "'|;|and|(|)|exec|insert|select|delete|update|count|*|%|$|@|!|(|)|+|=|&|/|\|chr|mid|master|truncate|char|declare" 
	Fy_Inf = split(Fy_In,"|") 
    For Fy_Xh = 0 To Ubound(Fy_Inf)
	If Instr(LCase(Str),Fy_Inf(Fy_Xh)) <> 0 Then
	IsSpecial=Ture
	Exit For
	Else
	IsSpecial=False
	End If 
	Next
End Function
'========================================================================
Function Printl(Str)				'行式输出
	If IsNull(Str) Then
		Str=""
	End If
	Response.Write Str & vbcrlf
End Function
'========================================================================
Function Print(Str)					'输出
	If IsNull(Str) Then
		Str=""
	End If
	Response.Write Str
End Function
'========================================================================
' 过滤不良词语的功能
Function ChkBadWords(fString)
	dim strBadWords2,n
	If strBadWords<>"" Then
      strBadWords2 = split(strBadWords, "|")   'BadWords   是数据库中定义的过滤词语
        for n = 0 to ubound(strBadWords2)
            fString = Replace(fString, strBadWords2(n), string(len(strBadWords2(n)),"*")) 
        next
   End if 
    ChkBadWords = fString
End Function

OpenData
Call CheckUser()							
'========================================================================
'UBB
Function UBBCode(strContent)
Dim re
Set re=new RegExp
re.IgnoreCase =True
re.Global=True

'颜色
re.Pattern="\[color=([^<>\]]*?)\](.*?)\[\/color]"
strContent=re.Replace(strContent,"<font color=""$1"">$2</font>")
'表情
re.Pattern="\[emot=default,(.*?)\/]"
strContent=re.Replace(strContent,"<img src=""img/emot/$1.gif"">")	
'表情1.0版本
re.Pattern="\{:em*([0-9]*)}"
strContent=re.Replace(strContent,"<img src=""img/emot/smile.gif"">")
'链接
re.Pattern = "\[url=(.[^\]]*)\](.*?)\[\/url]"
strContent = re.Replace(strContent,"<a href=""$1"" target=""_blank"">$2</a>")
re.Pattern = "\[[url](.*?)\[\/[url]"
strContent = re.Replace(strContent,"<a href=""$1"" target=""_blank"">$1</a>")
'图片
re.Pattern="\[img\](.*?)\[\/img]"
strContent=re.Replace(strContent,"<p><a class=""continue_reading_link zoom"" href=""$1"" ><img src=""$1"" width=""120"" title=""点击查看大图"" style=""border:#d2d2d2 1px solid;padding:2px;""></a></p>")	
'图片带文字
re.Pattern = "\[img=(.[^\]]*)\](.*?)\[\/img]"
strContent=re.Replace(strContent,"<p><a class=""continue_reading_link zoom"" href=""$2"" ><img src=""$2"" onload=""javascript:DrawImage(this);"" width=""120"" title=""$1"" style=""border:#d2d2d2 1px solid;padding:2px;""></a></p>")	
'视频
re.Pattern="\[flash\s*(?:=\s*(\d+)\s*,\s*(\d+)\s*)\]([\s\S]+?)\[\/flash\]"
strContent=re.Replace(strContent,"<p><embed type=""application/x-shockwave-flash"" wmode=""transparent"" src=""$3"" width=""400"" height=""300"" pluginspage=""http://www.macromedia.com/go/getflashplayer""></embed></p>")	
'音乐
re.Pattern="\[media\s*=\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\]([\s\S]+?)\[\/media\]"
strContent=re.Replace(strContent,"<p><embed type=""application/x-mplayer2"" classid=""clsid:6bf52a52-394a-11d3-b153-00c04f79faa6"" src=""$4"" enablecontextmenu=""false"" autostart=""0"" width=""300"" height=""68""></object></p>")

UBBCode=strContent
End Function
'===========================================================================
'友情链接
sql="select * from friendsite order by siteorder"
OpenRs(Sql)
Dim ff
For i=1 To Rs.RecordCount
ff="<li><a target=""_blank"" href="""&Trim(Rs("siteurl"))&""" title="""&Trim(Rs("sitename"))&""">"&Trim(Rs("sitename"))&"</a></li>"
friendsite=friendsite+ff
Rs.MoveNext
Next
Rs.Close
'========================================================================
'最新回复
Sql="Select Top 8 WeiBo.Content,WeiBo.Poster,WeiBo.parent From WeiBo where parent<>0 order by PostTime desc"
OpenRs(Sql)
For i=1 To Rs.RecordCount
Dim Content,iii,NewReply
Content=delHtml(Rs("Content"))
If Content="" Then Content="{表情}"
If strlen(Content)>50 Then Content=CutStr(Content,50)
iii="<li><a title=""显示["&Rs("poster")&"]的基本信息"" href=""javascript:memberinfo('"&Rs("Poster")&"')"">"&Rs("Poster")&"</a>：<a href=""text.asp?id="&Rs("parent")&""">"&Content&"</a></li>"
NewReply=NewReply+iii
Rs.MoveNext
Next
Rs.Close
'========================================================================
'粉丝
Dim membercount
sql="select count(*) from Users"
openrs(sql)
membercount=rs(0)
rs.close
'========================================================================
'博文数
Dim tallcount
sql="select count(*) from WeiBo where parent=0"
openrs(sql)
tallcount=rs(0)
rs.close
'========================================================================
'回复数
Dim allreply
sql="select count(*) from WeiBo where parent<>0"
openrs(sql)
allreply=rs(0)
rs.close
'========================================================================
'打开模板文件
Function OpenTemplate 
Set Rs = Server.CreateObject("Scripting.FileSystemObject")
Set TemplateHtml = Rs.OpenTextFile(Server.MapPath("include/"&HtmlFiles&".html"))
Template = TemplateHtml.ReadAll
End Function 
%>
