<!-- #include file="Config.asp"-->
<!-- #include file="about.asp" -->
<script language="JavaScript" src="css/page.js"></script>
<!--#include file="head.asp"-->
<!-- #include file="top.asp"-->
<%
Dim locat,LouHB_Key,LouHB_Type,urlstr,sqlstr,hiddenstr
locat=Trim(Request.QueryString("locat"))
urlstr=""
sqlstr=""
hiddenstr=""
if locat="ok" Then
	LouHB_Key=Trim(Request.QueryString("LouHB_Key"))
	LouHB_Type=Trim(Request.QueryString("LouHB_Type"))
	urlstr=urlstr&"LouHB_Key="&LouHB_Key&"&LouHB_Type="&LouHB_Type&"&"
	hiddenstr=hiddenstr&"<input type='hidden' name='LouHB_Key' value='"&LouHB_Key&"'>"
	hiddenstr=hiddenstr&"<input type='hidden' name='LouHB_Type' value='"&LouHB_Type&"'>"
	Select Case LCase(LouHB_Type)
	Case "xy_user"
		sqlstr=" where xy_user like '%"&LouHB_Key&"%'"
	Case "xy_dz"
		sqlstr=" where xy_dz like '%"&LouHB_Key&"%'"
	End Select
End if

page = CLng(request("page"))							'利用CLng函数把page值转换为Long型
judge=request("judge")
judge2=request("judge2")
styp=Trim(Request("styp"))
adtyp=Trim(Request("adtyp"))

judge3=0
styp_title=""
if adtyp="1" Then adtyp=1 Else adtyp=0
if adtyp=1 Then
	styp_title="(降)"
Else
	styp_title="(升)"
End if

orderBy="a.xy_admin"

s_u_t=""
s_u_p=""
s_u_dz=""
s_u_ktrq=""
s_u_xfrq=""
s_u_xfje=""
s_u_jie=""
s_u_dqrq=""
if styp<>"" Then
	urlstr=urlstr&"styp="&styp&"&adtyp="&adtyp&"&"
	select Case LCase(styp)
	Case "user"
		orderBy="xy_user"
		s_u_t=styp_title
	case "pass"
		orderBy="xy_password"
		s_u_p=styp_title
	Case "dz"
		orderBy="xy_dz"
		s_u_dz=styp_title
	Case "ktrq"
		orderBy="xy_ktrq"
		s_u_ktrq=styp_title
	Case "xfrq"
		orderBy="xy_xfrq"
		s_u_xfrq=styp_title
	Case "xfje"
		orderBy="xy_xfje"
		s_u_xfje=styp_title
	Case "zje"
		orderBy="xy_zje"
		s_u_zje=styp_title
	Case "dqrq"
		orderBy="xy_dqrq"
		s_u_dqrq=styp_title
	End Select
End if
if adtyp=1 Then
	orderBy=orderBy&" Desc "
	new_adtyp=0
Else
	orderBy=orderBy&" Asc "
	new_adtyp=1
End if

if styp<>"" Then hiddenstr=hiddenstr&"<input type='hidden' name='styp' value='"&styp&"'>"
if adtyp<>"" Then hiddenstr=hiddenstr&"<input type='hidden' name='adtyp' value='"&adtyp&"'>"
'delete
if request("doAct")="delete" then
ID=Request("LouHB_ID")
L_Table=Request("LouHB_Table")
Del="delete from xy_data where ID="&ID
Louconn.Execute Del
LouConn.close
Set LouConn=Nothing
response.Redirect("index.asp")
end if
%>

<table width="768" border="0" cellspacing="0" cellpadding="0" align="center">
<form name="form1" method="get" action="index.asp">
<input type="hidden" name="locat" value="ok">
<tr>
<td width="283">&nbsp;</td>
<td width="160">
<input type="text" name="LouHB_Key" size="20" value="<%=LouHB_Key%>">
</td>
<td width="64">
<select name="LouHB_Type" type="text" size=1 class="button1">
<option value="xy_user" <%if LouHB_Type="xy_user" Then Response.Write "selected"%>>账号</option>
<option value="xy_dz" <%if LouHB_Type="xy_dz" Then Response.Write "selected"%>>地址</option>
</select>
</td>
<td width="261"><input type="image" src="images/Search.gif" name="button" id="button" value="提交"></td>
</tr>
</form>
</table>
<%
Set Lours=server.createobject("adodb.recordset")
sql="select b.x_mz,a.* from xy_data a Left join xy_admin b on a.xy_admin = b.x_user "&sqlstr&" order by "&orderBy&""
Lours.open sql,Louconn,1,1
if Lours.EOF and Lours.BOF then
Response.Write "<center><br>查无数据</center>"
'response.write("<center><br>数据库中尚未添加任何个人电脑信息</center>")
else
%>
<script language="javascript">
function lookpass(uid,upass){
	var upassid=document.getElementById("upass_"+uid);
	upassid.innerHTML=upass;
}
</script>
<table width="768" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="50" height="30" align="center" background="images/bg.gif"><a href="index.asp?styp=user&adtyp=<%=new_adtyp%>">账号<%=s_u_t%></a></td>
<td width="120" align="center" background="images/bg.gif"><a href="index.asp?styp=pass&adtyp=<%=new_adtyp%>">密码<%=s_u_p%></a></td>
<td width="104" align="center" background="images/bg.gif"><a href="index.asp?styp=dz&adtyp=<%=new_adtyp%>">地址<%=s_u_dz%></a></td>
<td width="75" align="center" background="images/bg.gif"><a href="index.asp?styp=ktrq&adtyp=<%=new_adtyp%>">开通日期<%=s_u_ktrq%></a></td>
<td width="75" align="center" background="images/bg.gif"><a href="index.asp?styp=xfrq&adtyp=<%=new_adtyp%>">续费日期<%=s_u_xfrq%></a></td>
<td width="70" align="center" background="images/bg.gif"><a href="index.asp?styp=xfje&adtyp=<%=new_adtyp%>">续费金额<%=s_u_xfje%></a></td>
<td width="75" align="center" background="images/bg.gif"><a href="index.asp?styp=zje&adtyp=<%=new_adtyp%>">总金额<%=s_u_zje%></a></td>
<td width="75" align="center" background="images/bg.gif"><a href="index.asp?styp=dqrq&adtyp=<%=new_adtyp%>">到期日期<%=s_u_dqrq%></a></td>
<td width="75" align="center" background="images/bg.gif">操作/修改</td>
<%If Session("x_user")<> "" Then%>
<td width="49" align="center" background="images/bg.gif">管理</td>
<%End If%>
</tr>
<%
Lours.MoveFirst
Lours.PageSize = Page_Size  '每页显示的条数  
If page < 1 Then page = 1 
If page > Lours.PageCount Then page = Lours.PageCount  
Lours.AbsolutePage = page
ipage=1
Do while Not Lours.Eof
%>
<tr uid="<%=Lours("id")%>">
<td height="24" align=center class="table4"><%=Lours("xy_user")%></td>
<td align=center class="table4"><span id="upass_<%=Lours("id")%>">***********</span><a href="javascript:void(0);" onClick="lookpass(<%=Lours("id")%>,'<%=Lours("xy_password")%>');">(显)</a></td>
<td align=center class="table4"><%=Lours("xy_dz")%></td>
<td align=center class="table4"><%=Lours("xy_ktrq")%></td>
<td align=center class="table4"><%
If IsDate(Lours("xy_xfrq")) Then
	Response.write Lours("xy_xfrq")
Else
	Response.write "<Font Color='Red'>新用户</Font>"
ENd if
%></td>
<td align=center class="table4"><%if trim(Lours("xy_xfje"))="" then response.Write("未交费") else response.Write(Lours("xy_xfje")) end if%></td>
<td align=center class="table4"><%if trim(Lours("xy_zje"))="" then response.Write("无数据") else response.Write(Lours("xy_zje"))  end if%></td>
<td align=center class="table4"><%If Lours("xy_dqrq") = #1900-1-1# then
response.write("&nbsp;")
Else
response.write(""&Lours("xy_dqrq")&"")
End If
If date() > Lours("xy_dqrq") then
response.write("<br><Font Color='Red'>欠费用户</Font>")
End If
%></td>
<td align=center class="table4" title=""><%=Lours("x_mz")%></br><%
If IsDate(Lours("xy_xgrq")) Then
	Response.write "<span title='操作时间 "&Lours("xy_xgrq")&" 修改'>"&FormatDateTime(Lours("xy_xgrq"),vbShortDate)&"</span>"
Else
	Response.write "<Font Color='Red'>未修改</Font>"
ENd if
%></td>
<%If Session("x_user")<> "" Then%>
<td align=center class="table4"><a href="Add_User.asp?doActs=Modify&LouHB_ID=<%=Lours("ID")%>"><img src="images/gai.gif" alt="修改" border="0"></a>  <a href="?doAct=delete&LouHB_ID=<%=Lours("ID")%>" onClick="return confirm('操作提示：您确实要删除吗？');"><img src="images/Del.gif" alt="删除" border="0"></a></td><%End IF%>
<%
Lours.MoveNext
Loop  
%>
</table>
<table border="0" align="center" cellpadding="0" cellspacing="0" width="768" class="table4">
<form onSubmit="return test();" method="get" action="">
<tr> 
<td width="220" height="30" background="images/bg.gif"> 
<%
If page = 1 Then 
Response.Write "第一页　上一页" 
End If
If page <> 1 Then 
Response.Write "<a href=index.asp?"&urlstr&"page=1&LouHB_XingMing="& L_XingMing &">第一页</a>" 
Response.Write "　<a href=index.asp?"&urlstr&"page=" & (page - 1) & "&LouHB_XingMing="& L_XingMing &">上一页</a>" 
End If
If page <> Lours.PageCount Then 
Response.Write "　<a href=index.asp?"&urlstr&"page=" & (page + 1) & "&LouHB_XingMing="& L_XingMing &">下一页</a>" 
Response.Write "　<a href=index.asp?"&urlstr&"page=" & Lours.PageCount & "&LouHB_XingMing="& L_XingMing &">最后一页</a>" 
End If 
If page = Lours.PageCount Then 
Response.Write "　下一页　最后一页" 
End If
%></td>
<td colspan="2" background="images/bg.gif">跳转到第 
<%
Response.Write hiddenstr
Response.Write "<input type=text size=5 maxlength=4 name=page class=button1><input type=hidden name=judge value=1>"  '显示输入页数框并将page,judge参数传递下去%>
<input type="hidden" name="LouHB_XingMing" value=<%=L_XingMing%>>
页</td>
</tr>
<tr> 
<td height="25" colspan="3">使用总数：<%=Lours.recordCount%>　总页数：<%=Lours.PageCount%>　当前页次：<%=page%></td>
</tr>
</form>
</table>
<%
End If
Lours.Close
Set Lours=Nothing
LouConn.Close
Set LouConn=Nothing
Call Copy()
%>
