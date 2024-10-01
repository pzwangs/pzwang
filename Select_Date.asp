<%
'WEB Calendar
'By Chaiwei 2002-7-31
'--------------------------------

'月份名称定义
Dim Month_Name(12)
Month_Name(1) = "1月"
Month_Name(2) = "2月"
Month_Name(3) = "3月"
Month_Name(4) = "4月"
Month_Name(5) = "5月"
Month_Name(6) = "6月"
Month_Name(7) = "7月"
Month_Name(8) = "8月"
Month_Name(9) = "9月"
Month_Name(10) = "10月"
Month_Name(11) = "11月"
Month_Name(12) = "12月"

'年份处理,默认值为服务器当前年份
if request.querystring("year")<>"" then
Year_var=cint(request.querystring("year"))
else
Year_var=year(date())
end if

'上一年、下一年赋值
Previous_year=Year_var-1
Next_year=Year_var+1


'月份处理,默认值为服务器当前月份
if request.querystring("Month")<>"" then
Month_var=cint(request.querystring("Month"))
else
Month_var=month(date())
end if

'上一月、下一月赋值
if Month_var<=1 then
Next_month=Month_var+1
Previous_month=1
else
if Month_var>=12 then
Next_month=12
Previous_month=Month_var-1
else
Next_month=Month_var+1
Previous_month=Month_var-1
end if
end if

'当前天数定位计算
First_day=DateSerial(Year_var,Month_var,1)
Current_day=First_day-weekday(First_day)+1 '加1周日为第一天，加2周一为第一天

%>
<html>
<head>
<title>日期选择器</title>
<Script Language="JavaScript">

//前端日期选择函数

function pick(v) {
window.opener.document.<%=request.querystring("form")%>.<%=request.querystring("field")%>.value=v;window.close();return false;
}
</Script>
<style>
<!--
.page{text-decoration: none; color: #CAE3FF; font-size:9pt; font-family:Webdings}
.dayTable{ border: 1px dotted #E6E6E6; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1}
.day { font-family: Arial; font-size: 9pt; text-decoration: underline; color: #000000 }
:hover.day   { font-family: Arial; font-size: 9pt; text-decoration: none; color: #FF0000 }
.title { font-family: Arial; font-size: 9pt; color: #FFFFFF; font-weight: bold }
:hover.page  { text-decoration: underline; color: #FFFFFF; font-family:Webdings; font-size:9pt }
-->
</style>
</head>
<body topmargin="0" leftmargin="0" onLoad="window.focus();">
<div align="center">
<center>
<table border="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber1" cellpadding="0">
<tr>
<td width="100%" bgcolor="#003063">
<!--日历表头显示-->
<div align="center">
<center>
<table border="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber3" cellpadding="2">
<tr>
<td width="20%" align="center">
<a href="Select_Date.asp?year=<%=Previous_year%>&month=<%=Month_var%>&form=<%=request.querystring("form")%>&field=<%=request.querystring("field")%>" title="上一年" class="page">7</a>
<a href="Select_Date.asp?year=<%=Year_var%>&month=<%=Previous_month%>&form=<%=request.querystring("form")%>&field=<%=request.querystring("field")%>" title="上一月" class="page">3</a></td>
<td width="60%" align="center" class="title"><%response.write Month_Name(Month_var) & "　" & Year_var%>年</td>
<td width="20%" align="center"><a href="Select_Date.asp?year=<%=Year_var%>&month=<%=Next_month%>&form=<%=request.querystring("form")%>&field=<%=request.querystring("field")%>" title="下一月" class="page">4</a>
<a href="Select_Date.asp?year=<%=Next_year%>&month=<%=Month_var%>&form=<%=request.querystring("form")%>&field=<%=request.querystring("field")%>" title="下一年" class="page">8</a></td>
</tr>
</table>
</center>
</div>
</td>
</tr>
<tr>
<td width="100%">
<div align="center">
<center>
<table border="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber2" cellpadding="3">
<tr>
<td align="center" bgcolor="#31659C" class="title">日</td>
<td align="center" bgcolor="#31659C" class="title">一</td>
<td align="center" bgcolor="#31659C" class="title">二</td>
<td align="center" bgcolor="#31659C" class="title">三</td>
<td align="center" bgcolor="#31659C" class="title">四</td>
<td align="center" bgcolor="#31659C" class="title">五</td>
<td align="center" bgcolor="#31659C" class="title">六</td>
</tr>
<%
'日历内容 5行*7例 显示
'外层循环显示控制行

for i=0 to 5
%>
<tr>
<%
'内层循环显示控制列

for j=0 to 6
response.write "<td align='center' class='dayTable'"
'天数显示，“今天”显示
if Current_day = date then
response.write " bgcolor='#00FFFF'>"
%>
<a href="javascript:pick('<%=Current_day%>');" title="今天" class="day"><strong><%=day(Current_day)%></strong></a>
<%else '天数显示，非本月天数显示
if Month(Current_day) <> Month_var then
response.write "bgcolor='#F0F0F0'>"
%>
<a href="javascript:pick('<%=Current_day%>');" title="<%=Current_day%>" class="day"><font color="#CCCCCC"><%=day(Current_day)%></font></a>
<%else  '天数显示，本月天数显示
response.write ">"
%>
<a href="javascript:pick('<%=Current_day%>');" title="<%=Current_day%>" class="day"><%=day(Current_day)%></a>
<%end if
end if

'天数累加推算

Current_day = Current_day + 1
response.write "</td>"
next
%>
</tr>
<%
next
%>
</table>
</center>
</div>
</td>
</tr>
</table>
</center>
</div>
</body>
</html>
