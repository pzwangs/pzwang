<%
'WEB Calendar
'By Chaiwei 2002-7-31
'--------------------------------

'�·����ƶ���
Dim Month_Name(12)
Month_Name(1) = "1��"
Month_Name(2) = "2��"
Month_Name(3) = "3��"
Month_Name(4) = "4��"
Month_Name(5) = "5��"
Month_Name(6) = "6��"
Month_Name(7) = "7��"
Month_Name(8) = "8��"
Month_Name(9) = "9��"
Month_Name(10) = "10��"
Month_Name(11) = "11��"
Month_Name(12) = "12��"

'��ݴ���,Ĭ��ֵΪ��������ǰ���
if request.querystring("year")<>"" then
Year_var=cint(request.querystring("year"))
else
Year_var=year(date())
end if

'��һ�ꡢ��һ�긳ֵ
Previous_year=Year_var-1
Next_year=Year_var+1


'�·ݴ���,Ĭ��ֵΪ��������ǰ�·�
if request.querystring("Month")<>"" then
Month_var=cint(request.querystring("Month"))
else
Month_var=month(date())
end if

'��һ�¡���һ�¸�ֵ
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

'��ǰ������λ����
First_day=DateSerial(Year_var,Month_var,1)
Current_day=First_day-weekday(First_day)+1 '��1����Ϊ��һ�죬��2��һΪ��һ��

%>
<html>
<head>
<title>����ѡ����</title>
<Script Language="JavaScript">

//ǰ������ѡ����

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
<!--������ͷ��ʾ-->
<div align="center">
<center>
<table border="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber3" cellpadding="2">
<tr>
<td width="20%" align="center">
<a href="Select_Date.asp?year=<%=Previous_year%>&month=<%=Month_var%>&form=<%=request.querystring("form")%>&field=<%=request.querystring("field")%>" title="��һ��" class="page">7</a>
<a href="Select_Date.asp?year=<%=Year_var%>&month=<%=Previous_month%>&form=<%=request.querystring("form")%>&field=<%=request.querystring("field")%>" title="��һ��" class="page">3</a></td>
<td width="60%" align="center" class="title"><%response.write Month_Name(Month_var) & "��" & Year_var%>��</td>
<td width="20%" align="center"><a href="Select_Date.asp?year=<%=Year_var%>&month=<%=Next_month%>&form=<%=request.querystring("form")%>&field=<%=request.querystring("field")%>" title="��һ��" class="page">4</a>
<a href="Select_Date.asp?year=<%=Next_year%>&month=<%=Month_var%>&form=<%=request.querystring("form")%>&field=<%=request.querystring("field")%>" title="��һ��" class="page">8</a></td>
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
<td align="center" bgcolor="#31659C" class="title">��</td>
<td align="center" bgcolor="#31659C" class="title">һ</td>
<td align="center" bgcolor="#31659C" class="title">��</td>
<td align="center" bgcolor="#31659C" class="title">��</td>
<td align="center" bgcolor="#31659C" class="title">��</td>
<td align="center" bgcolor="#31659C" class="title">��</td>
<td align="center" bgcolor="#31659C" class="title">��</td>
</tr>
<%
'�������� 5��*7�� ��ʾ
'���ѭ����ʾ������

for i=0 to 5
%>
<tr>
<%
'�ڲ�ѭ����ʾ������

for j=0 to 6
response.write "<td align='center' class='dayTable'"
'������ʾ�������족��ʾ
if Current_day = date then
response.write " bgcolor='#00FFFF'>"
%>
<a href="javascript:pick('<%=Current_day%>');" title="����" class="day"><strong><%=day(Current_day)%></strong></a>
<%else '������ʾ���Ǳ���������ʾ
if Month(Current_day) <> Month_var then
response.write "bgcolor='#F0F0F0'>"
%>
<a href="javascript:pick('<%=Current_day%>');" title="<%=Current_day%>" class="day"><font color="#CCCCCC"><%=day(Current_day)%></font></a>
<%else  '������ʾ������������ʾ
response.write ">"
%>
<a href="javascript:pick('<%=Current_day%>');" title="<%=Current_day%>" class="day"><%=day(Current_day)%></a>
<%end if
end if

'�����ۼ�����

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
