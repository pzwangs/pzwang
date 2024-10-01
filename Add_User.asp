<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!-- #include file="Config.asp"-->
<!-- #include file="Check_Session.asp"-->
<!-- #include file="md5.asp"-->
<!-- #include file="about.asp" -->
<!--#include file="head.asp"-->
<!-- #include file="top.asp"-->
<%
Dim doAct
doAct=Trim(Request("doAct"))
if doAct="deal" Then
	'ID=Request("LouHB_ID") 					'仅用于修改
	'xy_admin=session("x_user") 					'管理员登陆
	xy_user=Request("LouHB_user")				'添加的新用户
	xy_password=Request("LouHB_password2") 		'密码
	xy_dz=Request("LouHB_dz") 					'地址
	xy_ktrq=Request("LouHB_ktrq") 				'开通日期
	xy_xfrq=Request("LouHB_xfrq")				'续费日期
	xy_xfje=Request("LouHB_xfje") 				'续费金额
	xy_dqrq=Request("LouHB_dqrq") 				'到期日期
	xy_bz=Request("LouHB_bz") 					'备注
	opex = Request("op") 						'执行操作指令
	cid=Trim(Request("cid"))
	if xy_user="" then Call ld_error("账号为必填项目",0)
	
	if xy_dz="" then Call ld_error("地址为必填项目",0)
	
	if Not IsDate(xy_dqrq) Then Call ld_Error("到期日期的格式不正确！",0)
	if Not IsNumeric(xy_xfje) Then Call ld_Error("续费金额必需为数字！",0)
	lactsql="select ID from xy_data where xy_user='"&xy_user&"'"
	sql="Select * From xy_data"
	altermsg="添加成功"
	if opex="Modifys" Then
		if Not IsNumeric(cid) Then cid=0
		lactsql=lactsql&" And ID<>"&cid&""
		sql=sql&" where ID="&cid&""
		altermsg="修改成功"
	Else
		if xy_password="" then Call ld_error("密码为必填项目",0)
		if Not IsDate(xy_ktrq) Then Call ld_Error("开通日期的格式不正确！",0)
	End if
	'判断会员是否已存在
	set rs=LouConn.Execute(lactsql)
	if Not(Rs.Eof Or Rs.BOf) Then
		Call ld_Error("此会员账号已经存在，无法添加或修改！",0)
	End if
	Rs.Close:Set Rs=Nothing
	Set uRs=server.createobject("adodb.recordset")
	uRs.Open sql,LouConn,3,3
	if opex="Add_User" Then
		uRs.addnew
		uRs("xy_ktrq")=xy_ktrq				'新用户开通日期
		uRs("xy_xfje")=0					'续费金额
	Else
		if uRs.Eof And uRs.Bof Then
			Call ld_Error("您要修改的会员不存在，无法修改！",0)
		End if
		uRs("xy_xfje")=xy_xfje					'续费金额
		uRs("xy_zje")=cint(xy_xfje)+cint(xy_zje_tem)			'总金额
	End if
		uRs("xy_admin")=x_user			'管理员，或者操作员记录
		uRs("xy_user")=xy_user				'新添加的用户
		if xy_password<>"" then
		'uRs("xy_password")=md5(xy_password,16)	'添加新用户的密码 加密
		uRs("xy_password")=xy_password	'添加新用户的密码 不加密
		end if
		uRs("xy_dz")=xy_dz					'地址
		uRs("xy_xfrq")=xy_xfrq				'续费日期
		uRs("xy_dqrq")=xy_dqrq				'到期日期
		uRs("xy_zje")=cint(xy_xfje)+cint(trim(request("xy_zje_tem")))  '总金额
		uRs("xy_xgrq")=Now()				'修改日期
		uRs("xy_bz")=xy_bz					'备注
	uRs.update
	uRs.close
	set uRs=nothing
	Call ld_Error(altermsg,"./index.asp")
End if
%>
<body>
<%
if request("doActs")="Add_New" then
%>
<table width="41%"  border="0" cellspacing="0" cellpadding="0" align="center" class="table4">
  <form name="Select_Date" method="post" action="add_user.asp">
    <tr align="center">
      <td height="30" colspan="2" background="images/bg.gif"> 新用户添加
        <input name="op" type="hidden" id="op" value="Add_User">
        <input name="doAct" type="hidden" id="doAct" value="deal">
      </td>
    </tr>
    <tr>
      <td width="30%" align="center"><div align="right">操作员： </div></td>
      <td width="70%">-= <%=web_x_mz%>=-</td>
    </tr>
    <tr>
      <td align="center"><div align="right"></div></td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td align="center"><div align="right">账号：</div></td>
      <td><input name="LouHB_user" type="text" class="button1" size="10" ></td>
    </tr>
    <tr>
      <td align="center"><div align="right"></div></td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td align="center"><div align="right">密码：</div></td>
      <td><input name="LouHB_password2" type="password" class="button1" id="LouHB_password2" value="" size="15"></td>
    </tr>
    <tr>
      <td align="center"><div align="right"></div></td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td align="center"><div align="right">址地：</div></td>
      <td><input name="LouHB_dz" type="text" class="button1" value="" size="15"></td>
    </tr>
    <tr>
      <td align="center"><div align="right"></div></td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td align="center"><div align="right">开通日期：</div></td>
      <td><table width="290" border="0">
          <tr>
            <td width="75"><input name="LouHB_ktrq" type="text" class="button1" id="LouHB_ktrq" value="<%=Date()%>" size="10"></td>
            <td width="205"><a href="#SelectDate" onClick="javascript:window.open('Select_Date.asp?form=Select_Date&field=LouHB_ktrq','','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=190,height=170');"><img src="images/RiLi.gif" width="30" height="30" border="0" alt="选择日期"></a></td>
          </tr>
        </table></td>
    </tr>
    
        <tr>
      <td align="center"><div align="right"></div></td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td align="center"><div align="right">缴纳金额：</div></td>
      <td><input name="LouHB_xfje" type="text" class="button1" value="" size="10" >
        元 </td>
    </tr>
    <tr>
      <td align="center"><div align="right"></div></td>
      <td>&nbsp;</td>
    </tr>

    <tr>
      <td align="center"><div align="right">到期日期：</div></td>
      <td><table width="253" border="0">
          <tr>
            <td width="75"><input name="LouHB_dqrq" type="text" class="button1" value="" size="10"></td>
            <td width="168"><a href="#SelectDate" onClick="javascript:window.open('Select_Date.asp?form=Select_Date&field=LouHB_dqrq','','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=190,height=170');"><img src="images/RiLi.gif" width="30" height="30" border="0" alt="选择日期"></a></td>
          </tr>
        </table></td>
    </tr>
    <tr>
      <td align="center">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td align="center"><div align="right">备注：</div></td>
      <td><textarea name="LouHB_bz" cols="40" class="button1"></textarea>
      </td>
    </tr>
    <tr align="center">
      <td colspan="2"><%
response.write"<input name=""LouHB_Reg"" type=""submit"" class=""button1"" value=""添加"">"
%></td>
    </tr>
    <tr>
      <td colspan="2">&nbsp;</td>
    </tr>
  </form>
</table>
<% end if
if request("doActs")="Modify" then

Set Rs=server.createobject("adodb.recordset")
		sql="Select * From xy_data where id="&trim(Request("LouHB_ID"))
		Rs.Open sql,LouConn,1,1
%>
<table width="41%"  border="0" cellspacing="0" cellpadding="0" align="center" class="table4">
  <form name="Select_Date" method="post" action="add_user.asp">
    <tr align="center">
      <td height="30" colspan="4" background="images/bg.gif">用户资料修改
        <input name="op" type="hidden" id="op" value="Modifys">
        <input name="doAct" type="hidden" id="doAct" value="deal">
        <input name="cid" type="hidden" id="cid" value="<%=trim(Request("LouHB_ID"))%>">
		<input name="xy_zje_tem" type="hidden" id="xy_zje_tem" value="<%=rs("xy_zje")%>"></td>
    </tr>
    <tr>
      <td width="17%" align="center"><div align="right">操作员： </div></td>
      <td colspan="3">-= <%=web_x_mz%>=-</td>
    </tr>
    <tr>
      <td align="center"><div align="right"></div></td>
      <td colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td align="right">账号：</td>
      <td width="18%"><input name="LouHB_user" type="text" class="button1" value="<%=rs("xy_user")%>" size="7"></td>
      <td width="24%"><font color="#FF0000">总交金额：<%=rs("xy_zje")%>元 </font></td>
      <td width="41%">开通日期：<%=formatdatetime(rs("xy_ktrq"),2)%></td>
    </tr>
    <tr>
      <td align="right"><div align="right"></div></td>
      <td colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td align="right">密码：</td>
      <td colspan="3"><input name="LouHB_password2" type="password" class="button1" id="LouHB_password2" size="12">
        <span style="color: #FF0000">密码不改请留空，改了请记住。</span></td>
    </tr>
    <tr>
      <td align="right"><div align="right"></div></td>
      <td colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td align="right">址地：</td>
      <td colspan="3"><input name="LouHB_dz" type="text" class="button1" value="<%=rs("xy_dz")%>" size="15"></td>
    </tr>
    <tr>
      <td align="right">&nbsp;</td>
      <td colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td align="right">续费日期：</td>
      <td colspan="3"><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="35%"><input name="LouHB_xfrq" type="text" class="button1" size="15" value="<%=Date()%>"></td>
    <td width="65%">
      <a href="#SelectDate" onClick="javascript:window.open('Select_Date.asp?form=Select_Date&field=LouHB_xfrq','','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=190,height=170');"><img src="images/RiLi.gif" width="30" height="30" border="0" alt="选择日期"></a></td>
  </tr>
</table>
</td>
    </tr>
    <tr>
      <td align="right">&nbsp;</td>
      <td colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td align="right">续费金额：</td>
      <td colspan="3"><input name="LouHB_xfje" type="text" class="button1" value="<%=rs("xy_xfje")%>" size="10">
        元 </td>
    </tr>
    <tr>
      <td align="right">&nbsp;</td>
      <td colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td align="right">到期日期：</td>
      <td colspan="3"><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="36%"><input name="LouHB_dqrq" type="text" class="button1" value="<%=DateAdd("m",1,formatdatetime(rs("xy_dqrq"),2))%>" size="15"<%If L_Group>2 then response.write(" readonly") End If%>></td>
    <td width="64%">
        <a href="#SelectDate" onClick="javascript:window.open('Select_Date.asp?form=Select_Date&field=LouHB_dqrq','','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=190,height=170');"><img src="images/RiLi.gif" width="30" height="30" border="0" alt="选择日期"></a></td>
  </tr>
</table>

        </td>
    </tr>
    <tr>
      <td align="right">&nbsp;</td>
      <td colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td align="right">备注：</td>
      <td colspan="3"><textarea name="LouHB_bz" cols="40" class="button1" ><%=rs("xy_bz")%></textarea></td>
    </tr>
    <tr>
      <td colspan="4">&nbsp;</td>
    </tr>
    <tr align="center">
      <td colspan="4"><%
response.write"<input name=""LouHB_Reg"" type=""submit"" class=""button1"" value=""修改"">"
%></td>
    </tr>
  </form>
</table>
<%end if
Call Copy()
%>
</body>
</html>
