<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%dbdns="../"%>
<!--#include file="chk.asp"-->
<!--#include file="../AppCode/fun/md5.asp"-->
<%
Call CheckAdminFlag(1)

Dim AdminId
Dim AdminName
Dim AdminPwd
Dim GroupId
Dim AdminLock

Select Case Trim(Request.Form("action"))
	Case "add"
		Call Add()
		Call SaveAdminLog("添加管理员：" & AdminName)
		Call CloseConn()
		Call ActionOk("Sys_admin.asp")
	Case "edit"
		Call Edit()
		Call SaveAdminLog("修改管理员(ID=" & AdminId & ")为：" & AdminName)
		Call CloseConn()
		Call ActionOk("Sys_admin.asp")
	Case "del"
		Call Del()
		Call SaveAdminLog("删除管理员(ID=" & AdminId & ")")
		Call CloseConn()
		Call ActionOk("Sys_admin.asp")
End Select

Private Sub Add()
	Call GetFormData()
	Sql = "select count(*) from Ok3w_Admin where AdminName='" & AdminName & "'"
	If Conn.Execute(Sql)(0)<>0 Then
		Call CloseConn()
		Session("ErrMsg") = "相同的管理员已经存在，请不要重复添加。"
		Call ActionOk("Sys_admin.asp")
	End If
	sql = "select * from Ok3w_Admin where 1=2"
	Rs.open sql,conn,1,3
	Rs.addnew
	Call UpdateRs(Rs,"add")
	Rs.update
	Rs.close
End Sub

Private Sub Edit()
	Call GetFormData()
	sql="select count(*) from Ok3w_Admin where AdminName='" & AdminName & "' and AdminId<>" & AdminId
	If Conn.Execute(sql)(0)<>0 Then
		Call CloseConn()
		Session("ErrMsg") = "相同的管理员已经存在，该修改无效。"
		Call ActionOk("Sys_admin.asp")
	End If
	sql = "select * from Ok3w_Admin where AdminId=" & AdminId
	Rs.Open Sql,Conn,1,3
	Call UpdateRs(Rs,"edit")
	Rs.Update
	Rs.Close
End Sub

Private Sub Del()
	Call GetFormData()
	sql = "delete from Ok3w_Admin where AdminId=" & AdminId
	Conn.Execute sql
End Sub

Private Sub GetFormData()
    AdminId = Trim(Request.Form("AdminId"))
    AdminName = Trim(Request.Form("AdminName"))
    AdminPwd = Trim(Request.Form("AdminPwd"))
    GroupId = "," & Replace(Request.Form("GroupId")," ","") & ","
    AdminLock = Trim(Request.Form("AdminLock"))
End Sub

Private Sub UpdateRs(ByRef Rs,action)
    Rs("AdminName") = AdminName
    If action="edit" And  AdminPwd="" Then
		Else
			Rs("AdminPwd") = MD5(AdminPwd)
			Response.Cookies("Ok3w")("AdminPwd") = MD5(AdminPwd)
	End If
    Rs("GroupId") = GroupId
    Rs("AdminLock") = AdminLock
End Sub
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台管理系统</title>
<link rel="stylesheet" type="text/css" href="images/Style.css">
</head>

<body>
<!--#include file="top.asp"-->
<br />
<table cellspacing="0" cellpadding="0" width="98%" align="center" border="0">
  <tbody>
    <tr>
      <td style="PADDING-LEFT: 2px; HEIGHT: 22px" 
    background="images/tab_top_bg.gif"><table cellspacing="0" cellpadding="0" width="477" border="0">
        <tbody>
          <tr>
            <td width="147"><table height="22" cellspacing="0" cellpadding="0" border="0">
              <tbody>
                <tr>
                  <td width="3"><img id="tabImgLeft__0" height="22" 
                  src="images/tab_active_left.gif" width="3" /></td>
                  <td class="mtitle" 
                background="images/tab_active_bg.gif">管理员管理</td>
                  <td width="3"><img id="tabImgRight__0" height="22" 
                  src="images/tab_active_right.gif" 
            width="3" /></td>
                </tr>
              </tbody>
            </table></td>
          </tr>
        </tbody>
      </table></td>
    </tr>
    <tr>
      <td bgcolor="#ffffff"><table cellspacing="0" cellpadding="0" width="100%" border="0">
        <tbody>
          <tr>
            <td width="1" background="images/tab_bg.gif"><img height="1" 
            src="images/tab_bg.gif" width="1" /></td>
            <td 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          valign="top"><div id="tabContent__0" style="DISPLAY: block; VISIBILITY: visible">
              <table cellspacing="1" cellpadding="1" width="100%" align="center" 
            bgcolor="#8ccebd" border="0">
                <tbody>
                  <tr>
                    <td 
                style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
                valign="top" bgcolor="#fffcf7"><table border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                      <form action="" method="post">
                        <tr>
                          <td align="center" bgcolor="#EBEBEB">用户名</td>
                          <td align="center" bgcolor="#EBEBEB">密码</td>
                          <td align="center" bgcolor="#EBEBEB">权限</td>
                          <td align="center" bgcolor="#EBEBEB">禁止登陆</td>
                          <td align="center" bgcolor="#EBEBEB">操作</td>
                        </tr>
                        <tr>
                          <td align="center" bgcolor="#FFFFFF"><input name="AdminName" type="text" id="AdminName" /></td>
                          <td align="center" bgcolor="#FFFFFF"><input name="AdminPwd" type="password" id="AdminPwd" /></td>
                          <td bgcolor="#FFFFFF"><input name="GroupId" type="checkbox" id="GroupId" value="1">
                            系统设置
                              <input name="GroupId" type="checkbox" id="GroupId" value="2">
                            网站说明<br>
                            <input name="GroupId" type="checkbox" id="GroupId" value="3">
                            新闻编辑
                            <input name="GroupId" type="checkbox" id="GroupId" value="4"> 
                            留言管理<br><input name="GroupId" type="checkbox" id="GroupId" value="5"> 
                            会员管理
							<input name="GroupId" type="checkbox" id="GroupId" value="6"> 
                            软件管理</td>
                          <td align="center" bgcolor="#FFFFFF"><input name="AdminLock" type="radio" value="0" checked="checked" />
                            否
                            <input name="AdminLock" type="radio" value="1" />
                            是</td>
                          <td align="center" bgcolor="#FFFFFF"><input type="button" name="Submit" value=" 添 加 " onClick="javascript:formsubmit(this.form,'add');" />
                              <input name="action" type="hidden" id="action" /></td>
                        </tr>
                      </form>
                    </table>
                      <br />
                      <table border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                        <tr>
                          <td align="center" bgcolor="#EBEBEB">用户名</td>
                          <td align="center" bgcolor="#EBEBEB">密码</td>
                          <td align="center" bgcolor="#EBEBEB">权限</td>
                          <td align="center" bgcolor="#EBEBEB">锁定</td>
                          <td align="center" bgcolor="#EBEBEB">操作</td>
                        </tr>
                        <%
sql = "select * from Ok3w_Admin "
Rs.Open sql,Conn,1,1
Do While Not Rs.Eof
    AdminId = Rs("AdminId")
    AdminName = Rs("AdminName")
    AdminPwd = Rs("AdminPwd")
    GroupId = Rs("GroupId")
    AdminLock = Rs("AdminLock")
%>
                        <form id="form1" name="form1" method="post" action="">
                          <tr>
                            <td align="center" bgcolor="#FFFFFF"><input name="AdminName" type="text" id="AdminName" value="<%=AdminName%>" readonly="readonly" /></td>
                            <td align="center" bgcolor="#FFFFFF"><input name="AdminPwd" type="password" id="AdminPwd" /></td>
                            <td bgcolor="#FFFFFF"><input name="GroupId" type="checkbox" id="GroupId" value="1" <%If InStr(GroupId,",1,")>=1 Then%>checked<%End If%>>
系统设置
  <input name="GroupId" type="checkbox" id="GroupId" value="2" <%If InStr(GroupId,",2,")>=1 Then%>checked<%End If%>>
网站说明<br>
<input name="GroupId" type="checkbox" id="GroupId" value="3" <%If InStr(GroupId,",3,")>=1 Then%>checked<%End If%>>
新闻编辑
<input name="GroupId" type="checkbox" id="GroupId" value="4" <%If InStr(GroupId,",4,")>=1 Then%>checked<%End If%>>
留言管理<br><input name="GroupId" type="checkbox" id="GroupId" value="5" <%If InStr(GroupId,",5,")>=1 Then%>checked<%End If%>> 
                            会员管理
							<input name="GroupId" type="checkbox" id="GroupId" value="6" <%If InStr(GroupId,",6,")>=1 Then%>checked<%End If%>> 
                            软件管理</td>
                            <td align="center" bgcolor="#FFFFFF"><input name="AdminLock" type="radio" value="0" <%If Not AdminLock Then%>checked="checked"<%End If%> />
                              否
                              <input name="AdminLock" type="radio" value="1" <%If AdminLock Then%>checked="checked"<%End If%> />
                              是</td>
                            <td align="center" bgcolor="#FFFFFF"><input type="button" name="Submit4" value="修改" onClick="javascript:formsubmit(this.form,'edit');"  />
                                <input type="button" name="Submit5" value="删除" onClick="javascript:if(!confirm('真的要删除吗？')) return false;formsubmit(this.form,'del');" <%If Request.Cookies("Ok3w")("AdminName")= AdminName Then%>disabled="disabled"<%End If%> />
                                <input name="AdminId" type="hidden" id="AdminId" value="<%=AdminId%>" />
                                <input name="action" type="hidden" id="action" /></td>
                          </tr>
                        </form>
                        <%
	Rs.MoveNext
Loop
Rs.Close
%>
                      </table>
                        <script language="JavaScript" type="text/javascript">
function formsubmit(frm,action)
{
	if(frm.AdminName.value.trim()=="")
	{
		ShowErrMsg("管理员名称不能为空，请输入");
		frm.AdminName.focus();
		return false;
	}
	if(frm.AdminPwd.value.trim()=="" && action =="add")
	{
		ShowErrMsg("管理员密码不能为空，请输入");
		frm.AdminPwd.focus();
		return false;
	}
	
	frm.action.value = action;
	frm.submit();
}
                      </script>
                      </td>
                  </tr>
                </tbody>
              </table>
            </div></td>
            <td width="1" background="images/tab_bg.gif"><img height="1" 
            src="images/tab_bg.gif" width="1" /></td>
          </tr>
        </tbody>
      </table></td>
    </tr>
    <tr>
      <td background="images/tab_bg.gif" bgcolor="#ffffff"><img height="1" 
      src="images/tab_bg.gif" width="1" /></td>
    </tr>
  </tbody>
</table>
</body>
</html>
<%
Call CloseConn()
Set Rs = Nothing
Set Admin = Nothing
%>

