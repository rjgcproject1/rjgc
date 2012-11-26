<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%dbdns="../"%>
<!--#include file="chk.asp"-->
<!--#include file="../AppCode/Class/Ok3w_SiteConfig.asp"-->
<!--#include file="../AppCode/fun/function.asp"-->
<%
Call CheckAdminFlag(1)

Select Case Trim(Request.Form("action"))
	Case "edit"
		Set SiteConfig = New Ok3w_SiteConfig
		Call SiteConfig.Edit()
		Call SaveAdminLog("编辑系统设置")
		Call CloseConn()
		Call ActionOk("Sys_Config.asp")
End Select

Function IsObjInstalled(ObjName)
	On Error ReSume Next
	Set Obj = Server.CreateObject(ObjName)
	If Err.Number<>0 Then
		Err.Clear
		IsObjInstalled = False
		Else
			IsObjInstalled = True
	End If
	On Error Goto 0
End Function
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
                  <td 
                background="images/tab_active_bg.gif" class="tab"><strong class="mtitle">后台管理系统</strong></td>
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
                valign="top" bgcolor="#fffcf7"><table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                      <form id="form1" name="form1" method="post" action="">
                        <tr>
                          <td colspan="2" bgcolor="#EBEBEB"><strong>网站运行设置</strong></td>
                        </tr>
                        <tr>
                          <td width="18%" align="right" bgcolor="#FFFFFF">网站关闭</td>
                          <td width="82%" bgcolor="#FFFFFF"><input name="SiteIsClose" type="checkbox" id="SiteIsClose" value="1" <%If Application("Ok3w_SiteIsClose")="1" Then%>checked="checked"<%End If%> /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">网站关闭原因</td>
                          <td bgcolor="#FFFFFF"><input name="SiteCloseNote" type="text" value="<%=Application("Ok3w_SiteCloseNote")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td colspan="2" align="left" bgcolor="#EBEBEB"><strong>网站基本设置</strong></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">网站名称</td>
                          <td bgcolor="#FFFFFF"><input name="SiteName" type="text" value="<%=Application("Ok3w_SiteName")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">网站标题</td>
                          <td bgcolor="#FFFFFF"><input name="SiteTitle" type="text" value="<%=Application("Ok3w_SiteTitle")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">网站关键字</td>
                          <td bgcolor="#FFFFFF"><input name="SiteKeyWords" type="text" value="<%=Application("Ok3w_SiteKeyWords")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">网站描述</td>
                          <td bgcolor="#FFFFFF"><input name="SiteDescription" type="text" value="<%=Application("Ok3w_SiteDescription")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">网站网址</td>
                          <td bgcolor="#FFFFFF"><input name="SiteUrl" type="text" value="<%=Application("Ok3w_SiteUrl")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">&nbsp;TCP/IP备案号</td>
                          <td bgcolor="#FFFFFF"><input name="SiteTCPIP" type="text" id="SiteTCPIP" value="<%=Application("Ok3w_SiteTCPIP")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">网友留言审核</td>
                          <td bgcolor="#FFFFFF"><input name="SiteIsGuest" type="checkbox" id="SiteIsGuest" value="1" <%if Application("Ok3w_SiteIsGuest")="1" then%>checked="checked"<%End If%> />
                            需要先审核后显示</td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">留言簿公告</td>
                          <td bgcolor="#FFFFFF"><textarea name="SiteGuestNews" cols="48" rows="3" id="SiteGuestNews"><%=Application("Ok3w_SiteGuestNews")%></textarea></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">流量统计代码</td>
                          <td bgcolor="#FFFFFF"><textarea name="SiteTongji" cols="48" rows="3" id="SiteTongji"><%=Application("Ok3w_SiteTongji")%></textarea></td>
                        </tr>
                        
                        
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">全局关键词</td>
                          <td bgcolor="#FFFFFF"><textarea name="SitePublicKeyWords" cols="80" rows="5" id="SitePublicKeyWords"><%=Application("Ok3w_SitePublicKeyWords")%></textarea>
                            <br>
                            一行一个，关键词和链接用“|”分隔</td>
                        </tr>
                        <tr style="display:none">
                          <td align="right" bgcolor="#FFFFFF">网站所在目录</td>
                          <td bgcolor="#FFFFFF"><input name="SitePath" type="text" id="SitePath" value="<%=Application("Ok3w_SitePath")%>" size="20" /></td>
                        </tr>
                        <tr style="display:none">
                          <td align="right" bgcolor="#FFFFFF">数据库所在目录</td>
                          <td bgcolor="#FFFFFF"><input name="SiteDbPath" type="text" id="SiteDbPath" value="<%=Application("Ok3w_SiteDbPath")%>" size="20" /></td>
                        </tr>
                        <tr style=" display:none">
                          <td align="right" bgcolor="#FFFFFF">后台地址</td>
                          <td bgcolor="#FFFFFF"><input name="SiteAdminPath" type="text" value="<%=Application("Ok3w_SiteAdminPath")%>" size="20" /></td>
                        </tr>
                        <tr>
                          <td colspan="2" bgcolor="#EBEBEB"><strong>联系方式设置</strong></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">公司名称</td>
                          <td bgcolor="#FFFFFF"><input name="SiteCoName" type="text" id="SiteCoName" value="<%=Application("Ok3w_SiteCoName")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">联系人</td>
                          <td bgcolor="#FFFFFF"><input name="SiteLinkMan" type="text" id="SiteLinkMan" value="<%=Application("Ok3w_SiteLinkMan")%>" size="20" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">电话</td>
                          <td bgcolor="#FFFFFF"><input name="SiteTel" type="text" value="<%=Application("Ok3w_SiteTel")%>" size="20" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">传真</td>
                          <td bgcolor="#FFFFFF"><input name="SiteFax" type="text" value="<%=Application("Ok3w_SiteFax")%>" size="20" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">地址</td>
                          <td bgcolor="#FFFFFF"><input name="SiteAddress" type="text" value="<%=Application("Ok3w_SiteAddress")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">邮编</td>
                          <td bgcolor="#FFFFFF"><input name="SiteZip" type="text" value="<%=Application("Ok3w_SiteZip")%>" size="10" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">邮箱</td>
                          <td bgcolor="#FFFFFF"><input name="SiteEmail" type="text" value="<%=Application("Ok3w_SiteEmail")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">QQ</td>
                          <td bgcolor="#FFFFFF"><input name="SiteQQ" type="text" value="<%=Application("Ok3w_SiteQQ")%>" size="20" /></td>
                        </tr>
                        <tr>
                          <td colspan="2" align="left" bgcolor="#EBEBEB"><strong>注册会员设置</strong></td>
                          </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">不允许注册的会员名称</td>
                          <td bgcolor="#FFFFFF"><input name="SiteUserRegKillName" type="text" id="SiteUserRegKillName" value="<%=Application("Ok3w_SiteUserRegKillName")%>" size="50" maxlength="50" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">会员注册条款</td>
                          <td bgcolor="#FFFFFF"><textarea name="SiteUserRegTK" cols="48" rows="5" id="SiteUserRegTK"><%=Application("Ok3w_SiteUserRegTK")%></textarea></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">会员注册成功赠送</td>
                          <td bgcolor="#FFFFFF"><input name="SiteUserRegJifen" type="text" id="SiteUserRegJifen" value="<%=Application("Ok3w_SiteUserRegJifen")%>" size="4">
                            积分</td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">每天第一次登陆赠送</td>
                          <td bgcolor="#FFFFFF"><input name="SiteUserLoginJifen" type="text" id="SiteUserLoginJifen" value="<%=Application("Ok3w_SiteUserLoginJifen")%>" size="4">
积分</td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">投稿通过审核，赠送</td>
                          <td bgcolor="#FFFFFF"><input name="SiteUserTGJifen" type="text" id="SiteUserTGJifen" value="<%=Application("Ok3w_SiteUserTGJifen")%>" size="4">
积分</td>
                        </tr>
                        
                        <tr>
                          <td colspan="2" align="left" bgcolor="#EBEBEB"><strong>上传图片水印设置</strong></td>
                        </tr>
                        
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">水印组件</td>
                          <td bgcolor="#FFFFFF"><input name="Sitesy_DLL" type="radio" value="0" <%If Application("Ok3w_Sitesy_DLL")=0 Then%>checked<%End If%>>
                            aspJpeg</td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">水印类型</td>
                          <td bgcolor="#FFFFFF"><input type="radio" name="Sitesy_Type" value="0" <%If Application("Ok3w_Sitesy_Type")=0 Then%>checked<%End If%>>
关闭水印功能
<%If IsObjInstalled("Persits.Jpeg") Then%>
  <input type="radio" name="Sitesy_Type" value="1" <%If Application("Ok3w_Sitesy_Type")=1 Then%>checked<%End If%>>
文字水印
<input type="radio" name="Sitesy_Type" value="2" <%If Application("Ok3w_Sitesy_Type")=2 Then%>checked<%End If%>>
图片水印
<%Else%>
<span class="red">注意：你还没有安装水印组件，无法启用水印功能。</span>
<%End If%></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">水印位置</td>
                          <td bgcolor="#FFFFFF"><input type="radio" name="Sitesy_Location" value="0" <%If Application("Ok3w_Sitesy_Location")=0 Then%>checked<%End If%>>
                            左上
                              <input type="radio" name="Sitesy_Location" value="1" <%If Application("Ok3w_Sitesy_Location")=1 Then%>checked<%End If%>>
                              右上
                              <input type="radio" name="Sitesy_Location" value="2" <%If Application("Ok3w_Sitesy_Location")=2 Then%>checked<%End If%>>
                              左下
                              <input type="radio" name="Sitesy_Location" value="3" <%If Application("Ok3w_Sitesy_Location")=3 Then%>checked<%End If%>>
                              右下
                              <input type="radio" name="Sitesy_Location" value="4" <%If Application("Ok3w_Sitesy_Location")=4 Then%>checked<%End If%>>
                              居中</td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">水印文字</td>
                          <td bgcolor="#FFFFFF"><input name="Sitesy_Text" type="text" id="Sitesy_Text" value="<%=Application("Ok3w_Sitesy_Text")%>" size="20">
                            字体
                              <input name="Sitesy_Family" type="text" id="Sitesy_Family" value="<%=Application("Ok3w_Sitesy_Family")%>" size="6">
                              字号
                              <input name="Sitesy_Size" type="text" id="Sitesy_Size" value="<%=Application("Ok3w_Sitesy_Size")%>" size="4">
                              px
                              颜色
                              #
                              <input name="Sitesy_Color" type="text" id="Sitesy_Color" value="<%=Application("Ok3w_Sitesy_Color")%>" size="8"></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">水印图片</td>
                          <td bgcolor="#FFFFFF"><input name="Sitesy_Logo" type="text" id="Sitesy_Logo" value="<%=Application("Ok3w_Sitesy_Logo")%>" size="50"></td>
                        </tr>
                        
                        <tr style="display:none">
                          <td colspan="2" align="left" bgcolor="#EBEBEB">邮件服务器设置</td>
                        </tr>
                        <tr style="display:none">
                          <td align="right" bgcolor="#FFFFFF">发送组件选择 </td>
                          <td bgcolor="#FFFFFF"><select name="SiteMailSmtpType" id="SiteMailSmtpType">
                              <option value="">关闭邮件发送功能</option>
                              <option value="Jmail" <%If Application("Ok3w_SiteMailSmtpType") = "Jmail" Then%>selected="selected"<%End If%>>Jmail</option>
                            </select>                          </td>
                        </tr>
                        <tr style="display:none">
                          <td align="right" bgcolor="#FFFFFF">发件人邮箱</td>
                          <td bgcolor="#FFFFFF"><input name="SiteMailSmtpFromEmail" type="text" value="<%=Application("Ok3w_SiteMailSmtpFromEmail")%>" /></td>
                        </tr>
                        <tr style="display:none">
                          <td align="right" bgcolor="#FFFFFF">发件人姓名</td>
                          <td bgcolor="#FFFFFF"><input name="siteMailSmtpFromUser" type="text" value="<%=Application("Ok3w_siteMailSmtpFromUser")%>" /></td>
                        </tr>
                        <tr style="display:none">
                          <td colspan="2" align="left" bgcolor="#EBEBEB">支付宝设置</td>
                        </tr>
                        <tr style="display:none">
                          <td align="right" bgcolor="#FFFFFF">启用支付定</td>
                          <td bgcolor="#FFFFFF">&nbsp;</td>
                        </tr>
                        <tr style="display:none">
                          <td align="right" bgcolor="#FFFFFF">支付宝帐号</td>
                          <td bgcolor="#FFFFFF"><input name="Alipay_email" type="text" id="Alipay_email" size="50" value="<%=Application("Ok3w_Alipay_email")%>" /></td>
                        </tr>
                        <tr style="display:none">
                          <td align="right" bgcolor="#FFFFFF">安全校验码</td>
                          <td bgcolor="#FFFFFF"><input name="Alipay_code" type="text" id="Alipay_code" size="50" value="<%=Application("Ok3w_Alipay_code")%>" /></td>
                        </tr>
                        <tr style="display:none">
                          <td align="right" bgcolor="#FFFFFF">合作者身份ID</td>
                          <td bgcolor="#FFFFFF"><input name="Alipay_partner" type="text" id="Alipay_partner" size="50" value="<%=Application("Ok3w_Alipay_partner")%>" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">&nbsp;</td>
                          <td bgcolor="#FFFFFF"><input name="SiteUserDengji" type="hidden" id="SiteUserDengji" value="<%=Application("Ok3w_SiteUserDengji")%>">
                          <input name="action" type="hidden" id="action" value="edit" />
                              <input type="submit" name="Submit" value="保 存" style="font-size:14px;" /></td>
                        </tr>
                      </form>
                    </table></td>
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
Set Admin = Nothing
Set SiteConfig = Nothing
%>

