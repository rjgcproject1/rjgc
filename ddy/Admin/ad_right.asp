<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%dbdns="../"%>
<!--#include file="chk.asp"-->
<%
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

Function DbSize()
	On Error ReSume Next
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFile(Server.MapPath(dbdns & SysSiteDbPath))
	DbSize = f.Size
	Set f = Nothing
	Set fso = Nothing
	On Error Goto 0
End Function
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��̨����ϵͳ</title>
<link rel="stylesheet" type="text/css" href="images/Style.css">
<style type="text/css">
<!--
.STYLE1 {
	color: #0000FF;
	font-weight: bold;
	font-size: 14px;
}
.STYLE2 {
	color: #FF0000;
	font-size: 14px;
}
-->
</style>
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
            <td><table height="22" cellspacing="0" cellpadding="0" border="0">
              <tbody>
                <tr>
                  <td width="3"><img id="tabImgLeft__0" height="22" 
                  src="images/tab_active_left.gif" width="3" /></td>
                  <td 
                background="images/tab_active_bg.gif" class="tab"><strong class="mtitle">��վ�����̨ &gt;&gt; ������ҳ </strong></td>
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
            <td width="100%" 
          valign="top" 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><table width="100%" cellpadding="0" cellspacing="0">
              <tr>
                <td height="100"><span class="STYLE1">&gt;&gt;��ӭ��ʹ����վ��̨����ϵͳ&lt;&lt;</span></td>
              </tr>
              <tr>
                <td><table border="0" cellpadding="5" cellspacing="0">
                    <tbody>
                      <tr>
                        <td align="right"><strong>����汾��</strong></td>
                        <td> 1.0 </td>
                      </tr>
                      <tr>
                        <td align="right"><strong>��Ȩ���У�</strong></td>
                        <td>�й����ʴ�ѧ���人����Ϣ����ѧԺ�������ϵ</td>
                      </tr>
                      <tr>
                        <td align="right"><strong>��ƿ�����</strong></td>
                        <td>������������ޡ���������ѡ���ǿ</td>
                      </tr>
                      <tr>
                        <td align="right"><strong>��ϵQQ��</strong></td>
                        <td>779351528</td>
                      </tr>
                      <tr>
                        <td align="right"><strong>��ϵ���䣺</strong></td>
                        <td>jieqiuming.2007@163.com</td>
                      </tr>
                    </tbody>
                </table></td>
              </tr>
            </table></td>
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
<br>
</body>
</html>
<%
Call CloseConn()
Set Rs = Nothing
Set Admin = Nothing
%>

