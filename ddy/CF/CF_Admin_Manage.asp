<%
'�˷���û�������
'����QQ��178575
'����EMail��yliangcf@163.com
'������վ��http://www.qqcf.com
'��ϸ��飺http://www.qqcf.com/?action=list&list=cfcount
'�����г���������ʾ����װ��ʾ��ʹ�����ѽ�����°汾���ص�����
'��Ϊ��Щ���ݿ���ʱ�����£���û�з��ڳ�������Լ�����վ�ϲ鿴
'�������汾����ʾ
%>
<%
Response.Expires= -1
Response.Addheader "pragma","no-cache"
Response.AddHeader "cache-control","no-store"
%>
<!--#include file="Conn.asp"-->
<!--#include file="CF_MyFunction.asp"-->
<!--#include file="CF_Md5.asp"-->
<!--#Include File="CF_Getcookie.asp"-->
<%
If Session("CFCountAdmin")="" Then
 Response.Cookies("CFCountAdminCookie")=""
 Call AlertUrl("�����µ�½!","Admin.asp")
End if

Action=Request("Action")
%>
<!--#Include File="CF_Admin_Manage_Do.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HTML><HEAD><TITLE><%=RsSet("SysTitle")%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<!--#include file="CF_Style.asp"-->
<%

If RsSet("emailsendtype")=1 Then
 show=show&"show_1();"
ElseIf RsSet("emailsendtype")=2 Then
 show=show&"show_2();"
End If
%>
</HEAD>
<BODY>
<!--#include file="Top.asp"-->

<table width="980px" align="center"  style="margin:0px;">
  <tr> 
    <td valign="top" width="202">
<!--#Include File="CF_Admin_Manage_Menu.asp"-->
</td>
<td width="778" valign="top">
<!--#Include File="CF_Admin_Manage_Select.asp"-->
</td>
</tr>
</table>

<!--#include file="Bottom.asp"-->
</BODY></HTML>

<%Call ConnClose()%>