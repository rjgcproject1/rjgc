<%@ CODEPAGE = "936" %>
<%
'=================================
'
'      �ෲ��վ����ͳ��ϵͳ
'    Ajiang   zhqi123@163.com 
'        www.jhqing.com
'
'     ��Ȩ���С���ϮŲ�ñؾ�
'
'=================================
%>
<!--#include file="inc_config.asp"-->
<%
'Ȩ�޼��
if session.Contents("master")=false and whatcan<6 then Response.Redirect "help.asp?id=004&error=��û�й����Զ���ͳ�Ƽ����������Ȩ�ޡ�"

'��ȡ����
name		=trim(Request("name"))
content		=trim(Request("content"))
wherestr	=trim(Request("wherestr"))
outtype		=trim(Request("outtype"))
overwrite	=cbool(trim(Request("overwrite")))

'���˵�����
name=replace(name,"'","")

'�������
if name="" then Response.Redirect "help.asp?id=002&error=��ΪҪ�洢�ļ�������ȡ�����֡�"
if outtype="" then Response.Redirect "help.asp?id=002&error=û��ָ��Ҫ��������ݣ���ȷ�����ڼ������ҳ�㱣����뱾ҳ�ġ�"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Copyright" content="Ajiang http://www.jhqing.com">
<title><%=countname%>�������Զ���ͳ�Ƶļ�������</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<!--#include file="inc_top.asp"-->
<br>


<%
'�����ݿ�
set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath(connpath)
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")

sql="select * from save where name='" & name & "'"
rs.Open sql,conn,3,2

if (not rs.EOF) and (not overwrite) then
%>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <form action="tj_save.asp" method=post id=form1 name=form1>
  <tr height="30">
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
    <td width="498"class="backq">
		&nbsp; <img src="images/tb_save.gif" align=absmiddle><font style="font-size:16px">&nbsp;</font>�ˡˡ� ����������� �ˡˡ�<br>

	<p class="p1"><img src='images/tb_error.gif' align=absmiddle><font style='font-size:15px'>&nbsp;</font><font color=red>ע�⣡��</font>
	<p class="p1">���ݿ����Ѿ�����һ����Ϊ��<%=name%>���ļ���������¼��
	<p class="p1" style='margin-bottom: 0;'><font class="fonts">��������</font>&nbsp;
	<%if rs("wherestr")="" then%>û�м�������<%else%><%=rs("wherestr")%><%end if%>��
	<p class="p1" style='margin-top: 0;margin-bottom: 0;'><font class="fonts">��ѯ��Ŀ</font>&nbsp; <%=rs("outtype")%>��
	<p class="p1" style='margin-top: 0;'><font class="fonts">˵������</font>&nbsp; <%=rs("content")%>
	<p class="p1">����Ҫ����ļ�������Ϊ��
	<p class="p1" style='margin-bottom: 0;'><font class="fonts">��������</font>&nbsp;
	<%if wherestr="" then%>û�м�������<%else%><%=wherestr%><%end if%>��
	<p class="p1" style='margin-top: 0;margin-bottom: 0;'><font class="fonts">��ѯ��Ŀ</font>&nbsp; <%=outtype%>��
	<p class="p1" style='margin-top: 0;'><font class="fonts">˵������</font>&nbsp; <%=content%>
	<p class="p1">��ѡ�����Ĵ���취����
	<p class="p1"><INPUT type="radio" name="overwrite" value="0" checked> ����������ָ�Ϊ&nbsp;
	<input name="name" size="25" class="input" value="<%=name%>">
	<p class="p1" style='margin-top: 0;'><INPUT type="radio" name="overwrite" value="1"> ����ԭ����������������ύ��Ϊ׼
	<p class="p1" style='margin-top: 0;' align="right">
	<a href="help.asp?id=002">����</a>
	<a href='javascript:history.back()'>ȡ��</a> 
	<a href='javascript:document.forms[0].submit();'>����</a>
	<input name="content" size="50" type="hidden" value="<%=content%>"><input type="hidden" name="wherestr" value="<%=wherestr%>"><input type="hidden" name="outtype" value="<%=outtype%>"><input type="submit" value=" " name="save" class="backc2"><font style="font-size:16px">&nbsp;</font>&nbsp;
	</td>
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
  </tr>
  </form>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<%
rs.Close

else

	if rs.EOF then rs.AddNew

	rs("wherestr")=wherestr
	rs("outtype")=outtype
	rs("name")=name
	rs("content")=content
	
	rs.Update
	rs.Close
%>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <form action="tj_save.asp" method=post id=form1 name=form1>
  <tr height="30">
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
    <td width="498"class="backq">
		&nbsp; <img src="images/tb_save.gif" align=absmiddle><font style="font-size:16px">&nbsp;</font>�ˡˡ� ������������ɹ� �ˡˡ�<br>

	<p class="p1"><font color=red>��ϲ����</font>
	<p class="p1">�Ѿ��ɹ��ı�������μ�����������
	<p class="p1"><font class="fonts">��������</font>&nbsp; <%=name%>
	<p class="p1" style='margin-top: 0;'><font class="fonts">��������</font>&nbsp;
	<%if wherestr="" then%>û�м�������<%else%><%=wherestr%><%end if%>��
	<p class="p1" style='margin-top: 0;'><font class="fonts">��ѯ��Ŀ</font>&nbsp; <%=outtype%>��
	<p class="p1" style='margin-top: 0;'><font class="fonts">˵������</font>&nbsp; <%=content%>
	<p class="p1" align="right"><a href='tj_search.asp'>����</a> <a href='tj_search.asp'><img src="images/arbutton.gif" align="absmiddle" border="0"></a> <font style="font-size:16px">&nbsp;</font>&nbsp;
	</td>
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
  </tr>
  </form>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<%
end if

set rs=nothing
conn.Close
set conn=nothing
%>
<!--#include file="inc_bottom.asp"-->
</body>
</html>
