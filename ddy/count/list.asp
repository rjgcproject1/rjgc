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
'Ȩ�޷���
if session.Contents("master")=false and whatcan<1 then Response.Redirect "help.asp?id=004&error=��ͳ��ϵͳ����Ա������ÿͲ鿴�κ���Ϣ��"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Copyright" content="Ajiang http://www.jhqing.com">
<meta http-equiv="refresh" content="300; url='list.asp'">
<title><%=countname%></title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body style="BACKGROUND-IMAGE:url('images/leftback.gif')" class=backq>
<table>
<tr height="240"><td valign=top>
<p style="margin-left: 17;margin-right: 0;	margin-top: 62;	margin-bottom: 5">
<%if session.Contents("master")=true or whatcan>0 then%><a href="main.asp" target="countmain">�� ��������</a><%end if%>
<%if session.Contents("master")=true or whatcan>2 then%><br><a href="tj_all.asp" target="countmain">�� ��ϸ��¼</a><%end if%>
<%if session.Contents("master")=true or whatcan>1 then%><br><a href="tj_hour.asp" target="countmain">�� 24Сʱͳ��</a>
<br><a href="tj_day.asp" target="countmain">�� ��ͳ��</a>
<br><a href="tj_week.asp" target="countmain">�� ��ͳ��</a>
<br><a href="tj_month.asp" target="countmain">�� ��ͳ��</a>
<br><a href="tj_come.asp" target="countmain">�� ��·ͳ��</a>
<br><a href="tj_page.asp" target="countmain">�� ����ҳ��</a>
<br><a href="tj_ip.asp" target="countmain">�� IPͳ��</a>
<br><a href="tj_soft.asp" target="countmain">�� �ͻ������</a>
<br><a href="tj_where.asp" target="countmain">�� �����ߵ���</a><%end if%>
</td></tr>
<tr height=60><td valign=top>
<%if session.Contents("master")=true or whatcan>3 then%>
<p style="margin-left: 17;margin-right: 0;	margin-top: 5;	margin-bottom: 5"><a href="tj_search.asp" target="countmain">�Զ���ͳ��</a>
<%set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath(connpath)
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")

sql="select top 2 * from save"
rs.Open sql,conn,1,1
do while not rs.EOF
%>
<br><a href="tj_searchgo.asp?wherestr=<%=server.URLEncode(rs("wherestr"))%>&outtype=<%=server.URLEncode(rs("outtype"))%>&name=<%=server.URLEncode(rs("name"))%>&content=<%=server.URLEncode(rs("content"))%>" target="countmain">�� <%=left(rs("name"),6)%></a>
<%
rs.MoveNext
loop
rs.Close
set rs=nothing
conn.Close
set conn=nothing

end if%>
<br>&nbsp; &nbsp; &nbsp; &nbsp; <a href="tj_search.asp" target="countmain">����...</a>
</td></tr>
<%if session.Contents("master")=false then%>
<form action="login.asp"  method=post target="_top">
<tr><td>
<br>&nbsp;<img src="images/tb_login.gif" align=absmiddle> <font color=white>����Ա��¼
<br>&nbsp;�û�:<input name="logname" class="input" size=9>
<br>&nbsp;����:<input type="password" name="logpsw" class="input" size=9></font>
<input type="submit" name="search" class="backc2" value=" ">
</td></tr>
</form>
<%else%>
<tr><td>
<br>&nbsp;<img src="images/tb_login.gif" align=absmiddle> <font color=white>����Ա�ǳ�
	<br>&nbsp; ����Ա�ѵ�¼��
<br>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; <a  target="_top" href='logout.asp' class=a2>�˳� <img src="images/arbutton.gif" align="absmiddle" border="0"></a>
</td></tr>
<%end if%>
</table>
</body>
</noframes> 
</html>
