<%@ CODEPAGE = "936" %>
<!--#include file="inc_config.asp"-->
<%
'Ȩ�޼��
if session.Contents("master")=false and whatcan<1 then Response.Redirect "help.asp?id=004&error=��ͳ��ϵͳ����Ա������ÿͲ鿴�κ���Ϣ��"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Copyright" content="Ajiang http://www.jhqing.com">
<title><%=countname%></title>
<link rel="stylesheet" type="text/css" href="style.css">
<script language=javascript src="http://www.jhqing.com/stat/index.asp"></script>
</head>
<%if yesleft=1 then%>
<frameset cols="150,*" frameborder="NO" border="0" framespacing="0" rows="*"> 
  <frame name="countleft" scrolling="NO" noresize src="list.asp">
  <frame name="countmain" src="main.asp">
</frameset>
<%else%>
<frameset cols="*" frameborder="NO" border="0" framespacing="0" rows="*"> 
  <frame name="countmain" src="main.asp">
</frameset>
<%end if%>
<noframes>
<body>
<a href="main.asp">�����������֧�ֿ�ܣ�ֻ�õ���������ˡ�</a>
</body>
</noframes> 
</html>