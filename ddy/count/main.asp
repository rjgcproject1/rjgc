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

set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath(connpath)
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath

'�����ܷ���������ʼ�������ڣ��������ݿ��ȡ��
    tmprs=conn.execute("Select count(id) As vtotal, first(vtime) as vfirst from view")
    vtotal=tmprs("vtotal")
    vfirst=tmprs("vfirst")
	if isnull(vtotal) then vtotal=0
	if isnull(vfirst) then vfirst=now()

if vtotal=0 then
	conn.Close
	set conn=nothing
	Response.Redirect "help.asp?error=ͳ��ϵͳ��û�����ã��в��ܲ鿴ͳ�Ʊ��档"
end if

'���շ����������շ��������Ӽ����ݳ���ȡ��
    tmprs=conn.execute("Select today,yesterday from vjian")
    vtoday=tmprs("today")
    vyesterday=tmprs("yesterday")
	if isnull(vtoday) then vtoday=0
	if isnull(vyesterday) then vyesterday=0
'���������
    tmprs=conn.execute("Select count(id) as vthisyear from view where vyear=" & year(now))
    vthisyear=tmprs("vthisyear")
	if isnull(vthisyear) then vthisyear=0
'���·�����
    tmprs=conn.execute("Select count(id) as vthismonth from view where vmonth=" & month(now) & " and vyear=" & year(now))
    vthismonth=tmprs("vthismonth")
	if isnull(vthismonth) then vthismonth=0
'�������顢ƽ��ÿ�������
	vdays=now()-vfirst
	vdayavg=vtotal/vdays
	vdays=int((vdays*10^mPrecision)+0.5)/(10^mPrecision)
	if vdays<1 then vdays="0" & vdays
	vdayavg=int((vdayavg*10^mPrecision)+0.5)/(10^mPrecision)
'Ԥ�ƽ��շ�����
	vdaylong=now()-date()
	vguess=int(((vtoday/vdaylong)+vyesterday)/2+0.5)
	if vguess< vtoday then vguess=int((vtoday/vdaylong)+0.5)
'��ǰ�û�����
	vuser=cint(Request.Cookies(mNameEn)("lao"))

conn.Close
set conn=nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Copyright" content="Ajiang http://www.jhqing.com">
<title><%=countname%></title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<!--#include file="inc_top.asp"-->
<br>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="4"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"></td>
    <td colspan="2" width="489"class="backq">
		&nbsp; <img src="images/tb_title.gif" align=absmiddle> &nbsp;�ˡˡ� �� �� �� �� �ˡˡ� 
	</td>
    <td width="1" class="backs"></td>
  </tr>
  <tr height="50">
    <td width="1" class="backs"></td>
    <td width="220"class="backq" valign=top>
		<p style="line-height: 140%; margin-left: 25; margin-right: 0; margin-top: 10; margin-bottom: 5" align=left>
		�ܷ�����: &nbsp; <%=vtotal%><br>
		���ķ�����: <%=vuser%><br>
		��ʼͳ����: <%=vfirst%><br>
		���շ�����: <%=vtoday%><br>
		���շ�����: <%=vyesterday%><br>
		���������: <%=vthisyear%><br>
		���·�����: <%=vthismonth%><br>
		ͳ������: &nbsp; <%=vdays%><br>
		ƽ���շ���: <%=vdayavg%><br>
		Ԥ�ƽ���: &nbsp; <%=vguess%><br>
	</td>
    <td width="278"class="backq" valign=top>
		<p style="line-height: 140%; margin-left: 10; margin-right: 0; margin-top: 10; margin-bottom: 0" align=left>
		��վ: <%=mName%><br>
		����: <a href="<%=mURL%>" class="a1" target="_blank"><%=mURL%></a><br>
		վ��: <%=masterName%><br>
		����: <a href="mailto:<%=masterEmail%>" class="a1"><%=masterEmail%></a><br>
		<p style="text-indent: -35; line-height: 140%; margin-left: 45; margin-right: 25; margin-top: 0; margin-bottom: 5" align=left>���: <%=SiteBrief%><br>
	</td>
    <td width="1" class="backs"></td>
  </tr>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<%if session.Contents("master")=false then%>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
<form action="login.asp"  method=post target="_top" id=form1 name=form1>
  <tr height="50">
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
    <td width=498 class="backq">
&nbsp;&nbsp; <img src="images/tb_login.gif" align=absmiddle> ����Ա��¼
&nbsp; &nbsp; �û�:<input name="logname" class="input" size=9>
&nbsp;����:<input type="password" name="logpsw" class="input" size=9>
&nbsp; &nbsp;��¼ <input type="submit" name="search" class="backc2" value=" ">
	</td>
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
  </tr>
</form>
  <tr><td colspan="3"><img src="images/photodown.gif"></td></tr>
</table>
<%else%>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
<form action="login.asp"  method=post target="_top" id=form1 name=form1>
  <tr height="50">
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
    <td width=498 class="backq">
&nbsp;&nbsp; <img src="images/tb_login.gif" align=absmiddle> ����Ա��¼
&nbsp; &nbsp; ����Ա�ѵ�¼��&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
<a href='logout.asp' target="_top">�˳� <img src="images/arbutton.gif" align="absmiddle" border="0"></a>	</td>
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
  </tr>
</form>
  <tr><td colspan="3"><img src="images/photodown.gif"></td></tr>
</table>
<%end if%>
<%if yestop=0 and yesleft=0 then%>
<br>
<!--#include file="inc_index.asp"-->
<%end if%>
<!--#include file="inc_bottom.asp"-->
</body>
</html>
