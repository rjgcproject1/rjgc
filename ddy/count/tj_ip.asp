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
if session.Contents("master")=false and whatcan<2 then Response.Redirect "help.asp?id=004&error=��û�в鿴IPͳ�Ƶ�Ȩ�ޡ�"

%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Copyright" content="Ajiang http://www.jhqing.com">
<title><%=countname%>��IPͳ��</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<!--#include file="inc_top.asp"-->
<br>
<table width=500 cellspacing=0 align=center>
  <tr><td>
    <p style="line-height: 100%;text-indent: -35; margin-left: 50; margin-top: 5; margin-bottom: 10">
    Tips: ������ָͼ����������ַ���Կ�����Ӧ�ķ�������Ҫ�õ�ĳһʱ�ε�ͳ����Ϣ����ʹ���Զ���ͳ�ơ�</p>
  </td></tr>
</table>
<%
set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath(connpath)
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")
%>

<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"></td>
    <td width="498"class="backq">
    
		&nbsp; <img src="images/tb_title.gif" align=absmiddle> &nbsp;�ˡˡ� IP��ַ�������� �ˡˡ�<br>

<table border="0" cellpadding="0" cellspacing="0" width="350" align=center>
<tr height="9"><td></td></tr>
<tr height="10">
	<td width="120"></td><td width="230"><img src="images/tu_back_up.gif"></td>
</tr>
<%

sql="select vip,count(id) as allip from view group by vip order by count(id) DESC"
rs.Open sql,conn,1,1

maxallip=0
sumallip=0
do while not rs.EOF
	if cint(rs("allip"))>maxallip then maxallip=cint(rs("allip"))
	sumallip=sumallip+cint(rs("allip"))
	rs.MoveNext
loop
'��ֹ����Ϊ0����
if maxallip=0 then maxallip=1
if sumallip=0 then sumallip=1

rs.MoveFirst 

j=0
do while not rs.EOF
theip=rs("vip")
vallip=rs("allip")
	thelen=len(theip)
	if thelen =0 then
		theip="main.asp"
		svip="ͨ���ղػ�ֱ��������ַ����"
	end if
	if thelen <= 33 and thelen > 0 then
		svip=theip
	end if
	if thelen >= 34 then
		svip=left(theip,31) & "..."
	end if
%>
<tr>
<td width="120" align=right><a title="<%=theip%>������<%=vallip%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vallip*1000/sumallip)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"><%=svip%></a>&nbsp;</td>
<td width="230" background="images/tu_back_2.gif" align=left>
<img style="BORDER-left: #000000 1px solid;" src="images/tu.gif"
	width="<%=(vallip/maxallip)*183%>" height="9"
	alt="<%=theip%>������<%=vallip%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vallip*1000/sumallip)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"> <%=vallip%></td>
</tr>
<%
rs.MoveNext
	j=j+1
	'�����¼����40�������˳�
	if j=40 then exit do
loop
%>
<tr height="10">
	<td width="120"></td><td width="230"><img src="images/tu_back_down.gif"></td>
</tr>

<tr height="5"><td colspan=29></td></tr>
</table>

	</td>
    <td width="1" class="backs"></td>
  </tr>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<%
rs.Close

set rs=nothing
conn.Close 
set conn=nothing
%>
<br>
<!--#include file="inc_bottom.asp"-->
</body>
</html>
<%
'����ָ�����ڵķ�����
function vdaycon(theday)
    theday=cdate(theday)
    thetday=cdate(theday+1)
    tmprs=conn.execute("Select count(id) as vdaycon from view where" & _
		" vtime>=datevalue('" & theday & "') and vtime<=datevalue('" & thetday & "')")
    vdaycon=tmprs("vdaycon")
	if isnull(vdaycon) then vdaycon=0
end function
%>