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
if session.Contents("master")=false and whatcan<2 then Response.Redirect "help.asp?id=004&error=��û�в鿴�ܷ���ͳ�Ƶ�Ȩ�ޡ�"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Copyright" content="Ajiang http://www.jhqing.com">
<title><%=countname%>���ܷ���ͳ��</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<!--#include file="inc_top.asp"-->
<br>
<table width=500 cellspacing=0 align=center>
  <tr><td>
    <p style="line-height: 100%; margin-left: 15; margin-top: 5; margin-bottom: 0">
    Tips: ������ָͼ��������ͼ�����µ����ֿ��Կ�����Ӧ�ķ�������</p>
  </td></tr>
</table>
<br>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"></td>
    <td width="498" class="backq">
    
		&nbsp; <img src="images/tb_title.gif" align=absmiddle> &nbsp;�ˡˡ� �ܷ�����ͳ�� �ˡˡ�<br>

<table width="90%" align=center><tr><td>

<table border="0" cellpadding="0" cellspacing="0" width="175" align=center>
<tr height="9"><td></td></tr>
<tr height="101">
<%
set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath(connpath)
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath

'�ҵ���ʼͳ�������������������7�죬������ǰ��Ŀռ�
tmprs=conn.execute("Select first(vtime) as vfirst from view")
vfirst=tmprs("vfirst")
if isnull(vfirst) then vfirst=now()
vdays=int(date()-vfirst+1)

'������ά���飬voutday(*,0)Ϊ������,voutday(*,1)Ϊ����,voutday(*,2)Ϊ����
dim vday(7,3),voutday(7,3)
maxday=0
sumday=0
for i=0 to 6
	vday(i,0)=vdaycon(date()-6+i)
	if vday(i,0)>maxday then maxday=vday(i,0)
	sumday=sumday+vday(i,0)
	vday(i,1)=day(date()-6+i)
	vday(i,2)=weekday(date()-6+i)
next
'��ֹ����Ϊ0������
if maxday=0 then maxday=1
if sumday=0 then sumday=1

'������ͳ����������ֵ����
if vdays>=7 then
	for i=0 to 6
		voutday(i,0)=vday(i,0)
		voutday(i,1)=vday(i,1)
		voutday(i,2)=vday(i,2)
	next
else
	for i=0 to 6
		if i<=vdays then
			voutday(i,0)=vday(i+6-vdays,0)
			voutday(i,1)=vday(i+6-vdays,1)
			voutday(i,2)=vday(i+6-vdays,2)
		else
			voutday(i,0)=0
			voutday(i,1)=""
			voutday(i,2)=0
		end if
	next
end if
%>
<td align=right valign=top>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13">
<font face="Arial"><%=int(maxday*10+0.5)/10%></font>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13">
<font face="Arial"><%=int(3*maxday*10/4+0.5)/10%></font>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13">
<font face="Arial"><%=int(maxday*10/2+0.5)/10%></font>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0">
<font face="Arial"><%=int(maxday*10/4+0.5)/10%><br></font></td>
<td width=10><img src="images/tu_back_left.gif"></td>
<%
for i= 0 to 6
%>
<td width=15 valign=bottom background="images/tu_back.gif" align=center>
<img style="BORDER-BOTTOM: #000000 1px solid" src="images/tu.gif"
	height="<%=(voutday(i,0)/maxday)*100%>" width="9"
	alt="<%=voutday(i,1)%>�գ�����<%=findweek(cint(voutday(i,2)))%>������<%=voutday(i,0)%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(voutday(i,0)*1000/sumday+0.5)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"></td>
<%
next
%>
<td width=10><img src="images/tu_back_right.gif"></td>
<td width=10></td>
</tr>

<tr height="18">
<td align=right>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0">
<font face="Arial">0</font></td>
<td width=10></td>
<%
for i= 0 to 6
%>
<td width=15 align=center>
<a title="<%=voutday(i,1)%>�գ�����<%=findweek(cint(voutday(i,2)))%>������<%=voutday(i,0)%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(voutday(i,0)*1000/sumday+0.5)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"><%
'���ݵ���������ò�ͬ����ɫ�������ڣ���������ɫ�������ú�ɫ
select case voutday(i,2)
case 1%>
<font face="Arial" style="letter-spacing: -1" color="red">
<%case 7%>
<font face="Arial" style="letter-spacing: -1" class="fonts">
<%case else%>
<font face="Arial" style="letter-spacing: -1">
<%end select%>
<%=findweek(voutday(i,2))%></font></a></td>
<%
next
%>
<td width=10></td>
<td width=10></td>
</tr>
<tr height="5"><td></td></tr>
</table>

</td><td>

<table border="0" cellpadding="0" cellspacing="0" width="175" align=center>
<tr height="9"><td></td></tr>
<tr height="101">
<%
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select vweek,count(id) as allweek from view group by vweek"
rs.Open sql,conn,1,1

dim vallweek(7)
maxallweek=0
sumallweek=0
do while not rs.EOF
	vallweek(cint(rs("vweek"))-1)=cint(rs("allweek"))
	if vallweek(cint(rs("vweek"))-1)>maxallweek then maxallweek=vallweek(cint(rs("vweek"))-1)
	sumallweek=sumallweek+vallweek(cint(rs("vweek"))-1)
	rs.MoveNext
loop
'��ֹ����Ϊ0������
if maxallweek=0 then maxallweek=1
if sumallweek=0 then sumallweek=1

%>
<td align=right valign=top>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13">
<font face="Arial"><%=int(maxallweek*10+0.5)/10%></font>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13">
<font face="Arial"><%=int(3*maxallweek*10/4+0.5)/10%></font>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13">
<font face="Arial"><%=int(maxallweek*10/2+0.5)/10%></font>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0">
<font face="Arial"><%=int(maxallweek*10/4+0.5)/10%><br></font></td>
<td width=10><img src="images/tu_back_left.gif"></td>
<%
for i= 0 to 6
%>
<td width=15 valign=bottom background="images/tu_back.gif" align=center>
<img style="BORDER-BOTTOM: #000000 1px solid;" src="images/tu.gif"
	height="<%=(vallweek(i)/maxallweek)*100%>" width="9"
	alt="����<%=findweek(i+1)%>������<%=vallweek(i)%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vallweek(i)*1000/sumallweek)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"></td>
<%
next
%>
<td width=10><img src="images/tu_back_right.gif"></td>
<td width=10></td>
</tr>
<tr height="18">
<td align=right>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0">
<font face="Arial">0</font></td>
<td width=10></td>
<%
for i= 0 to 6
%>
<td width=15 align=center><a title="����<%=findweek(i+1)%>������<%=vallweek(i)%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vallweek(i)*1000/sumallweek)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%">
	<font face="Arial" style="letter-spacing: -1"><%=findweek(i+1)%></font></a></td>
<%
next
%>
<td width=10></td>
<td width=10></td>
</tr>
<tr height="5"><td colspan=29></td></tr>
</table>

</td></tr>
<tr height="20" align=center><td>�� �����һ��</td><td>�� ȫ��ʱ��</td></tr>
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

'��������ŷ���Ϊ����
function findweek(theweek)
	select case theweek
	case 1
		findweek="��"
	case 2
		findweek="һ"
	case 3
		findweek="��"
	case 4
		findweek="��"
	case 5
		findweek="��"
	case 6
		findweek="��"
	case 7
		findweek="��"
	end select
end function
%>