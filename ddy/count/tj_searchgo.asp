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
if session.Contents("master")=false and whatcan<4 then Response.Redirect "help.asp?id=004&error=��û��ʹ���Զ���ͳ�Ƶ�Ȩ�ޡ�"

'********************************** �����ѯ���� **********************************

'================================
'
'     ��������������ݲ�����
'
'================================

'���������ȡ����

onyear		=trim(Request("onyear"))
onmonth		=trim(Request("onmonth"))
onday		=trim(Request("onday"))

offyear		=trim(Request("offyear"))
offmonth	=trim(Request("offmonth"))
offday		=trim(Request("offday"))

vip			=trim(Request("vip"))
vwhere		=trim(Request("vwhere"))
vos1		=trim(Request("vos1"))
vos2		=trim(Request("vos2"))
vsoft1		=trim(Request("vsoft1"))
vsoft2		=trim(Request("vsoft2"))
vcome		=trim(Request("vcome"))
vpage		=trim(Request("vpage"))

outtype		=trim(Request("outtype"))
wherestr	=trim(Request("wherestr"))
name		=trim(Request("name"))
content		=trim(Request("content"))

'�����ȡ������
ontime=onyear & "-" & onmonth & "-" & onday
if (ontime <> "--") and (not isdate(ontime)) then Response.Redirect "help.asp?id=001&error=������д���Ϻ�Ҫ��"

offtime=offyear & "-" & offmonth & "-" & offday
if (offtime <> "--") and (not isdate(offtime)) then Response.Redirect "help.asp?id=001&error=������д���Ϻ�Ҫ��"
if (offtime <> "--") then offtime=cdate(offtime)+1		'��ΪSQL���ǰ���ָ���ڵ�0ʱ����ģ�����Ӧ�ø�Ϊ��һ��

if (outtype ="") then Response.Redirect "help.asp?id=001&error=��û��ָ��Ҫ�鿴����Ŀ��"

vos=vos2
if vos = "" then vos=vos1
vsoft=vsoft2
if vsoft="" then vsoft=vsoft1


'************************************ �������� ************************************

'================================
'
'         ת��Ϊ SQL ָ��
'
'================================

'����ѯ����ת��ΪSQL��ѯָ��
if wherestr="" then
wherestr=" where "

'������ڲ�ѯ�������򽫲�ѯ���������ѯ�ִ�
if ontime <> "--" then wherestr=wherestr & "and (vtime>=datevalue('" & ontime & "')) "
if offtime <> "--" then wherestr=wherestr & "and (vtime<=datevalue('" & offtime & "')) "
if vip <> "" then wherestr=wherestr & "and (vip like '%" & vip & "%') "
if vwhere <> "" then wherestr=wherestr & "and ((vwhere like '%" & vwhere & "%') or (vwheref like '%" & vwhere & "%')) "
if vos <> "" then vherestr=wherestr & "and (vos like '%" & vos & "%') "
if vsoft <> "" then wherestr=wherestr & "and (vsoft like '%" & vsoft & "%') "
if vcome <> "" then wherestr=wherestr & "and (vcome like '%" & vcome & "%') "
if vpage <> "" then wherestr=wherestr & "and (vpage like '%" & vpage & "%') "

'ȥ����һ����ѯ������ where ֮��� and
wherestr=replace(wherestr,"where and","where")

'���û�в�ѯ���������ѯ�ִ�ӦΪ��
if trim(wherestr)="where" then wherestr=""
end if

'���Ҫ�鿴������ֻ����ϸ��¼����ֱ��ת����ϸ��¼ҳ
if outtype="��ϸ" then Response.Redirect "tj_all.asp?wherestr=" & server.URLEncode(wherestr)

'================================
'
'          �� �� ��
'
'================================

'�����ݿ�
set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath(connpath)
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Copyright" content="Ajiang http://www.jhqing.com">
<title><%=countname%>���Զ���ͳ��</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<!--#include file="inc_top.asp"-->
<br>
<table width=500 cellspacing=0 align=center>
  <tr><td>
    <p style="line-height: 100%; margin-left: 15; margin-top: 5; margin-bottom: 0">
    ��ϲ: ������ʾ�ľ�����Ҫ��ѯ�����ݵķ��������</p>
  </td></tr>
</table>
<br>

<%
'���ݲ�ѯ���������Ӧ������

if instr(outtype,"Сʱ") then
statname="24Сʱ���ʷ���"
tuwidth=15
thewhich="vhour"
tuhow=24
unit="ʱ"
statheng statname,tuwidth,thewhich,tuhow,unit,wherestr
end if

if instr(outtype,"��") then
statname="�շ��ʷ���"
tuwidth=13
thewhich="vday"
tuhow=31
unit="��"
statheng statname,tuwidth,thewhich,tuhow,unit,wherestr
end if

if instr(outtype,"��") then
statname="�ܷ��ʷ���"
tuwidth=40
thewhich="vweek"
tuhow=7
unit=""
statheng statname,tuwidth,thewhich,tuhow,unit,wherestr
end if

if instr(outtype,"��") then
statname="�·��ʷ���"
tuwidth=20
thewhich="vmonth"
tuhow=12
unit="��"
statheng statname,tuwidth,thewhich,tuhow,unit,wherestr
end if

if instr(outtype,"��") then
statname="����ʷ���"
leftwidth=50
tuwhich="vyear"
statshu statname,leftwidth,tuwhich,wherestr
end if

if instr(outtype,"IP") then
statname="IP����"
leftwidth=150
tuwhich="vIP"
statshu statname,leftwidth,tuwhich,wherestr
end if

if instr(outtype,"����") then
statname="�������ʷ���"
leftwidth=100
tuwhich="vwhere"
statshu statname,leftwidth,tuwhich,wherestr
end if

if instr(outtype,"�����") then
statname="��ͬ������û����ʷ���"
tuwidth=50
thewhich="vsoft"
tuhow=7
unit=""
statheng statname,tuwidth,thewhich,tuhow,unit,wherestr
end if

if instr(outtype,"����ϵͳ") then
statname="��ͬ����ϵͳ�û����ʷ���"
tuwidth=50
thewhich="vOS"
tuhow=7
unit=""
statheng statname,tuwidth,thewhich,tuhow,unit,wherestr
end if

if instr(outtype,"��·") then
statname="��·����"
leftwidth=220
tuwhich="vcome"
statshu statname,leftwidth,tuwhich,wherestr
end if

if instr(outtype,"ҳ��") then
statname="ҳ��������"
leftwidth=220
tuwhich="vpage"
statshu statname,leftwidth,tuwhich,wherestr
end if

'���ͬʱѡ������ϸ������ʾ�������
if instr(outtype,"��ϸ") then
%>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <form action="tj_all.asp" method=post id=form1 name=form1>
  <tr height="30">
    <td width="1" class="backs"></td>
    <td width="498"class="backq">
		&nbsp; <img src="images/tb_title.gif" align=absmiddle><font style="font-size:16px">&nbsp;</font>�ˡˡ� �鿴��ϸ��¼ �ˡˡ�<br>
		<p class="p1">��Ϊ��ϸ��¼��Ҫ��ҳ�����ܺ����ϼ�¼����ͬһҳ����������ļ�����ť�鿴��Ӧ���ݡ�
	<p class="p1" style='margin-top: 0;' align="right">
	<a href='javascript:document.forms[0].submit();'>����</a>
	<input type="hidden" name="wherestr" value="<%=wherestr%>">
	<input type="submit" value=" " name="save" class="backc2"><font style="font-size:16px">&nbsp;</font>&nbsp;
	</td>
    <td width="1" class="backs"></td>
  </tr>
  </form>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<%
end if

'�������ݱ�
if session.Contents("master")=true or whatcan>=6 then
%>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <form action="tj_save.asp" method=post>
  <tr height="30">
    <td width="1" class="backs"></td>
    <td width="498"class="backq">
		&nbsp; <img src="images/tb_save.gif" align=absmiddle><font style="font-size:16px">&nbsp;</font>�ˡˡ� ������μ������� �ˡˡ�<br>

	<p class="p1"><font class="fonts"><b>��������</b></font>&nbsp; <%if wherestr="" then%>û�м�������<%else%><%=wherestr%><%end if%>��
	<p class="p1" style='margin-bottom: 0;margin-top: 0;'><font class="fonts"><b>��ѯ��Ŀ</b></font>&nbsp; <%=outtype%>��
	<p class="p1" style='line-height: 100%;margin-bottom: 0;margin-top: 0;'><font class="fonts"><b>ȡ������</b></font>&nbsp; <input name="name" size="16" class="input" value="<%=name%>">&nbsp;
	<INPUT type="radio" name="overwrite" value="0"<%if name="" then%> checked<%end if%>> ͬ��ʱ��ʾ&nbsp;
	<INPUT type="radio" name="overwrite" value="1"<%if name<>"" then%> checked<%end if%>> ͬ��ʱ����
	<p class="p1" style='line-height: 100%;margin-bottom: 0;margin-top: 0;'><font class="fonts"><b>�Ӹ�����</b></font>&nbsp; <input name="content" size="50" class="input" value="<%=content%>">
	<p class="p1" style='margin-top: 0;' align="right">
	<a href="help.asp?id=002">����</a>
	<a href='javascript:document.forms[0].submit();'>����</a>
	<input type="hidden" name="wherestr" value="<%=wherestr%>"><input type="hidden" name="outtype" value="<%=outtype%>"><input type="submit" value=" " name="save" class="backc2"><font style="font-size:16px">&nbsp;</font>&nbsp;

	</td>
    <td width="1" class="backs"></td>
  </tr>
  </form>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<%
end if
%>
<!--#include file="inc_bottom.asp"-->
</body>
</html>
<%
' �ر����ݿ�
set rs=nothing
conn.Close
set conn=nothing
%>





<%

'******************************** �Զ��庯�����ӳ��� ********************************

'================================
'
'   ���ͼ�����ݵ��ӳ���(��)
'
'================================
sub statheng(statname,tuwidth,tuwhich,thehow,unit,wherestr)

'statname	��������
'tuwidth	��ͼÿ�����
'thewhich	Ҫ��ѯ����Ŀ
'thehow		���ж���������
'unit		��Ŀ�ĵ�λ
'wherestr	��ѯ����

'������������������
dim val(),alt(),theper()
redim val(thehow),alt(thehow),theper(thehow)

'��ʼ����
sql="select " & thewhich & ",count(id) as theval from view " & wherestr & _
	" group by " & thewhich & " order by " & thewhich
rs.Open sql,conn,1,1

themax=0
thesum=0

for i=0 to tuhow-1
	if not rs.EOF then
		alt(i)=rs(thewhich)
		val(i)=rs("theval")
		if cint(val(i))>themax then themax=cint(val(i))
		thesum=thesum + cint(val(i))
		rs.MoveNext
	else
		alt(i)=""
		val(i)=0
	end if	
next
'��ֹ����Ϊ0������
if themax=0 then themax=1
if thesum=0 then thesum=1

'����ٷ���
for i=0 to tuhow-1
	theper(i)=int(val(i)*1000/thesum)/10
	if csng(theper(i))<1 then theper(i)="0" & theper(i)
next

'�����ѯ��Ŀ�����ڣ���תaltΪ����
if thewhich="vweek" then
  for i=0 to tuhow-1
    alt(i)=findweek(alt(i))
  next
end if
rs.Close

'����hengout�ӳ�������������
hengout statname,tuwidth,themax,tuhow,val,alt,theper,unit

end sub





'================================
'
'   ��ʾ����ͼ����ӳ���(��)
'
'================================

sub hengout(statname,tuwidth,themax,tuhow,val,alt,theper,unit)

'statname	��������
'tuwidth	��ͼÿ�����
'themax		x�������ֵ
'tuhow		��ͼ�ж�����
'val()		x����ֵ����
'alt()		x��������
'theper()	��ͼ�ٷ���
'unit		x���굥λ
%>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
    <td width="498"class="backq">
    
		&nbsp; <img src="images/tb_title.gif" align=absmiddle> &nbsp;�ˡˡ� <%=statname%> �ˡˡ�<br>
<table border="0" cellpadding="0" cellspacing="0" width="<%=tuhow * tuwidth + 70%>" align=center>
<tr height="9"><td></td></tr>
<tr height="101">
  <td align=right valign=top>
	<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"><font face="Arial"><%=int(themax*10+0.5)/10%></font>
	<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"><font face="Arial"><%=int(3*themax*10/4+0.5)/10%></font>
	<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"><font face="Arial"><%=int(themax*10/2+0.5)/10%></font>
	<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0"><font face="Arial"><%=int(themax*10/4+0.5)/10%><br></font></td>
  <td width=10><img src="images/tu_back_left.gif"></td>
<%
for i= 0 to tuhow-1
%>
  <td width="<%=tuwidth%>" valign=bottom background="images/tu_back.gif" align=center>
  <img style="BORDER-BOTTOM: #000000 1px solid;" src="images/tu.gif"
	height="<%=(val(i)/themax)*100%>" width="9"
	alt="<%=alt(i)%><%=unit%>������<%=val(i)%>�Σ�<%=theper(i)%>%"></td>
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
for i= 0 to tuhow-1
%>
  <td width="<%=tuwidth%>" align=center><a
	title="<%=alt(i)%><%=unit%>������<%=val(i)%>�Σ�<%=theper(i)%>%"><font
	face="Arial" style="letter-spacing: -1"><%=alt(i)%></font></a></td>
<%
next
%>
  <td width=10></td>
  <td width=10></td>
</tr>
<tr height="5"><td></td></tr>
</table>

	</td>
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
  </tr>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<%
end sub
%>

<%

'================================
'
'   ���ͼ�����ݵ��ӳ���(��)
'
'================================

sub statshu(statname,leftwidth,thewhich,wherestr)

'statname	��������
'leftwidth	��ͼÿ�����
'thewhich	Ҫ��ѯ����Ŀ
'wherestr	��ѯ����

'������������������
dim val(),alt(),theper(),alta()

on error resume next
'��ʼ����
sql="select " & thewhich & ",count(id) as theval from view " & wherestr & _
	" group by " & thewhich & " order by count(id) DESC"
rs.Open sql,conn,1,1

themax=0
thesum=0
tuhow=rs.recordcount
if tuhow>100 then tuhow=100

redim val(tuhow),alt(tuhow),theper(tuhow),alta(tuhow)
i=0
do while not rs.EOF
	alta(i)=rs(thewhich)
	if (thewhich = "vcome") or (thewhich = "vpage") then
		thelen=len(alta(i))
		if thelen =0 then
			alta(i)="#"
			alt(i)="ͨ���ղػ�ֱ��������ַ����"
		end if
		if thelen <= 33 and thelen > 0 then
			alt(i)=alta(i)
		end if
		if thelen >= 34 then
			alt(i)=left(alta(i),31) & "..."
		end if
	else
		alt(i)=alta(i)
	end if
	val(i)=rs("theval")
	if cint(val(i))>themax then themax=cint(val(i))
	thesum=thesum + cint(val(i))
	if i=100 then exit do
	i=i+1
	rs.MoveNext
loop
'��ֹ����Ϊ0������
if themax=0 then themax=1
if thesum=0 then thesum=1

'����ٷ���
for i=0 to tuhow-1
	theper(i)=int(val(i)*1000/thesum)/10
	if csng(theper(i))<1 then theper(i)="0" & theper(i)
next

'�����ѯ��Ŀ�����ڣ���תaltΪ����
if thewhich="vweek" then
  for i=0 to tuhow-1
    alt(i)=findweek(alt(i))
  next
end if
rs.Close

'����hengout�ӳ�������������
shuout statname,leftwidth,themax,tuhow,val,alt,alta,theper

end sub






'================================
'
'   ��ʾ����ͼ����ӳ���(��)
'
'================================

sub shuout(statname,leftwidth,themax,tuhow,val,alt,alta,theper)
'statname	��������
'leftwidth	ͼ������������ֵĿ��
'themax		x�������ֵ
'tuhow		�����м�¼
'val()		��¼��ֵ(������)
'alt()		˵������
'alta()		���ֵ�����
'theper()	��ͼ�ٷ���
%>

<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
    <td width="498"class="backq">
    
		&nbsp; <img src="images/tb_title.gif" align=absmiddle> &nbsp;�ˡˡ� <%=statname%> �ˡˡ�<br>

<table border="0" cellpadding="0" cellspacing="0" width="<%=leftwidth+230%>" align=center>
  <tr height="9"><td></td></tr>
  <tr height="10">
    <td width="<%=leftwidth%>"></td><td width="230"><img src="images/tu_back_up.gif"></td>
  </tr>
<%for i=0 to tuhow-1%>
  <tr>
    <td width="220" align=right><a href="<%=alta(i)%>" target="_blank"
	title="<%=alta(i) & vbcrlf%>����<%=val(i)%>�Σ�<%=theper(i)%>%"><%=alt(i)%></a>&nbsp;</td>
  <td width="230" background="images/tu_back_2.gif" align=left>
  <img style="BORDER-left: #000000 1px solid;" src="images/tu.gif"
	width="<%=(val(i)/themax)*183%>" height="9"
	alt="<%=alta(i) & vbcrlf%>������<%=val(i)%>�Σ�<%=theper(i)%>%"> <%=val(i)%></td>
</tr>
<%
next
%>
<tr height="10">
	<td width="220"></td><td width="230"><img src="images/tu_back_down.gif"></td>
</tr>
<tr height="5"><td></td></tr>
</table>

	</td>
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
  </tr>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<%
end sub


'================================
'
'   ������תΪ���ֵ��Զ��庯��
'
'================================
function findweek(theweek)
	select case cstr(theweek)
	case "1"
		findweek="������"
	case "2"
		findweek="����һ"
	case "3"
		findweek="���ڶ�"
	case "4"
		findweek="������"
	case "5"
		findweek="������"
	case "6"
		findweek="������"
	case "7"
		findweek="������"
	end select
end function
%>