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
if session.Contents("master")=false and whatcan<4 then Response.Redirect "help.asp?id=004&error=��û�в鿴�Զ��������¼�ͽ����Զ�������ĵ�Ȩ�ޡ�"
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
    Tips: ��������������д���Ĳ�ѯ������֧����ϲ�ѯ��ģ����ѯ��</p>
  </td></tr>
</table>
<br>
<%
'���Ȩ��
if session.Contents("master")=true or whatcan>=5 then
%>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"></td>
    <td width="498"class="backq">
<%
dim OS(7)
	OS(0)="Win 2000"
	OS(1)="Win XP"
	OS(2)="Win NT"
	OS(3)="Win 9x"
	OS(4)="��Unix"
	OS(5)="Mac"
	OS(6)="Other"

dim soft(7)
	soft(0)="NetCaptor"
	soft(1)="MSIE 6.x"
	soft(2)="MSIE 5.x"
	soft(3)="MSIE 4.x"
	soft(4)="Netscape"
	soft(5)="Opera"
	soft(6)="Other"
%>
		&nbsp; <img src="images/tb_title.gif" align=absmiddle> &nbsp;�ˡˡ� �Զ���ͳ����Ŀ �ˡˡ�<br>
<table border="0" cellpadding="0" cellspacing="0" width="430" align=center>
<form action="tj_searchgo.asp" method="post" name="search">
<tr height="9"><td></td></tr>
<tr><td>
	ʱ �� �Σ�<input name="onyear" class="input" size="5"> ��
			  <input name="onmonth" class="input" size="3"> ��
			  <input name="onday" class="input" size="3"> ��
		   �� <input name="offyear" class="input" size="5"> ��
			  <input name="offmonth" class="input" size="3"> ��
			  <input name="offday" class="input" size="3"> ��
<br>�ɣе�ַ��<input name="vip" class="input" size="16">
<br>���Ե�����<input name="vwhere" class="input" size="25">
<br>�� �� �ˣ�<SELECT name="vOS1" class="input" style="width:70" id="vos1">			  <OPTION value="" selected>����ϵͳ</OPTION><%for i=0 to 6%>			  <OPTION value="<%=OS(i)%>"><%=OS(i)%></OPTION><%next%>
			  </SELECT>
			  <INPUT name="vOS2" class="input" size="10" id="vos">&nbsp; 

			  <SELECT name="vsoft1" class="input" style="width:70">			  <OPTION value="" selected>�����</OPTION><%for i=0 to 6%>			  <OPTION value="<%=soft(i)%>"><%=soft(i)%></OPTION><%next%>
			  </SELECT>
			  <INPUT name="vsoft2" class="input" size="10">

<br>������ҳ��<input name="vcome" class="input" size="35">
<br>������ҳ��<input name="vpage" class="input" size="35">
<br>������ͣ�<INPUT type="checkbox" name="outtype" value="Сʱ">Сʱ����
				<INPUT type="checkbox" name="outtype" value="��">�ա�����
				<INPUT type="checkbox" name="outtype" value="��">�ܡ�����
				<INPUT type="checkbox" name="outtype" value="��">��
				<br>����������<INPUT type="checkbox" name="outtype" value="��">�ꡡ����
				<INPUT type="checkbox" name="outtype" value="IP">IP������
				<INPUT type="checkbox" name="outtype" value="����">��������
				<INPUT type="checkbox" name="outtype" value="�����">�������
				<br>����������<INPUT type="checkbox" name="outtype" value="����ϵͳ">����ϵͳ
				<INPUT type="checkbox" name="outtype" value="��·">��·����
				<INPUT type="checkbox" name="outtype" value="ҳ��">ҳ�桡��
				<INPUT type="checkbox" name="outtype" value="��ϸ">��ϸ
</td></tr>
<tr><td align=right><a href="help.asp?id=001">����</a> <a href='javascript:document.forms[0].submit();'>�鿴���</a> <input type="submit" name="search" class="backc2" value=" "> &nbsp;</td></tr>
</form>
</table>

	</td>
    <td width="1" class="backs"></td>
  </tr>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<%
end if		'���Ȩ��

'��ʾԭ��������Զ������
'�����ݿ�
set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath(connpath)
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")

sql="select * from save"
rs.Open sql,conn,1,1

if not rs.EOF then
%>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
    <td width="498"class="backq">
		&nbsp; <img src="images/tb_title.gif" align=absmiddle><font style="font-size:16px">&nbsp;</font>�ˡˡ� �ѱ���ļ������� �ˡˡ�<br><br>

<table width="430" border="0" cellspacing=0 cellpadding=0 align=center>
  <tr height="1" class="backs"><td colspan=2><img src="images/touming.gif" width="1" height="1"></td></tr>
<%do while not rs.EOF%>
  <form action="tj_searchgo.asp" method=post>
  <tr height="28">
    <td>
      �� <font class=fonts><%=rs("name")%></font>
    </td>
    <td align=right>
      <input name="content" size="50" type="hidden" value="<%=rs("content")%>"
      ><input type="hidden" name="wherestr" value="<%=rs("wherestr")%>"
      ><input type="hidden" name="outtype" value="<%=rs("outtype")%>"
      ><input type="hidden" name="name" value="<%=rs("name")%>"
      ><%if session.Contents("master")=true or whatcan>=6 then%><a href="tj_save_del.asp?id=<%=rs("id")%>">ɾ��</a> <%end if%>�鿴��� <input type="submit" value=" " name="save" class="backc2"> &nbsp;
    </td>
  </tr>
  <%if rs("content") <> "" then%>
  <tr>
    <td colspan=2><p style='text-indent: 25;line-height: 130%;margin-left: 17;margin-right: 5;margin-top: 0;	margin-bottom: 5'><font color="#555555"><%=rs("content")%></font></td>
  </tr>
  <%end if%>
  <tr height="1" class="backs"><td colspan=2><img src="images/touming.gif" width="1" height="1"></td></tr>
  </form>
<%
rs.MoveNext
loop%>
</table>
<br>
	</td>
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
  </tr>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<%
end if

rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>
<!--#include file="inc_bottom.asp"-->
</body>
</html>