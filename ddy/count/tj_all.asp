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
if session.Contents("master")=false and whatcan<3 then Response.Redirect "help.asp?id=004&error=��û�в鿴��ϸ��¼��Ȩ�ޡ�"

wherestr=Request("wherestr")

pagesize=mPageSize
if request("page")="" then
  	curpage = 1
else
	curpage = cint(request("page"))
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Copyright" content="Ajiang http://www.jhqing.com">
<title><%=countname%>����ϸ��¼</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<!--#include file="inc_top.asp"-->
<br>
<table width=500 cellspacing=0 align=center>
  <tr><td>
    <p style="line-height: 100%; margin-left: 15; margin-top: 5; margin-bottom: 0">
    Tips: ������ָIP��ַ�ɿ������ڵ�������ָ�������ڿɲ鿴������ҳ�档</p>
  </td></tr>
</table>
<br>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"></td>
    <td width="498"class="backq">
		&nbsp; <img src="images/tb_title.gif" align=absmiddle> &nbsp;�ˡˡ� �� ϸ �� ¼ �ˡˡ�<br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#5EC707" width="480" align=center>
  <tr align=center height="25">
    <td width="115">ʱ &nbsp; ��</td><td width="85">IP</td><td width="55">����ϵͳ</td><td width="55">�����</td><td width="170">�� Դ �� ҳ</td>
  </tr>
<%
set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath(connpath)
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")

sql = "SELECT * FROM view " & wherestr & " ORDER BY id DESC"
rs.open sql, conn, 1, 1
if rs.bof and rs.eof then
	rs.close
	response.write "<tr><td colspan='5'><br><center>" & wherestr & "û�з��������ļ�¼</center><br></td></tr></table></td></tr></table>"
else
	dim i
	rs.pagesize = pagesize
	if rs.pagecount < curpage then
		rs.absolutepage=rs.pagecount
		curpage=rs.pagecount
	else
		rs.absolutepage = curpage
	end if
for i = 1 to rs.pagesize
%>
  <tr height="1" bgcolor="#5EC707"><td colspan=5></td></tr>
  <tr height="18">
	<td><a title="<%=rs("vpage")%>"><%=rs("vtime")%></a></td>
	<td><font face="Times new roman"><a title="<%=rs("vwhere") & "��" & rs("vwheref")%>"><%=rs("vip")%></a></font></td>
	<td><font face="Times new roman"><%=rs("vOS")%></font></td>
	<td><font face="Times new roman"><%=rs("vsoft")%></font></td>
	<td><%
	vcome=rs("vcome")
	thelen=len(vcome)
	if thelen=0 then Response.Write ""
	if thelen <= 33 and thelen > 0 then
		svcome=right(vcome,thelen-6)
		Response.Write "<a href='" & vcome & "' target='_blank'>" & svcome & "</a>"
	end if
	if thelen >= 34 then
		svcome=left(right(vcome,thelen-6),24) & "..."
		Response.Write "<a title='" & vcome & "' href='" & vcome & "' target='_blank'>" & svcome & "</a>"
	end if
	%></td>
  </tr>
<%
  rs.movenext
    if rs.eof then
      i = i + 1
      exit for
    end if
  next
%>
</table>
	</td>
    <td width="1" class="backs"></td>
  </tr>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>

<br>
<table width=500 cellspacing=0 align=center cellpadding=0>
<form method=post action="tj_all.asp" id=form2 name=form2>
  <tr><td colspan=3><img src=images/photoup.gif></td></tr>
  <tr height=65>
    <td width=1 class=backs></td>
    <td class=backq width="498">
		&nbsp; <img src="images/gb-search.gif" align=absmiddle> �ˡˡ� �� ҳ �ˡˡ�<br><br>
     &nbsp; ��<font class="fonts"><%=cstr(curpage)%></font>ҳ ��<font class="fonts"><%=cstr(rs.pagecount)%></font>ҳ ��ҳ<font class="fonts"><%=cstr(i-1)%></font>�� ��<font class="fonts"><%=cstr(rs.recordcount)%></font>�� &nbsp;
      <%
      wherestr=server.URLEncode(wherestr)
      if curpage <> 1 then%>
		<a href="tj_all.asp?page=1&wherestr=<%=wherestr%>">&lt;&lt;��ҳ </a>
		<a href="tj_all.asp?page=<%=cstr(curpage-1)%>&wherestr=<%=wherestr%>">&lt;��һҳ </a>
      <%else%>
        <font color="#888888">&lt;&lt;��ҳ &lt;��һҳ</font>
      <%end if
      if curpage <> rs.pagecount then%>
		<a href="tj_all.asp?page=<%=cstr(curpage+1)%>&wherestr=<%=wherestr%>"> ��һҳ&gt;</a>
		<a href="tj_all.asp?page=<%=cstr(rs.pagecount)%>&wherestr=<%=wherestr%>"> βҳ&gt;&gt;</a>
      <%else%>
         <font color="#888888">��һҳ&gt; βҳ&gt;&gt;</font>
      <%end if%>
		<input type=hidden name=wherestr value="<%=wherestr%>">&nbsp;<input type=text name='page' size=9 class=input> <input type=submit value=' ' class="backc2">
	</td>
    <td width=1 class=backs></td>
  </tr>
  <tr><td colspan=3><img src=images/photodown.gif></td></tr>
</form>
</table>
<br>
<%
'����Ǽ������������ʾ�������ݱ�
if wherestr <> "" then
wherestr=Request("wherestr")
'�������ݱ�
if session.Contents("master")=true or whatcan>=6 then

%>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <form action="tj_save.asp" method=post id=form1 name=form1>
  <tr height="30">
    <td width="1" class="backs"></td>
    <td width="498"class="backq">
		&nbsp; <img src="images/tb_save.gif" align=absmiddle><font style="font-size:16px">&nbsp;</font>�ˡˡ� ������μ������� �ˡˡ�<br>

	<p class="p1"><font class="fonts"><b>��������</b></font>&nbsp; <%if wherestr="" then%>û�м�������<%else%><%=wherestr%><%end if%>��
	<p class="p1" style='margin-bottom: 0;margin-top: 0;'><font class="fonts"><b>��ѯ��Ŀ</b></font>&nbsp; ��ϸ��
	<p class="p1" style='line-height: 100%;margin-bottom: 0;margin-top: 0;'><font class="fonts"><b>ȡ������</b></font>&nbsp; <input name="name" size="16" class="input">&nbsp;
	<INPUT type="radio" name="overwrite" value="0" checked> ͬ��ʱ��ʾ&nbsp;
	<INPUT type="radio" name="overwrite" value="1"> ͬ��ʱ����
	<p class="p1" style='line-height: 100%;margin-bottom: 0;margin-top: 0;'><font class="fonts"><b>�Ӹ�����</b></font>&nbsp; <input name="content" size="50" class="input">
	<p class="p1" style='margin-top: 0;' align="right">
	<a href="help.asp?id=002">����</a>
	<a href='javascript:document.forms[0].submit();'>����</a>
	<input type="hidden" name="wherestr" value="<%=wherestr%>"><input type="hidden" name="outtype" value="��ϸ"><input type="submit" value=" " name="save" class="backc2"><font style="font-size:16px">&nbsp;</font>&nbsp;

	</td>
    <td width="1" class="backs"></td>
  </tr>
  </form>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<%
end if
end if
%>

<%
rs.close
end if

set rs=nothing
conn.Close 
set conn=nothing
%>
<!--#include file="inc_bottom.asp"-->
</body>
</html>
