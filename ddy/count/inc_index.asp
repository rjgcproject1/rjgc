<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"></td>
    <td width="498"class="backq">
		&nbsp; <img src="images/tb_index.gif" align=absmiddle> &nbsp;�ˡˡ� <%=countname%> �ˡˡ�
		<p style="line-height: 140%; margin-left: 25; margin-right: 10; margin-top: 10; margin-bottom: 5" align=left>
		<a href="main.asp">�� ��������</a>
		<a href="tj_hour.asp">�� 24Сʱ</a>
		<a href="tj_day.asp">�� ��ͳ��</a>
		<a href="tj_month.asp">�� ��/��ͳ��&nbsp;</a>
		<a href="tj_week.asp">�� ����ͳ��</a>
		<a href="tj_come.asp">�� ��·</a>
		<a href="tj_page.asp">�� ҳ��</a>
		<br>
		<a href="tj_ip.asp">�� ������IP</a>
		<a href="tj_soft.asp">�� ������Ͳ���ϵͳ</a>
		<a href="tj_where.asp">�� �����ߵ���</a>
		<%if session.Contents("master")=true or whatcan>2 then%><a href="tj_all.asp">�� ��ϸ��¼</a><%end if%>
		<%if session.Contents("master")=true or whatcan>3 then%><a href="tj_search.asp">�� �Զ���ͳ��</a><%end if%>
	</td>
    <td width="1" class="backs"></td>
  </tr>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
