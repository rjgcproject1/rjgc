<%@ CODEPAGE = "936" %>
<!--#include file="inc_config.asp"-->
<%
'���������ȡ����
id=Request("id")
error=Request("error")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Copyright" content="Ajiang http://www.jhqing.com">
<title><%=countname%>��������Ϣ</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<!--#include file="inc_top.asp"-->
<br>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
    <td width="498"class="backq">
		&nbsp; <img src="images/tb_help.gif" align=absmiddle> <font style="font-size:16px">&nbsp;</font>�� �� �� Ϣ<br>
	<%if error <> "" then Response.Write "<p class='p1'><font color='red'><img src='images/tb_error.gif' align=absmiddle><font style='font-size:15px'>&nbsp;</font>���� " & error & "</font>"%>
<%select case id
case ""%>
	<li style="	margin-left: 50; margin-top: 15;margin-bottom: 5"><a href="help.asp?id=001">�Զ���ͳ������Ӧ��������д�أ�</a>
	<li style="	margin-left: 50; margin-top: 5;	margin-bottom: 5"><a href="help.asp?id=002">��α����Զ�����������ļ���������</a>
	<li style="	margin-left: 50; margin-top: 5;	margin-bottom: 5"><a href="help.asp?id=003">����޸Ļ�ɾ���Ѿ�����ļ���������</a>
	<li style="	margin-left: 50; margin-top: 5;	margin-bottom: 5"><a href="help.asp?id=004">�ÿͺ͹���Ա�ֱ������ЩȨ�ޣ�</a>
	<li style="	margin-left: 50; margin-top: 5;	margin-bottom: 5"><a href="help.asp?id=005">��λ�ñ�ϵͳ�����°汾��</a>
	<li style="	margin-left: 50; margin-top: 5;	margin-bottom: 5"><a href="help.asp?id=006">��ϵͳ�İ�Ȩ������</a>
	
<%case "001"%>
	<p class="p1"><font class="fonts"><b>�Զ���ͳ������Ӧ��������д�أ�</b></font>
	<p class="p1">��ֹ���ڣ��ꡢ�¡��ն�������д�����Ҷ����������֡����ǿ���ֻ����ʼ���ڣ�������ʾ����д����ʼ�������������ʱ��Σ�ͬ��Ҳ����ֻ��д��ֹ���ڣ�������ʾ�ӿ�ʼͳ������������д��������ڣ�ע���ֹ���ڱ����ڿ�ͨͳ������֮��
	<p class="p1">IP��ַ�����Ե������������֧��ģ����ѯ�ģ�������IP��ַ�������롰61.������ɲ鵽IP��ַ�ǡ�61.163.233.10�����ߡ�202.61.33.25�������ķ��ʼ�¼��ͳ�Ʒ�����
	<p class="p1">����ϵͳ�����������������Դ��б���ѡ��Ҳ����ֱ���ں���ı��������룬Ҳ֧��ģ����ѯ���������ı���������win������Բ鵽WIN 2K��WIN XP��WIN 9X��ȫ����¼��ͳ�Ʒ�����
	<p class="p1">���Ժͱ�����ҳ��������ѡ����������鵽�ض���·��Ŀ��ķ��ʼ�¼��ͳ�Ʒ�����֧��ģ����ѯ��
	<p class="p1">������ͣ�����ѡ��Ҫ�鿴�ķ��������ͣ��ɶ�ѡ����Ҫע����ѡ���Խ�࣬��Ҫ�ȴ���ʱ���Խ������Ȼ����Ҳ����һ������ѡ��

<%case "002"%>
	<p class="p1"><font class="fonts"><b>��α����Զ�����������ļ���������</b></font>
	<p class="p1">�ڽ������Զ���ͳ�Ƶļ���֮���ڸ�ҳ��ײ���һ����Ϊ�����汾�μ�������������Ŀ�򣬸ÿ��е�ѡ��������ΪҪ���������ȡ���������顣
	<p class="p1">����ǰ����Ϊ������ȡһ�����֣������պ��޷�ʹ��������ļ���������
	<p class="p1">�����ΪҪ����ļ�������дһ����飬��Ϊ����д��̫����Ӱ��ҳ�����۵ģ�Ҫϸ�µÿ����������������˼��д������Ǳ�Ҫ�ġ�
	<p class="p1">�����ֵ�����������������ѡ�򣬿�������ָ���������еļ�����Ŀ�����Ǹ��ǵ�ԭ�е��ػ�����ϵͳ��ʾ����ȡ���֣���������������Ϊ���޸�������Ŀ��ȷ������ûд��Ļ�������һ��������ѡ��ͬ��ʱ���ʡ�����ȷ���������ݵİ�ȫ��
	<p class="p1">ֻ�й���Ա�ſ��Խ�������������
	
<%case "003"%>
	<p class="p1"><font class="fonts"><b>����޸Ļ�ɾ���Ѿ�����ļ���������</b></font>
	<p class="p1">���޸ģ���ֻ��Ҫ�����µļ����������м������ڼ������ҳ��ĵײ��ı�������ʹ��ԭ�е����ֱ��棬��ѡ��ͬ��ʱ���ǡ����������޸���ԭ���������Ŀ��
	<p class="p1">��ɾ�����ڡ��Զ���ͳ�ơ�ҳ����Զ�����������б��е����ɾ�������ӣ�����ɾ��ҳ�棬ϵͳ��ʾ�Ƿ����Ҫɾ�������ȷ�����ɡ�
	<p class="p1">ֻ�й���Ա�ſ��Խ�������������

<%case "004"%>
	<p class="p1"><font class="fonts"><b>�ÿͺ͹���Ա�ֱ������ЩȨ�ޣ�</b></font>
	<p class="p1">�ෲ��վ����ͳ�ƶԷ���Ȩ��ʵ�еȼ���������Աʼ�վ��� 6 ����Ȩ������ӵ������Ȩ�����ǹ���Ա���ÿͣ���ӵ�е�Ȩ���ɹ���Ա�ڰ�װʱ�������ļ������á��ȼ����ֵİ취�ǣ�</p>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="320" align=center>
	<tr height="18" align=center class=backzq><td width="110">�� ��</td><td width="30">0</td><td width="30">1</td><td width="30">2</td><td width="30">3</td><td width="30">4</td><td width="30">5</td><td width="30">6</td></tr>
	<tr height="18" align=center><td align=left>&nbsp;�鿴��������</td><td></td><td>��</td><td>��</td><td>��</td><td>��</td><td>��</td><td>��</td></tr> 
	<tr height="18" align=center class=backzq><td align=left>&nbsp;�鿴ȱʡͳ������</td><td></td><td></td><td>��</td><td>��</td><td>��</td><td>��</td><td>��</td></tr> 
	<tr height="18" align=center><td align=left>&nbsp;�鿴��ϸ��¼</td><td></td></td><td></td><td><td>��</td><td>��</td><td>��</td><td>��</td></tr> 
	<tr height="18" align=center class=backzq><td align=left>&nbsp;�鿴�Զ����¼</td><td></td><td></td><td></td><td></td><td>��</td><td>��</td><td>��</td></tr> 
	<tr height="18" align=center><td align=left>&nbsp;�Զ�������ͳ��</td><td></td><td></td><td></td><td></td><td></td><td>��</td><td>��</td></tr> 
	<tr height="18" align=center class=backzq><td align=left>&nbsp;�����Զ����</td><td></td><td></td><td></td><td></td><td></td><td></td><td>��</td></tr> 
</table>
	<p class="p1">����ÿ͵ȼ�������Ϊ 0����ÿͽ�û���κ�Ȩ��������ÿ͵ȼ�������Ϊ 6����ÿͽ����к͹���Ա��ȵ�Ȩ����Ĭ�ϵķÿ�Ȩ�޵ȼ�Ϊ 4��
	<p class="p1">��ǰ�ÿ͵�Ȩ�޵ȼ�Ϊ��<%=whatcan%>

<%case "005"%>
	<p class="p1"><font class="fonts"><b>��λ�ñ�ϵͳ�����°汾��</b></font>
	<p class="p1">������ע�ҵ���ҳ <a href="http://www.jhqing.com" target="_blank">http://www.jhqing.com</a> �һ����Ƴ��°�ĵ�һʱ������ҳ�Ϸ�������Ҳ�����ෲ����ҳ�ϻ���ෲ��д����������
	<p class="p1">�����ע�Ᵽ���ෲ�İ�Ȩ��Ϣ���һ��Ѻܶྫ���������ֳ����������ʹ�ã�����һ�ֳɾ͸У��Ҳ��������ȡ������ҫ��
	<p class="p1">����л��ʹ���ෲ��д�ĳ���ϣ�����Եõ����Ľ���ͼ���֧�֡�

<%case "006"%>
	<p class="p1"><font class="fonts"><b>��ϵͳ�İ�Ȩ��Ϣ</b></font>
	<p class="p1">��ϵͳ��Ȩ�����ෲ(zhqi123@163.com)���У��������������ģ�
	<br>����1���ڸ��˻���ҵ��վ��ʹ�ñ����򲢱����ෲ�İ�Ȩ��Ϣ��
	<br>����2���ڱ���ԭѹ���������Ե���������·ַ�������
	<br>����3���޸ı���������޸���С������֮һ����ע�����޸����ෲ����վͳ��ϵͳ�������������ෲ���估��վ���ӣ�
	<br>����4���޸ı����򳬹�����֮һ�ģ�����ɾ���й��ෲ���κ���Ϣ�����·ַ������ڷַ�ǰӦ֪ͨ�ෲ��
	<br>����5�������ĳ��������ñ������д���Ƭ�ϣ�����ע�����������ෲ����վͳ��ϵͳ����
	<p class="p1">��������ǲ�����ģ�
	<br>����1���޸ĳ���������֮һ���޸İ�Ȩ���·ַ���
	<br>����2����������վ��ʹ�ñ�ϵͳȫû�б����ෲ��Ȩ��Ϣ�����ӣ�
	<br>����3��ֱ�ӻ���ʹ�ñ�����׬ȡ����
	<p class="p1">������������û�ʹ�ñ����򣬵���������İ�Ȩ�󴫲������������������޸��ҵ�̽����򣨽����޸��˰�Ȩ�������Լ�����ҳ�Ϸ�����˵���Լ������ģ���������������ෲ�е��ǳ��ź���ϣ���������鲻Ҫ�ٷ������ҵ�ͳ��ϵͳ�ϡ�
	<p class="p1">����Ҳ�����Խ�������������ҵ��վ�����κ��˶����ܴ��л�����������վ���蹫˾����ϵͳ���ڸ��ͻ���������վ��ʱ������Ϊ���������ͻ����ӷ��ã�Ҳ���ý���������͡���ϵͳ��Ϊ�Ϳͻ�̸�е�һ��������
	<p class="p1">����л��ʹ���ෲ��д�ĳ���ϣ�����Եõ����Ľ���ͼ���֧�֡�
	
<%case else%>
<p class="p1">Ŀǰ��û�������صİ�����Ϣ��

<%end select%>
<p class="p1" align="right"><a href='javascript:history.back()'>����</a> <a href='javascript:history.back()'><img src="images/arbutton.gif" align="absmiddle" border="0"></a> <font style="font-size:16px">&nbsp;</font>&nbsp;
	</td>
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
  </tr>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<!--#include file="inc_bottom.asp"-->
</body>
</html>
