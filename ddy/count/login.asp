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
if Request("logpsw")=masterpsw then
	session.Contents("master")=true
	Response.Redirect "index.asp"
else
	Response.Redirect "help.asp?error=��¼�����������"
end if
%>