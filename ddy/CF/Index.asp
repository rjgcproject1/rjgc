<%
'�˷���û�������
'����QQ��178575
'����EMail��yliangcf@163.com
'������վ��http://www.qqcf.com
'��ϸ��飺http://www.qqcf.com/?action=list&list=cfcount
'�����г���������ʾ����װ��ʾ��ʹ�����ѽ�����°汾���ص�����
'��Ϊ��Щ���ݿ���ʱ�����£���û�з��ڳ�������Լ�����վ�ϲ鿴
'�������汾����ʾ
%>
<!--#Include File="Conn.asp"-->
<!--#Include File="CF_MyFunCtion.asp"-->
<!--#Include File="CF_Md5.asp"-->
<!--#Include File="CF_Getcookie.asp"-->
<%
Action=Request("Action")
%>
<!--#Include File="CF_Do.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HTML><HEAD><TITLE><%=RsSet("SysTitle")%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<!--#Include File="CF_Style.asp"-->
</HEAD>

<BODY>
<!--#Include File="Top.asp"-->

<table class="tb_2 tb_color" style="margin-top:4px;margin-bottom:0px">
  <tr>
    <td>&nbsp;&nbsp;&nbsp;<a href="Manage.asp?Action=lylist">��Դ��ϸ</a>|&nbsp;&nbsp;<a href="Manage.asp?Action=lytj">��Դͳ��</a>|&nbsp;&nbsp;<a href="Manage.asp?Action=onlinetj">����ͳ��</a>|&nbsp;&nbsp;<a href="Manage.asp?Action=daytj">ÿ��ͳ��</a>|&nbsp;&nbsp;<a href="Manage.asp?Action=hourtj">Сʱͳ��</a>|&nbsp;&nbsp;<a href="Manage.asp?Action=backtj">��ͷ��ͳ��</a>|&nbsp;&nbsp;<a href="Manage.asp?Action=searchtj">��������ͳ��</a>|&nbsp;&nbsp;<a href="Manage.asp?Action=keywordtj">�����ؼ���ͳ��</a>|&nbsp;&nbsp;<a href="Manage.asp?Action=jsqset">��������</a>
	</td>
  </tr>
</table>

<table align="center">
  <tr> 
    <td valign="top"><!--#Include File="CF_UserLogin.asp"--></td>
    <td><IMG src="images/banner.jpg" hspace="3" vspace="2" border="0"></td>
  </tr>
</table>

<TABLE class="tb_2" style="margin-top:2px;margin-bottom:0px;">
    <TR  class="tr_1"> 
      <TD  class="td_1"><%=RsSet("SysTitle")%>���</TD>
    </TR>

    <TR> 
      <TD>

<TABLE width=100% border=0 align=center cellPadding=0 cellSpacing=0 
bgColor=#ffffff class="te2">
                    <TBODY>
                      <TR>
                        <TD vAlign=bottom align=middle width=10 height=218>&nbsp; </TD>
                        <TD width=706 >�����ܣ� 
                          <br>
                          1.�ܹ�60�ּ�����ͼƬ��ʽ����ѡ�񣬻�������ͳ��ͼ�����ѡ����ȫ�����������<br>
                          2.�������ü�������ʾ���֣���ʾλ�����������Ƿ����أ�ͳ����Ϣ�Ƿ񹫿��ȡ�<br>
                          3.ҳ����ʾ������IP��ˢ�¼������ּ���ģʽ��֧��Script��վ��ʽ��Img�����෽ʽ���ü��������롣<br>
                          4.���Լ�¼���ÿ͵���ԴIP��ַ����Դҳ����Ϣ������������<br>
                          5.ÿ�¡�ÿ���ÿСʱ�ķ�������ͳ�ƣ���ͷ��ͳ�ƣ�ÿ����ҳ���ͳ�ơ�<br>
                          6.��������ͳ�ƣ��������Լ������������棬�����ؼ���ͳ�ơ�<br>
                          7.ע���û��һ����빦�ܣ��û���������ͳ�ƣ�ɾ��ע���˺š�<br>
                          8.��ȫ�ԣ�����MD5���ܣ��һ������MD5���ܣ�ע�ᡢ��½ʹ����֤�룬��ȫ��Sqlע�롣<br> ���еĹ��ܣ� 
                          <br>
                          1.����ʹ�û��漼���������ں�ʮ��ǿ����<br>
                          2.����������ͼƬ��ͳ��ͼ�����ֻ��ƹ��棬�ڶ����ÿɵ���<br>
                          3.Script��վ��Img���������ַ�ʽ���ü�������Img�����෽ʽ�������������κ��ܲ���ͼƬ�ĵط�ʹ�á�<br>
                          4.���еĴ����Զ��޸����ƣ����ڼ���������������Զ��޸���<br>
                          5.��ȫ�ž������߳����׶����ݿ���ɵ��𻵣������������վ��ʹ�ñ��ֺ��ȶ���<br>
                          6.������ƣ��ڻ����б������ݣ����������������������ٶ����ݿ�����ӣ�ɾ��Ƶ���Ĳ�����<br>
                          7.�ȶ��ԡ���ȫ�ԡ��ٶ��ϱ��ֶ������㣬������ȫ�����뼯�ɳ̶ȸߡ���ȫ������רҵ��������ȫ��ѡ�<br> 
                        </TD>
                        <TD vAlign=bottom align=middle width=21 height=218>&nbsp;</TD>
                      </TR>
                    </TBODY>
                  </TABLE>
				  
      </TD>
    </TR>
</TABLE>



<!--#Include File="Bottom.asp"-->
</BODY>
</HTML>

<%Call ConnClose()%>