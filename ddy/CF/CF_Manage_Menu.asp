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
<table class="tb_3 td_1 tb_menu" style="margin:5px 0px 0px 0px;width:100%">
  <tr class="tr_1 td_1"> 
    <td >
	�û���<%=UserName%><br> <%
Set Rs= Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
Rs.Open Sql,Conn,1,1
%> 
<%If Rs("ParentName")<>"" And Session("CFCountSubAdmin")<>"" Then
 Response.write "<input type='submit' name='Submit' value='���ظ��˺�' onclick=""location.href='?Action=parentgoto'"">"
End If
If Session("CFCountUser_View")<>"" Then
 Response.write "�����鿴��"
End If
%>
</td>
  </tr>
  <tr> 
    <td ><img src=images/a.gif> <b><a href=?Action=tjsurvey>ͳ�Ƹſ�</a></b></td>
  </tr>
  <tr> 
    <td ><img src=images/closef.gif> <b><a href=?Action=lastvisit>���·���</a></b></td>
  </tr>
  <tr> 
    <td ><img src=images/closef.gif> <b><a href=?Action=lylist>��Դ��ϸ</a></b></td>
  </tr>
  <tr> 
    <td ><img src=images/closef.gif> <b><a href=?Action=onlinetj>����ͳ��</a></b></td>
  </tr>
  <tr> 
    <td ><img src=images/closef.gif> <b><a href=?Action=webtj>ҳ��ͳ��</a></b></td>
  </tr>
  <tr> 
    <td height=24 width=100%><img src=images/openf.gif id=aimgboard1  align="absmiddle"> 
      <b><a href="#">ͳ�Ʊ���</a></b></td>
  </tr>
  <tr> 
    <td style="display:yes;" id='submenuboard1'><table cellpadding=0 cellspacing=0 align=left >
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=hourtj">��Сʱͳ��</a></td>
        </tr>
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=daytj">������ͳ��</a></td>
        </tr>
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=monthtj">���·�ͳ��</a></td>
        </tr>

        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=backtj">��ͷ��ͳ��</a></td>
        </tr>
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=alexatj">Alext��ͳ��</a></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height=24 width=100%><img src=images/openf.gif id=aimgboard1  align="absmiddle"> 
      <b><a href="#">��������ͳ��</a></b></td>
  </tr>
  <tr> 
    <td style="display:yes;" id='submenuboard1'> <table cellpadding=0 cellspacing=0 align=left >
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=searchtj">����ͳ��</a></td>
        </tr>
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=keywordtj">�ؼ���ͳ��</a></td>
        </tr>
      </table></td>
  </tr>
<%
If Session("CFCountUser_View")="" Then'���Ƕ����鿴��ʱ
%>
<tr> 
    <td height=24 width=100%><img src=images/openf.gif id=aimgboard1  align="absmiddle"> 
      <b><a href="#">����������</a></b></td>
  </tr>
  <tr> 
    <td style="display:yes;" id='submenuboard1'> <table cellpadding=0 cellspacing=0 align=left >
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=jsqset">��������</a></td>
        </tr>
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=jsqquickset">������������</a></td>
        </tr>
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=stylelist">ͼƬ��ʽѡ��</a></td>
        </tr>
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=imglist">��������ʽѡ��</a></td>
        </tr>

        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=jsqreset">����������</a></td>
        </tr>
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=userdel">�˺�ɾ��</a></td>
        </tr>

      </table></td>
  </tr>
  <tr> 
    <td height=24 width=100%><img src=images/openf.gif id=aimgboard1  align="absmiddle"> 
      <b><a href="#">�޸ĸ�������</a></b></td>
  </tr>
  <tr> 
    <td style="display:yes;" id='submenuboard1'> <table cellpadding=0 cellspacing=0 align=left >
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href=?Action=pwdmodify>�޸�����</a></td>
        </tr>
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href=?Action=datamodify>�޸�����</a></td>
        </tr>
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href=?Action=passwordanswermodify>�޸����뱣��</a></td>
        </tr>
      </table></td>
  </tr>

  <tr> 
    <td ><img src=images/openf.gif> <b><a href=?Action=gbooklist>�����б�</a></b></td>
  </tr>
<%End If%>

<tr> 
    <td ><img src=images/openf.gif> <b><a href=?Action=getcode>���ͳ�ƴ���</a></b></td>
  </tr>
  <tr> 
    <td ><img src=images/exit.gif> <b><a href=?Action=userlogout>�˳�����</a></b></td>
  </tr>
</table>
