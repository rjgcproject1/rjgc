<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Const dbdns = "./"
Const Htmldns = "./"
Const Base_Target = "target=""_blank"""
%>

<!--#include file="AppCode/Conn.asp"-->
<!--#include file="AppCode/fun/function.asp"-->
<!--#include file="vbs.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!--<title>��ӭ�����й����ʴ�ѧ���人���о�����ҵ��Ϣ��</title>-->
<title><%=Application("Ok3w_SiteTitle")%></title>
<meta name="keywords" content="<%=Application("Ok3w_SiteKeyWords")%>">
<meta name="description" content="<%=Application("Ok3w_SiteDescription")%>">

<link rel="stylesheet" href="css/reset.css" />
<link rel="stylesheet" href="css/text.css" />
<link rel="stylesheet" href="css/960_24_col.css" />
<link rel="stylesheet" href="css/index.css" />

<script language="javascript" src="js/js.js"></script>
<SCRIPT src="js/sohuflash_1.js" type=text/javascript></SCRIPT>
<link href="images/fish.css" rel="stylesheet" type="text/css" />

<style type="text/css">
<!--

-->
</style>

<script type="text/javascript" src="js/jquery-1.4.2.min.js"></script>
<script type="text/javascript" src="js/embed_index.js" ></script>
<script type="text/javascript">
	$().ready(function(){
		onloadlst("index","info-list1",5,"","ȫְ","");//��λ��ʾ���б��ID��������ְλ���ְλ���ʣ���λ����
		onloadlst("index","info-list2",5,"","ʵϰ","");
	});
</script>
<link href="css/menu.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="js/menu.js" ></script>

</head>

<body>
<div class="container_24 all">

    	<div class="header">
            <div class="grid_24">
                <img src="img/banner.gif" align="banner" />
            </div>
            <div class="grid_24 nav">                
                    <%Call Ok3w_Article_Class_NavMenus()%>
            </div>
            <div class="clear"> </div>
		</div>
        <div class="grid_8 content">
        	<div class="left_top">
            	<!--<img src="img/left_top.gif" />-->
                <%Call Ok3w_Article_ImgFlash("",290,250)%>
            </div>
            <div class="left_middle">
            	<div class="menu">
                	<div class="text">
                    	��ҵ��Ϣ
                    </div>
                    <div class="more">
                    <a href="list.asp?id=32">More</a>
                    </div>
                </div>
                <div class="menudata width1">
                	<%Call Ok3w_Article_List(32,7,1,20,False,False,True,"5",False,"")%>
                </div>
                
            </div>
            <div class="left_bottom">
            	<div class="menu">
                	<div class="text">����ר��</div>
                    <div class="more">
                    <a href="list.asp?id=114">More</a>
                    </div>
                </div>
                <div class="menudata width1">
                	<%Call Ok3w_Article_List(114,7,1,20,False,False,True,"5",False,"")%>
                </div>
            </div>
        </div>
        <div class="grid_9 content">
        	<div class="center_top">
            	<div class="menu">
                	<div class="text">
                    	����֪ͨ
                    </div>
                    <div class="more">
                    <a href="list.asp?id=110">More</a>
                    </div>
                </div>
                <div class="menudata">
                	<%Call Ok3w_Article_List(110,7,1,30,False,False,True,"5",False,"")%>
                </div>
            </div>
            <div class="center_middle">
            	<div class="menu">
                	<div class="text">
                    	ѧ����Ϣ
                    </div>
                    <div class="more">
                    <a href="list.asp?id=126">More</a>
                    </div>
                </div>
                <div class="menudata">
                	<%Call Ok3w_Article_List(126,7,1,20,False,False,True,"5",False,"")%>
                </div>
            </div>
            <div class="center_bottom">
            	<div class="menu">
                	<div class="text">
                    	ѧ�ƽ���
                    </div>
                    <div class="more">
                    <a href="list.asp?id=6">More</a>
                    </div>
                </div>
                <div class="menudata">
                	<%Call Ok3w_Article_List(6,7,1,20,False,False,True,"5",False,"")%>
                </div>
            </div>
        </div>
        <div class="grid_7 content">
        	<div class="right_top">
            	<div class="menu">
                	<div class="text">
                    	��ѧͨ��
                    </div>
                    <div class="more">
                    <a href="list.asp?id=4">More</a>
                    </div>
                  
                </div>
                <div class="menudata2">
                	<p>
                    <span class="p">�����졿</span>
           			<marquee direction="up" height="207" scrollamount="2" scrolldelay="150" class="ggtitle1a" onmouseover="this.stop()" onmouseout="this.start()">
                    <%Call Ok3w_Article_List(51,4,1,40,False,False,False,"",False,"new")%>
                    </marquee><br>
            		<span class="q">�����졿</span> 
            		<marquee direction="up" height="207" scrollamount="2" scrolldelay="150" class="ggtitle1a" onmouseover="this.stop()"  onmouseout="this.start()">
                    <%Call Ok3w_Article_List(99,4,1,40,False,False,False,"",False,"new")%></marquee><br>
          			</p>
                </div>
            </div>
            <div class="right_bottom">
            	<div class="menu">
                	<div class="text">
                    	��������
                    </div>
                    <div class="more">
                    <img src="img/link.png" />
                    </div>
                </div>
                <div class="menudata">
					<select onchange=javascript:window.open(this.options[this.selectedIndex].value) name=D2>
                    <option value="#"> ==��ҵ����== </option>
                    <option value="http://www.mapgis.com.cn/">�е�����</option>
                    </select>
                    
                    <select onchange=javascript:window.open(this.options[this.selectedIndex].value) name=D1>
                    <option value="#">   ==У����վ== </option>
                    <option value="http://www.cug.edu.cn">�й����ʴ�ѧ���人��</option>
                    <option value="http://graduate.cug.edu.cn/">�ش��о���Ժ</option>
                    <option value="http://xgxy.cug.edu.cn/ygb/">��Ϣ����ѧԺ</option>                    
                    </select>
                    
                    <select onchange=javascript:window.open(this.options[this.selectedIndex].value) name=D1>
                    <option value="#">   ==�ֵ�ԺУ== </option>
                    <option value="http://www.cugb.edu.cn/">�й����ʴ�ѧ��������</option>
                    <option value="http://sse.hust.edu.cn/">���пƼ���ѧ���ѧԺ</option>
                    <option value="http://iss.whu.edu.cn/">�人��ѧ�������ѧԺ</option>                               
                    </select>
                </div>
            </div>
        </div>
        <div class="grid_24 footer">
           <div class="right">
				��Ȩ���У��й����ʴ�ѧ���人����Ϣ����ѧԺ�������ϵ
    		</div>
    	</div>
</div>
</body>
</html>
