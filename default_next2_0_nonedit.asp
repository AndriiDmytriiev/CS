<%
on error resume next
set conn2 = server.CreateObject("Adodb.connection")
      
        connStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & _
         Server.MapPath("missyou.mdb")
         conn2.Open connStr
      ' Recordset object
      
        
        
      
'-��������� ������ �� ������ �����---------
Session("b1")=request.querystring("b1")
Session("b2")=request.querystring("b2")
Session("b3")=request.querystring("k1")
Session("b4")=request.querystring("b3")
Session("b5")=request.querystring("b4")
Session("b6")=request.querystring("k2")
Session("b7")=request.querystring("b5")
Session("b8")=request.querystring("b6")
Session("b9")=request.querystring("k3")
Session("b10")=request.querystring("b7")
Session("b11")=request.querystring("b8")
Session("b12")=request.querystring("k4")
Session("b13")=request.querystring("b9")
Session("b14")=request.querystring("b10")
Session("b15")=request.querystring("k5")
Session("b16")=request.querystring("b11")
Session("b17")=request.querystring("b12")
Session("b18")=request.querystring("k6")
Session("b19")=request.querystring("b13")
Session("b20")=request.querystring("b14")
Session("b21")=request.querystring("b15")
Session("b22")=request.querystring("b16")
Session("b23")=request.querystring("b17")
Session("b24")=request.querystring("b18")
Session("b25")=request.querystring("k7")
Session("b26")=request.querystring("b19")
Session("b27")=request.querystring("k8")
Session("b28")=request.querystring("b20")
Session("b29")=request.querystring("b21")
Session("b30")=request.querystring("k9")
Session("b31")=request.querystring("b22")
Session("b32")=request.querystring("b23")
Session("b33")=request.querystring("k10")
Session("b34")=request.querystring("b24")
Session("b35")=request.querystring("b25")
Session("b36")=request.querystring("k11")
Session("b37")=request.querystring("b26")
Session("b38")=request.querystring("b27")
Session("b39")=request.querystring("b28")
Session("b40")=request.querystring("b29")
Session("b41")=request.querystring("b30")
Session("b42")=request.querystring("b31")
Session("b43")=request.querystring("b32")
Session("b44")=request.querystring("b33")

'------------------------------------------


        
%>
<HTML>
  <HEAD>
  
    <title>Default page</title>
     <link rel="stylesheet" type="text/css" href="index.css">
<META content="text/html; charset=windows-1251" http-equiv=Content-Type> 
<META NAME="Keywords" CONTENT="">

  </HEAD>
  <BODY background="bg.jpg" alink="#0000ff" vlink="#0000ff" link="#0000ff">


        
<center>
<img src="relationshipsromance.jpeg"></img>
</center>


<br>
<br>
<br>
<%if Session("login")<>"" then%>
<%
'Session("day")=Request.Form("day")
'        Session("month")=Request.Form("month")
'        Session("year")=Request.Form("year")
%>


<table>
<tr>
<td>���������:<%=Session("Login")%></td><td>|</td><td>������:<%=Session("client")%></td>
</tr>
<tr>
<td>����:<%=date()%></td><td>|</td><td>���������� ����, ���:<%=Session("contact")%></td>
</tr>
<tr>
<td>���� ������:<%=Session("day") & "." & Session("month") & "." & Session("year") %></td><td>|</td><td>���� ������ �����:<%=date()%></td>
</tr>
</table>

<%end if%>

<td>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

</td>
<td>
<table>
<tr>
<td align=right valign=top>




<form method=get >
<!--------------
<table border=0 width=100%>
<tr><td align=left>��������: ���������� �������� ������, ���������� ������. ������, ���������� � ���� ������������ ������ ���������������!
</td></tr>
<tr><td align=left>1.<b><u>���������� ������ ���������:</u></b></u></b></td></tr>
</table>
<br>


������������, ����������, ���������� �� ���� �������, � ������� ���������� ����������� ������.



<table border=1>
<tr>
<td>&nbsp;</td><td>����������</td><td>��.���.</td><td>������ � 1</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>1.</td><td><b>������� ������ � ������/�������</b></td><td></td><td>�����</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>1.1</td><td>���������</td><td>��.</td><td>
<input type=textbox name=t1 size="20"></td><td>
<input type=checkbox name = ch1 ></td><td>
<select name="Period1">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>1.2</td><td>��� ���������� �������</td><td>���</td><td>
<input type=textbox name=t2 size="20"></td><td>
<input type=checkbox name = ch2 ></td><td>
<select name="Period2">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>1.3</td><td>������ ��������</td><td>�.</td><td>
<input type=textbox name=t3 size="20"></td><td>
<input type=checkbox name = ch3 ></td><td>
<select name="Period3">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>2.</td><td><b>����� �������</b></td><td>��.�</td><td>
<input type=textbox name=t4 size="20"></td><td>
<input type=checkbox name = ch4 ></td><td>
<select name="Period4">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>2.1</td><td>������� ������� �����</td><td>��.�</td><td>
<input type=textbox name=t5 size="20"></td><td>
<input type=checkbox name = ch5 ></td><td>
<select name="Period5">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>2.2</td><td>�������� VIP</td><td>���-��</td><td>
<input type=textbox name=t6_1 size="20"></td><td>
<input type=checkbox name = ch6 ></td><td>
<select name="Period6">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t6_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.3</td><td>������� ���������</td><td>���-��</td><td>
<input type=textbox name=t7_1 size="20"></td><td>
<input type=checkbox name = ch7 ></td><td>
<select name="Period7">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t7_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.4</td><td>��������� ���������</td><td>���-��</td><td>
<input type=textbox name=t8_1 size="20"></td><td>
<input type=checkbox name = ch8 ></td><td>
<select name="Period8">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t8_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.5</td><td>��������� ���������</td><td>���-��</td><td>
<input type=textbox name=t9_1 size="20"></td><td>
<input type=checkbox name = ch9 ></td><td>
<select name="Period9">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t9_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.6</td><td>����������� ���������, �������</td><td>���-��</td><td>
<input type=textbox name=t10_1 size="20"></td><td>
<input type=checkbox name = ch10 ></td><td>
<select name="Period10">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t10_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>


<tr>
<td>2.7</td><td>��������</td><td>���-��</td><td>
<input type=textbox name=t11_1 size="20"></td><td>
<input type=checkbox name = ch11 ></td><td>
<select name="Period11">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t11_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.8</td><td>��������</td><td>���-��</td><td>
<input type=textbox name=t12_1 size="20"></td><td>
<input type=checkbox name = ch12 ></td><td>
<select name="Period12">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t12_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.9</td><td>�����</td><td>���-��</td><td>
<input type=textbox name=t13_1 size="20"></td><td>
<input type=checkbox name = ch13 ></td><td>
<select name="Period13">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t13_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.10</td><td>����������</td><td>���-��</td><td>
<input type=textbox name=t14_1 size="20"></td><td>
<input type=checkbox name = ch14 ></td><td>
<select name="Period14">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t14_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.11</td><td>�������</td><td>���-��</td><td>
<input type=textbox name=t15_1 size="20"></td><td>
<input type=checkbox name = ch15 ></td><td>
<select name="Period15">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t15_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.12</td><td>�����, ���������</td><td>���-��</td><td>
<input type=textbox name=t16_1 size="20"></td><td>
<input type=checkbox name = ch16 ></td><td>
<select name="Period16">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t16_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.13</td><td>������ ������� (�� ����.������������)</td><td></td><td>
<input type=textbox name=t17_0 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>���-��</td><td>
<input type=textbox name=t17_1 size="20"></td><td>
<input type=checkbox name = ch17 ></td><td>
<select name="Period17">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t17_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>


<tr>
<td>3.</td><td><b>�-�� ����������� ����� �������� � �����������, ���. � ������� � ����</b></td><td>���.</td><td>
<input type=textbox name=t18 size="20"></td><td>
<input type=checkbox name = ch18 ></td><td>&nbsp;</td>
</tr>


<tr>
<td>4.</td><td><b>������������� � ����������� �/� ���������� ����������� (��������� ������ � �����)</b></td><td>���-��</td><td>&nbsp;</td><td>&nbsp;</td><td>
<select name="Period19">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>��������� ������</td><td>���./���.</td><td>
<input type=textbox name=t19_1 size="20"></td><td>
<input type=checkbox name = ch19_1 ></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>������ ����</td><td>����/���.</td><td>
<input type=textbox name=t19_2 size="20"></td><td>
<input type=checkbox name = ch19_2 ></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>��� ���������</td><td>����/���.</td><td>
<input type=textbox name=t19_3 size="20"></td><td>
<input type=checkbox name = ch19_3 ></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>�������� ������� �/�������</td><td>��./���.</td><td>
<input type=textbox name=t19_4 size="20"></td><td>
<input type=checkbox name = ch19_4 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.</td><td><b>�����������</b></td><td>���.</td><td>&nbsp;</td><td>
<input type=checkbox name = ch20 ></td><td>
<select name="Period20">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>

<tr>
<td>5.1</td><td>������ �������� (��������)</td><td>��.�</td><td>
<input type=textbox name=t20_1 size="20"></td><td>
<input type=checkbox name = ch20_1 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.2</td><td>����������� �������� (��������, ������, ���������, �������� ���, �������)</td><td>��.�</td><td>
<input type=textbox name=t20_2 size="20"></td><td>
<input type=checkbox name = ch20_2 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.3</td><td>������� �������� (������, ������, ������) ���/�����</td><td>��.�</td><td>
<input type=textbox name=t20_3 size="20"></td><td>
<input type=checkbox name = ch20_3 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.4</td><td>���������� �����������</td><td>��.�</td><td>
<input type=textbox name=t20_4 size="20"></td><td>
<input type=checkbox name = ch20_4 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.5</td><td>������������� �����������</td><td>��.�</td><td>
<input type=textbox name=t20_5 size="20"></td><td>
<input type=checkbox name = ch20_5 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.6</td><td>������� �����������</td><td>��.</td><td>
<input type=textbox name=t20_6 size="20"></td><td>
<input type=checkbox name = ch20_6 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.7</td><td>������� �����</td><td>��.</td><td>
<input type=textbox name=t20_7 size="20"></td><td>
<input type=checkbox name = ch20_7 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.8</td><td>������� ������</td><td>��.</td><td>
<input type=textbox name=t20_8 size="20"></td><td>
<input type=checkbox name = ch20_8 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.9</td><td>����������� ������</td><td>��.</td><td>
<input type=textbox name=t20_9 size="20"></td><td>
<input type=checkbox name = ch20_9 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.10</td><td>���������� ������</td><td>��.</td><td>
<input type=textbox name=t20_10 size="20"></td><td>
<input type=checkbox name = ch20_10 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.11</td><td>������ �����������</td><td>��./��.�</td><td>
<input type=textbox name=t20_11 size="20"></td><td>
<input type=checkbox name = ch20_11 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>6.</td><td><b>������������ � ��� ������</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
<input type=checkbox name = ch21 ></td><td>
<select name="Period21">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>


<tr>
<td>6.1</td><td>������ ���������� �������� ����������� ������</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>�:</td><td>���</td><td>
<input type=textbox name = t22 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>��:</td><td>���</td><td>
<input type=textbox name = t23 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>���� � ����:</td><td>���</td><td>
<input type=textbox name = t24 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>6.2</td><td>������ ���������� �������������� ������</td><td></td><td></td><td></td><td></td>
</tr>

<tr>
<td>&nbsp;</td><td>�:</td><td>���</td><td>
<input type=textbox name = t25 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>��:</td><td>���</td><td>
<input type=textbox name = t26 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>���� � ����:</td><td>���</td><td>
<input type=textbox name = t27 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>



<tr>
<td>6.3</td><td>���������� ���������</td><td>���.</td><td>
<input type=textbox name = t28 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.</td><td><b>��������� ��� ���������� ����������������� ��������� � ������������</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
<input type=checkbox name = ch22 ></td><td>
<select name="Period22">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>

<tr>
<td>7.1</td><td>������� ���������</td><td>��/���</td><td><select name="choice1">
<option value="��" >��</option>
<option value="���" >���</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.2</td><td>������� �������� � ���������</td><td>��/���</td><td><select name="choice2">
<option value="��" >��</option>
<option value="���" >���</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.3</td><td>������� ���������</td><td>��.�</td><td>
<input type=textbox name = t29 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>8.</td><td><b>��������� ��� ���������� ��������������� ��������� � ������������������ ������� ���� (��� �������� ����� 10000 ��.�)</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
<input type=checkbox name = ch23 ></td><td>
<select name="Period23">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>

<tr>
<td>8.1</td><td>������� ���������</td><td>��/���</td><td><select name="choice3">
<option value="��" >��</option>
<option value="���" >���</option>
</select></td><td></td><td></td>
</tr>
<tr>
<td>8.2</td><td>������� ���������</td><td>��.�</td><td>
<input type=textbox name = t30 size="20"></td><td></td><td></td>
</tr>
<tr>
<td>8.3</td><td>���-�� ������������������ ������� ����</td><td>��.</td><td><input type=textbox name = t31 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>



</table>

<br>

<table border=0 width=100%>
<tr><td align=left>��������: ���������� �������� ������, ���������� ������. ������, ���������� � ���� ������������ ������ ���������������!
</td></tr>
<tr><td align=left>2.<b><u>������ ����������:</u></b></td></tr>
</table>


<br>

������������, ����������, ���������� �� ���� �������, � ������� ���������� ����������� ������.
<br>
<table border=1>
<tr>
<td>&nbsp;</td><td>����������</td><td>��.���.</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>1.</td><td><b>����� ��������� ����������</b></td><td>��.�</td><td><input type=textbox name=b1></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>1.1</td><td>������� � ��������� �������� (��������, �������, ������, �������, ������)</td><td>&nbsp;</td><td><input type=textbox name=b2></td><td><input type=checkbox name = k1></td><td>
<select name="b3">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ��� " >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>1.1.1</td><td>��������</td><td>���</td><td><input type=textbox name=b4></td><td><input type=checkbox name = k2></td><td>
<select name="b5">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ��� " >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>1.1.2</td><td>�������</td><td>�.</td><td><input type=textbox name=b6></td><td><input type=checkbox name = k3></td><td>
<select name="b7">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ��� " >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>1.1.3</td><td>������</td><td>�.</td><td><input type=textbox name=b8></td><td><input type=checkbox name = k4></td><td>
<select name="b9">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ��� " >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>1.1.4</td><td>�������</td><td>�.</td><td><input type=textbox name=b10></td><td><input type=checkbox name = k5></td><td>
<select name="b11">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ��� " >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>2.</td><td><b>������������ � ��� ������ ����������</b></td><td>��.�</td><td><input type=textbox name=b12></td><td><input type=checkbox name = k6></td><td>
<select name="b13">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ��� " >��� � ���</option>
</select>
</td>
</tr>

<tr>
<td>2.1</td><td>������ ������</td><td></td><td></td><td></td><td></td>
</tr>

<tr>
<td>&nbsp;</td><td>�:</td><td>���</td><td><input type=textbox name = b14></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>��:</td><td>���</td><td><input type=textbox name = b15></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>���� � ����:</td><td>���</td><td><input type=textbox name = b16></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>


<tr>
<td>2.2</td><td>���������� ���������</td><td>���.</td><td><input type=textbox name = b17></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>


<tr>
<td>3.</td><td><b>������������������ ��������� ������� (������������) ��</b></td><td>���.</td><td><input type=textbox name=b18></td><td><input type=checkbox name = k7></td><td>&nbsp;</td>
</tr>
<tr>
<td>3.1</td><td>������� ���������</td><td>���</td><td><input type=textbox name=b19></td><td><input type=checkbox name = k8></td><td>
<select name="b20">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ��� " >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>3.2</td><td>������� ���������</td><td>���</td><td><input type=textbox name=b21></td><td><input type=checkbox name = k9></td><td>
<select name="b22">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ��� " >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>3.3</td><td>������� ���������</td><td>���</td><td><input type=textbox name=b23></td><td><input type=checkbox name = k10></td><td>
<select name="b24">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ��� " >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>3.4</td><td>������� ���������</td><td>���</td><td><input type=textbox name=b25></td><td><input type=checkbox name = k11></td><td>
<select name="b26">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ��� " >��� � ���</option>
</select>
</td>
</tr>

<tr>
<td>4.</td><td><b>������� ����� ��� ������� ��������� �������,</b> �������������� ���������� � ������� ��� �������� ��������� ���������� � ���������, � ����� ���������� ������� �������, ������������ � ���������</td><td>��/���</td><td><select name="b27">
<option value="��" >��</option>
<option value="���" >���</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>5.</td><td><b>����������� � ������� ������, ������ ����� (������� � ���-�� ���)</b></td><td>��/���</td><td><select name="b28">
<option value="��" >��</option>
<option value="���" >���</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>5.1</td><td>�������</td><td>��.�</td><td><input type=textbox name=b29></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td></td><td>���������� ������� ������� � ������� �����</td><td>���-��</td><td><input type=textbox name=b30></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>6.</td><td><b>�������� ������ ���������� ����� </b>(A - �� �������������� ���� �����/ � - �� ��������)</td><td>�/�</td><td><select name="b31">
<option value="1" selected>�</option>
<option value="2" >�</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.</td><td><b>����� ����� � ������ ������ (����� ������ �� ����)</b></td><td>���.�.</td><td><input type=textbox name=b32></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>

<tr>
<td>8.</td><td><b>����� ���, ��������� �������� ����� ������</b></td><td>���.�.</td><td><input type=textbox name=b33></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

</table>
--------->

<align=left>
<br>
<table border=0 width=100%>
<tr><td align=left>��������: ���������� �������� ������, ���������� ������. ������, ���������� � ���� ������������ ������ ���������������!</td></tr>
<tr><td align=left>3.<b><u>�������������� ������:</u></b></td></tr>
</table>
<table border=1 width=100%>
<tr>
<td>&nbsp;</td><td>����������</td><td>��.���.</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td><b>1.</b></td><td><b>����������� ������ ���������</b></td><td>��/���</td><td><select name="c1">
<option value="<%=session("c1")%>" selected><%=session("c1")%></option>
<option value="��" >��</option>
<option value="���" >���</option>
</select></td><td><input type=checkbox name = c2 value='<%=session("c2")%>'></td>
</tr>
<tr>
<td>1.1</td><td>������� ���������-���������</td><td>��/���</td><td><select name="c3">
<option value="<%=session("c3")%>" selected><%=session("c3")%></option>
<option value="��" >��</option>
<option value="���" >���</option>
</select></td><td><input type=checkbox name = c4 value='<%=session("c4")%>'></td>
</tr>
<tr>
<td>1.1.1</td><td>���������� �������������� ���������</td><td>���.</td><td><input type=text name = c5 value='<%=session("c5")%>'></td><td><input type=checkbox name = c6 value='<%=session("c6")%>'></td>
</tr>
<tr>
<td>1.2</td><td>����������� � ������ ����������, ����� � �����</td><td>��</td><td><input type=text name = c7 value='<%=session("c7")%>'></td><td><input type=checkbox name = c8 value='<%=session("c8")%>'></td>
</tr>
<tr>
<td><b>2.</b></td><td><b>�������� ���������� ������ �������� ��������</b></td><td>��/���</td><td><select name="c9">
<option value="<%=session("c9")%>" selected><%=session("c9")%></option>
<option value="��" selected>��</option>
<option value="���" >���</option>
</select></td><td><input type=checkbox name = c10 value='<%=session("c10")%>'></td>
</tr>
<tr>
<td>2.1</td><td>������� �������� ��������</td><td>��.�</td><td><input type=text name = c11 value='<%=session("c11")%>'></td><td><input type=checkbox name = c12 value='<%=session("c12")%>'></td>
</tr>
<tr>
<td>2.2</td><td>������������� ����������</td><td>&nbsp;</td><td><input type=text name = c13 value='<%=session("c13")%>'></td> <td>&nbsp;</td>
</tr>
<tr>
<td><b>3.</b></td><td><b>��������� ����������� ���� �� ����������� ����</b></td><td>��/���</td><td><select name="c14">
<option value="<%=session("c14")%>" selected><%=session("c14")%></option>
<option value="��" selected>��</option>
<option value="���" >���</option>
</select></td><td><input type=checkbox name = c15 value='<%=session("c15")%>'></td>
</tr>
<tr>
<td>3.1</td><td>������� ���������</td> <td>��.�</td><td><input type=text name = c16 value='<%=session("c16")%>'></td>
</tr>
<tr>
<td>3.2</td><td>������������� ��������� ����</td> <td>&nbsp;</td><td><input type=text name = c17 value='<%=session("c17")%>'></td><td>&nbsp;</td>
</tr>
<tr>
<td><b>4.</b></td><td><b>����� ���� (������� ���� � ����� �������)</b></td> <td>��/���</td><td><select name="c18">
<option value="<%=session("c18")%>" selected><%=session("c18")%></option>
<option value="��" selected>��</option>
<option value="���" >���</option>
</select></td><td><input type=checkbox name = c19 value='<%=session("c19")%>'></td>
</tr>
<tr>
<td>4.1</td><td>������ ������ � ����� (� ����)</td> <td>��.�</td><td><input type=text name = c20 value='<%=session("c20")%>'></td><td><input type=checkbox name = c21 value='<%=session("c21")%>'></td>
</tr>
<tr>
<td>4.2</td><td>������������ (�� ���������)</td> <td>��.�</td><td><input type=text name = c22 value='<%=session("c22")%>'></td><td><input type=checkbox name = c23 value='<%=session("c23")%>'></td>
</tr>
<tr>
<td>4.3</td><td> � ������� ������������ ����������� </td> <td>��.�</td><td><input type=text name = c24 value='<%=session("c24")%>'></td><td><input type=checkbox name = c25 value='<%=session("c25")%>'></td>
</tr>
<tr>
<td><b>5.</b></td><td><b>������ ������, ������� �� ������ �� ��������</b></td><td>&nbsp;</td> <td>&nbsp;</td><td><input type=checkbox name = c26 value='<%=session("c26")%>'></td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td><input type=text name = c27 value='<%=session("c27")%>'></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td><input type=text name = c28 value='<%=session("c28")%>'></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td><input type=text name = c29 value='<%=session("c29")%>'></td><td>&nbsp;</td>
</tr>

<tr>
<td></td><td></td><td>&nbsp;</td><td>&nbsp;</td><td><input type=submit value="��������� >>"></td>
</tr>
</table>
</align>
</form>
<center>
<table border=0 >
<td>
<tr>		
		<td>
		<a href="search.asp"><font  size=3>����� ������ �� ��������</font></a>
		<hr>
		</td>
		<td>
		<a href="finansi.asp"><font  size=3>�������</font></a>
		<hr>
		</td>
		<td>
		<a href="history.asp"><font  size=3>������</font></a>
		<hr>
		</td>
		<td>
		<a href="reg.asp"><font  size=3>������������������</font></a>
		<hr>
		</td>
</tr>
</table>
</center>


</td>
</tr>
		
</table>		

</center>
</FONT>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

<center><FONT size="1" color="#0000ff" style="FONT-WEIGHT: bold">&copy;2005, Solva SoftWare inc. All rights Reserved

</font></center>

  </BODY>
</HTML>