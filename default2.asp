<%
on error resume next

set conn2 = server.CreateObject("Adodb.connection")
      
        connStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & _
         Server.MapPath("missyou.mdb")
         conn2.Open connStr
      ' Recordset object
      
        Set rs3 = Server.CreateObject("ADODB.Recordset")
        sql="select Login,Password from Logins where Login='" &trim(Request.Form("login1"))& "' and " & "Password='"&trim(Request.Form("pass1"))&"'"
       
        Set rs3 = conn2.Execute(sql)
        
        rs3.MoveFirst
        
        if not rs3.EOF  then
			Session("Login")=trim(Request.Form("login1"))
			Session("Password")=trim(Request.Form("pass1"))
		else
			Session("Login")=""
			Session("Password")=""
        end if
        Session("Client")=Request.Form("client")
        Session("contact")=Request.Form("contact")
         
    Dim ID
    Dim rs7
      ID = Request("ID")
      
   If Len(ID) > 0 Then
      
   
   
   ' Connection String
   
      Set rs7 = Server.CreateObject("ADODB.Recordset")
      
      ' opening connection

sql1="select SaleID,ClientID,datebeg,dateend,status,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15," 
sql1= sql1 & "a16,a17,a18,a19,a20,a21,a22,a23,a24,a25,a26,a27,a28,a29,a30,a31,a32,a33,a34,a35,a36,a37,a38,a39,a40,a41,a42,a43"
sql1= sql1 & ",a44,a45,a46,a47,a48,a49,a50,a51,a52,a53,a54,a55,a56,a57,a58,a59,a60,a61,a62,a63,a64,a65,a66,a67,a68,a69,a70,a71,a72,a73"
sql1= sql1 & ",a74,a75,a76,a77,b1,b2,b3,b4,b5,b6,b7,b8,b9,b10,b11,b12,b13,b14,b15,b16,b17,b18,b19,b20,b21,b22,b23,b24,b25,b26,b27,b28,b29,b30,b31,b32,b33"
sql1= sql1 & ",c1,c2,c3,c4,c5,c6,c7,c8,c9,c10,c11,c12,c13,c14,c15,c16,c17,c18,c19,c20,c21,c22,c23,c24,c25,c26,c27,c28,c29,c30,c31,c32,c33 from InternClean " 
sql1= sql1 & " where ID = " & cint(ID) & ""

      set rs7 = conn2.exec(sql1)

  End If


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

<form method=get action='default_next1.asp' ID="Form1">

<table border=0 width=100% ID="Table1">
<tr><td align=left>��������: ���������� �������� ������, ���������� ������. ������, ���������� � ���� ������������ ������ ���������������!
</td></tr>
<tr><td align=left>1.<b><u>���������� ������ ���������:</u></b></u></b></td></tr>
</table>
<br>


������������, ����������, ���������� �� ���� �������, � ������� ���������� ����������� ������.



<table border=1 ID="Table2">
<tr>
<td>&nbsp;</td><td>����������</td><td>��.���.</td><td>������ � 1</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>1.</td><td><b>������� ������ � ������/�������</b></td><td></td><td>�����</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>1.1</td><td>���������</td><td>��.</td><td>
<input type=textbox name=t1 size="20" ID="Textbox1"></td><td>
<input type=checkbox name = ch1 ID="Checkbox1"></td><td>
<select name="Period1" ID="Select1">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>1.2</td><td>��� ���������� �������</td><td>���</td><td>
<input type=textbox name=t2 size="20" ID="Textbox2"></td><td>
<input type=checkbox name = ch2 ID="Checkbox2"></td><td>
<select name="Period2" ID="Select2">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>1.3</td><td>������ ��������</td><td>�.</td><td>
<input type=textbox name=t3 size="20" ID="Textbox3"></td><td>
<input type=checkbox name = ch3 ID="Checkbox3"></td><td>
<select name="Period3" ID="Select3">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>2.</td><td><b>����� �������</b></td><td>��.�</td><td>
<input type=textbox name=t4 size="20" ID="Textbox4"></td><td>
<input type=checkbox name = ch4 ID="Checkbox4"></td><td>
<select name="Period4" ID="Select4">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>2.1</td><td>������� ������� �����</td><td>��.�</td><td>
<input type=textbox name=t5 size="20" ID="Textbox5"></td><td>
<input type=checkbox name = ch5 ID="Checkbox5"></td><td>
<select name="Period5" ID="Select5">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>2.2</td><td>�������� VIP</td><td>���-��</td><td>
<input type=textbox name=t6_1 size="20" ID="Textbox6"></td><td>
<input type=checkbox name = ch6 ID="Checkbox6"></td><td>
<select name="Period6" ID="Select6">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t6_2 size="20" ID="Textbox7"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.3</td><td>������� ���������</td><td>���-��</td><td>
<input type=textbox name=t7_1 size="20" ID="Textbox8"></td><td>
<input type=checkbox name = ch7 ID="Checkbox7"></td><td>
<select name="Period7" ID="Select7">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t7_2 size="20" ID="Textbox9"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.4</td><td>��������� ���������</td><td>���-��</td><td>
<input type=textbox name=t8_1 size="20" ID="Textbox10"></td><td>
<input type=checkbox name = ch8 ID="Checkbox8"></td><td>
<select name="Period8" ID="Select8">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t8_2 size="20" ID="Textbox11"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.5</td><td>��������� ���������</td><td>���-��</td><td>
<input type=textbox name=t9_1 size="20" ID="Textbox12"></td><td>
<input type=checkbox name = ch9 ID="Checkbox9"></td><td>
<select name="Period9" ID="Select9">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t9_2 size="20" ID="Textbox13"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.6</td><td>����������� ���������, �������</td><td>���-��</td><td>
<input type=textbox name=t10_1 size="20" ID="Textbox14"></td><td>
<input type=checkbox name = ch10 ID="Checkbox10"></td><td>
<select name="Period10" ID="Select10">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t10_2 size="20" ID="Textbox15"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>


<tr>
<td>2.7</td><td>��������</td><td>���-��</td><td>
<input type=textbox name=t11_1 size="20" ID="Textbox16"></td><td>
<input type=checkbox name = ch11 ID="Checkbox11"></td><td>
<select name="Period11" ID="Select11">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t11_2 size="20" ID="Textbox17"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.8</td><td>��������</td><td>���-��</td><td>
<input type=textbox name=t12_1 size="20" ID="Textbox18"></td><td>
<input type=checkbox name = ch12 ID="Checkbox12"></td><td>
<select name="Period12" ID="Select12">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t12_2 size="20" ID="Textbox19"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.9</td><td>�����</td><td>���-��</td><td>
<input type=textbox name=t13_1 size="20" ID="Textbox20"></td><td>
<input type=checkbox name = ch13 ID="Checkbox13"></td><td>
<select name="Period13" ID="Select13">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t13_2 size="20" ID="Textbox21"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.10</td><td>����������</td><td>���-��</td><td>
<input type=textbox name=t14_1 size="20" ID="Textbox22"></td><td>
<input type=checkbox name = ch14 ID="Checkbox14"></td><td>
<select name="Period14" ID="Select14">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t14_2 size="20" ID="Textbox23"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.11</td><td>�������</td><td>���-��</td><td>
<input type=textbox name=t15_1 size="20" ID="Textbox24"></td><td>
<input type=checkbox name = ch15 ID="Checkbox15"></td><td>
<select name="Period15" ID="Select15">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t15_2 size="20" ID="Textbox25"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.12</td><td>�����, ���������</td><td>���-��</td><td>
<input type=textbox name=t16_1 size="20" ID="Textbox26"></td><td>
<input type=checkbox name = ch16 ID="Checkbox16"></td><td>
<select name="Period16" ID="Select16">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t16_2 size="20" ID="Textbox27"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.13</td><td>������ ������� (�� ����.������������)</td><td></td><td>
<input type=textbox name=t17_0 size="20" ID="Textbox28"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>���-��</td><td>
<input type=textbox name=t17_1 size="20" ID="Textbox29"></td><td>
<input type=checkbox name = ch17 ID="Checkbox17"></td><td>
<select name="Period17" ID="Select17">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t17_2 size="20" ID="Textbox30"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>


<tr>
<td>3.</td><td><b>�-�� ����������� ����� �������� � �����������, ���. � ������� � ����</b></td><td>���.</td><td>
<input type=textbox name=t18 size="20" ID="Textbox31"></td><td>
<input type=checkbox name = ch18 ID="Checkbox18"></td><td>&nbsp;</td>
</tr>


<tr>
<td>4.</td><td><b>������������� � ����������� �/� ���������� ����������� (��������� ������ � �����)</b></td><td>���-��</td><td>&nbsp;</td><td>&nbsp;</td><td>
<select name="Period19" ID="Select18">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>��������� ������</td><td>���./���.</td><td>
<input type=textbox name=t19_1 size="20" ID="Textbox32"></td><td>
<input type=checkbox name = ch19_1 ID="Checkbox19"></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>������ ����</td><td>����/���.</td><td>
<input type=textbox name=t19_2 size="20" ID="Textbox33"></td><td>
<input type=checkbox name = ch19_2 ID="Checkbox20"></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>��� ���������</td><td>����/���.</td><td>
<input type=textbox name=t19_3 size="20" ID="Textbox34"></td><td>
<input type=checkbox name = ch19_3 ID="Checkbox21"></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>�������� ������� �/�������</td><td>��./���.</td><td>
<input type=textbox name=t19_4 size="20" ID="Textbox35"></td><td>
<input type=checkbox name = ch19_4 ID="Checkbox22"></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.</td><td><b>�����������</b></td><td>���.</td><td>&nbsp;</td><td>
<input type=checkbox name = ch20 ID="Checkbox23"></td><td>
<select name="Period20" ID="Select19">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>

<tr>
<td>5.1</td><td>������ �������� (��������)</td><td>��.�</td><td>
<input type=textbox name=t20_1 size="20" ID="Textbox36"></td><td>
<input type=checkbox name = ch20_1 ID="Checkbox24"></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.2</td><td>����������� �������� (��������, ������, ���������, �������� ���, �������)</td><td>��.�</td><td>
<input type=textbox name=t20_2 size="20" ID="Textbox37"></td><td>
<input type=checkbox name = ch20_2 ID="Checkbox25"></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.3</td><td>������� �������� (������, ������, ������) ���/�����</td><td>��.�</td><td>
<input type=textbox name=t20_3 size="20" ID="Textbox38"></td><td>
<input type=checkbox name = ch20_3 ID="Checkbox26"></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.4</td><td>���������� �����������</td><td>��.�</td><td>
<input type=textbox name=t20_4 size="20" ID="Textbox39"></td><td>
<input type=checkbox name = ch20_4 ID="Checkbox27"></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.5</td><td>������������� �����������</td><td>��.�</td><td>
<input type=textbox name=t20_5 size="20" ID="Textbox40"></td><td>
<input type=checkbox name = ch20_5 ID="Checkbox28"></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.6</td><td>������� �����������</td><td>��.</td><td>
<input type=textbox name=t20_6 size="20" ID="Textbox41"></td><td>
<input type=checkbox name = ch20_6 ID="Checkbox29"></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.7</td><td>������� �����</td><td>��.</td><td>
<input type=textbox name=t20_7 size="20" ID="Textbox42"></td><td>
<input type=checkbox name = ch20_7 ID="Checkbox30"></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.8</td><td>������� ������</td><td>��.</td><td>
<input type=textbox name=t20_8 size="20" ID="Textbox43"></td><td>
<input type=checkbox name = ch20_8 ID="Checkbox31"></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.9</td><td>����������� ������</td><td>��.</td><td>
<input type=textbox name=t20_9 size="20" ID="Textbox44"></td><td>
<input type=checkbox name = ch20_9 ID="Checkbox32"></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.10</td><td>���������� ������</td><td>��.</td><td>
<input type=textbox name=t20_10 size="20" ID="Textbox45"></td><td>
<input type=checkbox name = ch20_10 ID="Checkbox33"></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.11</td><td>������ �����������</td><td>��./��.�</td><td>
<input type=textbox name=t20_11 size="20" ID="Textbox46"></td><td>
<input type=checkbox name = ch20_11 ID="Checkbox34"></td><td>&nbsp;</td>
</tr>

<tr>
<td>6.</td><td><b>������������ � ��� ������</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
<input type=checkbox name = ch21 ID="Checkbox35"></td><td>
<select name="Period21" ID="Select20">
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
<input type=textbox name = t22 size="20" ID="Textbox47"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>��:</td><td>���</td><td>
<input type=textbox name = t23 size="20" ID="Textbox48"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>���� � ����:</td><td>���</td><td>
<input type=textbox name = t24 size="20" ID="Textbox49"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>6.2</td><td>������ ���������� �������������� ������</td><td></td><td></td><td></td><td></td>
</tr>

<tr>
<td>&nbsp;</td><td>�:</td><td>���</td><td>
<input type=textbox name = t25 size="20" ID="Textbox50"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>��:</td><td>���</td><td>
<input type=textbox name = t26 size="20" ID="Textbox51"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>���� � ����:</td><td>���</td><td>
<input type=textbox name = t27 size="20" ID="Textbox52"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>



<tr>
<td>6.3</td><td>���������� ���������</td><td>���.</td><td>
<input type=textbox name = t28 size="20" ID="Textbox53"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.</td><td><b>��������� ��� ���������� ����������������� ��������� � ������������</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
<input type=checkbox name = ch22 ID="Checkbox36"></td><td>
<select name="Period22" ID="Select21">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>

<tr>
<td>7.1</td><td>������� ���������</td><td>��/���</td><td><select name="choice1" ID="Select22">
<option value="��" selected>��</option>
<option value="���" >���</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.2</td><td>������� �������� � ���������</td><td>��/���</td><td><select name="choice2" ID="Select23">
<option value="��" selected>��</option>
<option value="���" >���</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.3</td><td>������� ���������</td><td>��.�</td><td>
<input type=textbox name = t29 size="20" ID="Textbox54"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>8.</td><td><b>��������� ��� ���������� ��������������� ��������� � ������������������ ������� ���� (��� �������� ����� 10000 ��.�)</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
<input type=checkbox name = ch23 ID="Checkbox37"></td><td>
<select name="Period23" ID="Select24">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>

<tr>
<td>8.1</td><td>������� ���������</td><td>��/���</td><td><select name="choice3" ID="Select25">
<option value="��" selected>��</option>
<option value="���" >���</option>
</select></td><td></td><td></td>
</tr>
<tr>
<td>8.2</td><td>������� ���������</td><td>��.�</td><td>
<input type=textbox name = t30 size="20" ID="Textbox55"></td><td></td><td></td>
</tr>
<tr>
<td>8.3</td><td>���-�� ������������������ ������� ����</td><td>��.</td><td><input type=textbox name = t31 size="20" ID="Textbox56"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>



</table>
<br>
<!------
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
<option value="1" selected>�������</option>
<option value="2" >�����</option>
<option value="3" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>1.1.1</td><td>��������</td><td>���</td><td><input type=textbox name=b4></td><td><input type=checkbox name = k2></td><td>
<select name="b5">
<option value="1" selected>�������</option>
<option value="2" >�����</option>
<option value="3" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>1.1.2</td><td>�������</td><td>�.</td><td><input type=textbox name=b6></td><td><input type=checkbox name = k3></td><td>
<select name="b7">
<option value="1" selected>�������</option>
<option value="2" >�����</option>
<option value="3" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>1.1.3</td><td>������</td><td>�.</td><td><input type=textbox name=b8></td><td><input type=checkbox name = k4></td><td>
<select name="b9">
<option value="1" selected>�������</option>
<option value="2" >�����</option>
<option value="3" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>1.1.4</td><td>�������</td><td>�.</td><td><input type=textbox name=b10></td><td><input type=checkbox name = k5></td><td>
<select name="b11">
<option value="1" selected>�������</option>
<option value="2" >�����</option>
<option value="3" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>2.</td><td><b>������������ � ��� ������ ����������</b></td><td>��.�</td><td><input type=textbox name=b12></td><td><input type=checkbox name = k6></td><td>
<select name="b13">
<option value="1" selected>�������</option>
<option value="2" >�����</option>
<option value="3" >��� � ���</option>
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
<option value="1" selected>�������</option>
<option value="2" >�����</option>
<option value="3" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>3.2</td><td>������� ���������</td><td>���</td><td><input type=textbox name=b21></td><td><input type=checkbox name = k9></td><td>
<select name="b22">
<option value="1" selected>�������</option>
<option value="2" >�����</option>
<option value="3" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>3.3</td><td>������� ���������</td><td>���</td><td><input type=textbox name=b23></td><td><input type=checkbox name = k10></td><td>
<select name="b24">
<option value="1" selected>�������</option>
<option value="2" >�����</option>
<option value="3" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>3.4</td><td>������� ���������</td><td>���</td><td><input type=textbox name=b25></td><td><input type=checkbox name = k11></td><td>
<select name="b26">
<option value="1" selected>�������</option>
<option value="2" >�����</option>
<option value="3" >��� � ���</option>
</select>
</td>
</tr>

<tr>
<td>4.</td><td><b>������� ����� ��� ������� ��������� �������,</b> �������������� ���������� � ������� ��� �������� ��������� ���������� � ���������, � ����� ���������� ������� �������, ������������ � ���������</td><td>��/���</td><td><select name="b27">
<option value="1" selected>��</option>
<option value="2" >���</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>5.</td><td><b>����������� � ������� ������, ������ ����� (������� � ���-�� ���)</b></td><td>��/���</td><td><select name="b28">
<option value="1" selected>��</option>
<option value="2" >���</option>
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
<option value="��" selected>��</option>
<option value="���" >���</option>
</select></td><td><input type=checkbox name = c2></td>
</tr>
<tr>
<td>1.1</td><td>������� ���������-���������</td><td>��/���</td><td><select name="c3">
<option value="��" selected>��</option>
<option value="���" >���</option>
</select></td><td><input type=checkbox name = c4></td>
</tr>
<tr>
<td>1.1.1</td><td>���������� �������������� ���������</td><td>���.</td><td><input type=text name = c5></td><td><input type=checkbox name = c6></td>
</tr>
<tr>
<td>1.2</td><td>����������� � ������ ����������, ����� � �����</td><td>��</td><td><input type=text name = c7></td><td><input type=checkbox name = c8></td>
</tr>
<tr>
<td><b>2.</b></td><td><b>�������� ���������� ������ �������� ��������</b></td><td>��/���</td><td><select name="c9">
<option value="��" selected>��</option>
<option value="���" >���</option>
</select></td><td><input type=checkbox name = c10></td>
</tr>
<tr>
<td>2.1</td><td>������� �������� ��������</td><td>��.�</td><td><input type=text name = c11></td><td><input type=checkbox name = c12></td>
</tr>
<tr>
<td>2.2</td><td>������������� ����������</td><td>&nbsp;</td><td><input type=text name = c13></td> <td>&nbsp;</td>
</tr>
<tr>
<td><b>3.</b></td><td><b>��������� ����������� ���� �� ����������� ����</b></td><td>��/���</td><td><select name="c14">
<option value="��" selected>��</option>
<option value="���" >���</option>
</select></td><td><input type=checkbox name = c15></td>
</tr>
<tr>
<td>3.1</td><td>������� ���������</td> <td>��.�</td><td><input type=text name = c16></td>
</tr>
<tr>
<td>3.2</td><td>������������� ��������� ����</td> <td>&nbsp;</td><td><input type=text name = c16></td><td>&nbsp;</td>
</tr>
<tr>
<td><b>4.</b></td><td><b>����� ���� (������� ���� � ����� �������)</b></td> <td>��/���</td><td><select name="c17">
<option value="��" selected>��</option>
<option value="���" >���</option>
</select></td><td><input type=checkbox name = c18></td>
</tr>
<tr>
<td>4.1</td><td>������ ������ � ����� (� ����)</td> <td>��.�</td><td><input type=text name = c19></td><td><input type=checkbox name = c20></td>
</tr>
<tr>
<td>4.2</td><td>������������ (�� ���������)</td> <td>��.�</td><td><input type=text name = c21></td><td><input type=checkbox name = c22></td>
</tr>
<tr>
<td>4.3</td><td> � ������� ������������ ����������� </td> <td>��.�</td><td><input type=text name = c23></td><td><input type=checkbox name = c24></td>
</tr>
<tr>
<td><b>5.</b></td><td><b>������ ������, ������� �� ������ �� ��������</b></td><td>&nbsp;</td> <td>&nbsp;</td><td><input type=checkbox name = c25></td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td><input type=text name = c26></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td><input type=text name = c27></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td><input type=text name = c28></td><td>&nbsp;</td>
</tr>
----->
<tr>
<td></td><td></td><td>&nbsp;</td><td>&nbsp;</td><td><input type=submit value="������ >>" ID="Submit1" NAME="Submit1"></td>
</tr>
</table>
</align>

<center>
<table border=0 ID="Table3">
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