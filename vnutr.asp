<%
on error resume next
if Session("NickName")<>Session("pr") then Session("pr")=Session("NickName")

set conn2 = server.CreateObject("Adodb.connection")
      
        connStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & _
         Server.MapPath("missyou.mdb")
         conn2.Open connStr
      ' Recordset object
      


        Set rs = Server.CreateObject("ADODB.Recordset")
        sql1="select [SaleID],[ClientID],[datebeg],[dateend],[status],[a1],[a2],[a3],[a4],[a5],[a6],[a7],[a8],[a9],[a10],[a11],[a12],[a13],[a14],[a15]," 
sql1= sql1 &"[a16],[a17],[a18],[a19],[a20],[a21],[a22],[a23],[a24],[a25],[a26],[a27],[a28],[a29],[a30],[a31],[a32],[a33],[a34],[a35],[a36],[a37],[a38],[a39],[a40],[a41],[a42],[a43]"
sql1= sql1 & ",[a44],[a45],[a46],[a47],[a48],[a49],[a50],[a51],[a52],[a53],[a54],[a55],[a56],[a57],[a58],[a59],[a60],[a61],[a62],[a63],[a64],[a65],[a66],[a67],[a68],[a69],[a70],[a71],[a72],[a73]"
sql1= sql1 & ",[a74],[a75],[a76],[a77],[a78],[a79],[a80],[a81],[a82],[a83],[a84],[a85],[a86],[a87],[a88],[a89],[a90],[a91],[a92],[a93],[a94],[a95],[a96],[a97],[a98],[a99],[a100]"
sql1= sql1 & ",[a101],[a102],[a103],[a104],[a105],[a106],[a107],[a108],[a109],[a110],[a111],[a112],[a113],[a114],[a115],[a116],[a117],[a118],"
sql1= sql1 & "[b1],[b2],[b3],[b4],[b5],[b6],[b7],[b8],[b9],[b10],[b11],[b12],[b13],[b14],[b15],[b16],[b17],[b18],[b19],[b20],[b21],[b22],[b23],[b24],[b25],[b26],[b27],[b28],[b29],[b30],[b31],[b32],[b33]"
sql1= sql1 & ",[b34],[b35],[b36],[b37],[b38],[b39],[b40],[b41],[b42],[b43],[b44]"
sql1= sql1 & ",[c1],[c2],[c3],[c4],[c5],[c6],[c7],[c8],[c9],[c10],[c11],[c12],[c13],[c14],[c15],[c16],[c17],[c18],[c19],[c20],[c21],[c22],[c23],[c24],[c25],[c26],[c27],[c28],[c29],[c30],[c31],[c32],[c33] from InternClean " 
sql1= sql1 & " where ID = " & Session("ID") & ""
       
        Set rs = conn2.Execute(sql1)
        
        rs.MoveFirst



'response.write sql1

'response.write rs(0)
'response.write err.description

session("a1")=rs("a1")
session("a2")=rs("a2")
session("a3")=rs("a3")
session("a4")=rs("a4")
session("a5")=rs("a5")
session("a6")=rs("a6")
session("a7")=rs("a7")
session("a8")=rs("a8")
session("a9")=rs("a9")
session("a10")=rs("a10")
session("a11")=rs("a11")
session("a12")=rs("a12")
session("a13")=rs("a13")
session("a14")=rs("a14")
session("a15")=rs("a15")
session("a16")=rs("a16")
session("a17")=rs("a17")
session("a18")=rs("a18")
session("a19")=rs("a19")
session("a20")=rs("a20")
session("a21")=rs("a21")
session("a22")=rs("a22")
session("a23")=rs("a23")
session("a24")=rs("a24")
session("a25")=rs("a25")
session("a26")=rs("a26")
session("a27")=rs("a27")
session("a28")=rs("a28")
session("a29")=rs("a29")
session("a30")=rs("a30")
session("a31")=rs("a31")
session("a32")=rs("a32")
session("a33")=rs("a33")
session("a34")=rs("a34")
session("a35")=rs("a35")
session("a36")=rs("a36")
session("a37")=rs("a37")
session("a38")=rs("a38")
session("a39")=rs("a39")
session("a40")=rs("a40")
session("a41")=rs("a41")
session("a42")=rs("a42")
session("a43")=rs("a43")
session("a44")=rs("a44")
session("a45")=rs("a45")
session("a46")=rs("a46")
session("a47")=rs("a47")
session("a48")=rs("a48")
session("a49")=rs("a49")
session("a50")=rs("a50")
session("a51")=rs("a51")
session("a52")=rs("a52")
session("a53")=rs("a53")
session("a54")=rs("a54")
session("a55")=rs("a55")
session("a56")=rs("a56")
session("a57")=rs("a57")
session("a58")=rs("a58")
session("a59")=rs("a59")
session("a60")=rs("a60")
session("a61")=rs("a61")
session("a62")=rs("a62")
session("a63")=rs("a63")
session("a64")=rs("a64")
session("a65")=rs("a65")
session("a66")=rs("a66")
session("a67")=rs("a67")
session("a68")=rs("a68")
session("a69")=rs("a69")
session("a70")=rs("a70")
session("a71")=rs("a71")
session("a72")=rs("a72")
session("a73")=rs("a73")
session("a74")=rs("a74")
session("a75")=rs("a75")
session("a76")=rs("a76")
session("a77")=rs("a77")
session("a78")=rs("a78")
session("a79")=rs("a79")
session("a80")=rs("a80")
session("a81")=rs("a81")
session("a82")=rs("a82")
session("a83")=rs("a83")
session("a84")=rs("a84")
session("a85")=rs("a85")
session("a86")=rs("a86")
session("a87")=rs("a87")
session("a88")=rs("a88")
session("a89")=rs("a89")
session("a90")=rs("a90")
session("a91")=rs("a91")
session("a92")=rs("a92")
session("a93")=rs("a93")
session("a94")=rs("a94")
session("a95")=rs("a95")
session("a96")=rs("a96")
session("a97")=rs("a97")
session("a98")=rs("a98")
session("a99")=rs("a99")
session("a100")=rs("a100")
session("a101")=rs("a101")
session("a102")=rs("a102")
session("a103")=rs("a103")
session("a104")=rs("a104")
session("a105")=rs("a105")
session("a106")=rs("a106")
session("a107")=rs("a107")
session("a108")=rs("a108")
session("a109")=rs("a109")
session("a110")=rs("a110")
session("a111")=rs("a111")
session("a112")=rs("a112")
session("a113")=rs("a113")
session("a114")=rs("a114")
session("a115")=rs("a115")
session("a116")=rs("a116")
session("a117")=rs("a117")
session("a118")=rs("a118")

session("b1")=rs("b1")
session("b2")=rs("b2")
session("b3")=rs("b3")
session("b4")=rs("b4")
session("b5")=rs("b5")
session("b6")=rs("b6")
session("b7")=rs("b7")
session("b8")=rs("b8")
session("b9")=rs("b9")
session("b10")=rs("b10")
session("b11")=rs("b11")
session("b12")=rs("b12")
session("b13")=rs("b13")
session("b14")=rs("b14")
session("b15")=rs("b15")
session("b16")=rs("b16")
session("b17")=rs("b17")
session("b18")=rs("b18")
session("b19")=rs("b19")
session("b20")=rs("b20")
session("b21")=rs("b21")
session("b22")=rs("b22")
session("b23")=rs("b23")
session("b24")=rs("b24")
session("b25")=rs("b25")
session("b26")=rs("b26")
session("b27")=rs("b27")
session("b28")=rs("b28")
session("b29")=rs("b29")
session("b30")=rs("b30")
session("b31")=rs("b31")
session("b32")=rs("b32")
session("b33")=rs("b33")
session("b34")=rs("b34")
session("b35")=rs("b35")
session("b36")=rs("b36")
session("b37")=rs("b37")
session("b38")=rs("b38")
session("b39")=rs("b39")
session("b40")=rs("b40")
session("b41")=rs("b41")
session("b42")=rs("b42")
session("b43")=rs("b43")
session("b44")=rs("b44")

session("c1")=rs("c1")
session("c2")=rs("c2")
session("c3")=rs("c3")
session("c4")=rs("c4")
session("c5")=rs("c5")
session("c6")=rs("c6")
session("c7")=rs("c7")
session("c8")=rs("c8")
session("c9")=rs("c9")
session("c10")=rs("c10")
session("c11")=rs("c11")
session("c12")=rs("c12")
session("c13")=rs("c13")
session("c14")=rs("c14")
session("c15")=rs("c15")
session("c16")=rs("c16")
session("c17")=rs("c17")
session("c18")=rs("c18")
session("c19")=rs("c19")
session("c20")=rs("c20")
session("c21")=rs("c21")
session("c22")=rs("c22")
session("c23")=rs("c23")
session("c24")=rs("c24")
session("c25")=rs("c25")
session("c26")=rs("c26")
session("c27")=rs("c27")
session("c28")=rs("c28")
session("c29")=rs("c29")
session("c30")=rs("c30")
session("c31")=rs("c31")
session("c32")=rs("c32")
session("c33")=rs("c33")

'response.write session("a1")






%>
<HTML>
  <HEAD>
  
    <title>Default page</title>
     <link rel="stylesheet" type="text/css" href="index.css">
<META content="text/html; charset=windows-1251" http-equiv=Content-Type> 
<META NAME="Keywords" CONTENT="">

  </HEAD>
  <BODY background="tiptop/bg.jpg" alink="#0000ff" vlink="#0000ff" link="#0000ff">


        
<center>
<img src="relationshipsromance.jpeg"></img>
</center>


<br>
<br>
<br>
<table>
<tr>
<td><font size=3>�������� ������: <%=Session("pr")%></font></td>
</tr>

<tr>
<td><font size=3>���������: <%=Session("login")%></font></td>
</tr>
<tr>
 <table>
  
 <td valign="top">
<a href="common.asp?ID=<%=Session("ID")%>"><font  size=1>����� ������(�������������)</font></a>
		<hr>
		<a href="vnutr.asp?ID=<%=Session("ID")%>"><font  size=1>������ (�������������)</font></a>
		<hr>


 </td>
<td>


<form method=get action='save1.asp'>
<!---------
<table border=0 width=100%>
<tr><td align=left>��������: ���������� �������� ������, ���������� ������. ������, ���������� � ���� ������������ ������ ���������������!
</td></tr>
<tr><td align=left>1.<b><u>���������� ������ ���������:</u></b></u></b></td></tr>
</table>
<br>


������������, ����������, ���������� �� ���� �������, � ������� ���������� ����������� ������.
----->


<table border=1>
<tr>
<td>&nbsp;</td><td>����������</td><td>��.���.</td><td>������ � 1</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>1.</td><td><b>������� ������ � ������/�������</b></td><td></td><td>�����</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>1.1</td><td>���������</td><td>��.</td><td>
<input type=textbox name=t1 size="20" value='<%=session("a1")%>'></td><td>
<input type=checkbox name = ch1  value='<%=session("a2")%>'></td><td>
<select name="Period1">
<option value="<%=session("a3")%>" selected><%=session("a3")%></option>
<option value="�������" >�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>1.2</td><td>��� ���������� �������</td><td>���</td><td>
<input type=textbox name=t2 size="20"  value='<%=session("a4")%>'></td><td>
<input type=checkbox name = ch2  value='<%=session("a5")%>'></td><td>
<select name="Period2">
<option value="<%=session("a6")%>" selected><%=session("a6")%></option>
<option value="�������">�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>1.3</td><td>������ ��������</td><td>�.</td><td>
<input type=textbox name=t3 size="20"  value='<%=session("a7")%>'></td><td>
<input type=checkbox name = ch3  value='<%=session("a8")%>'></td><td>
<select name="Period3">
<option value="<%=session("a9")%>" selected><%=session("a9")%></option>
<option value="�������">�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>2.</td><td><b>����� �������</b></td><td>��.�</td><td>
<input type=textbox name=t4 size="20"  value='<%=session("a10")%>'></td><td>
<input type=checkbox name = ch4  value='<%=session("a11")%>'></td><td>
<select name="Period4">
<option value="<%=session("a12")%>" selected><%=session("a12")%></option>
<option value="�������">�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>2.1</td><td>������� ������� �����</td><td>��.�</td><td>
<input type=textbox name=t5 size="20"  value='<%=session("a13")%>'></td><td>
<input type=checkbox name = ch5  value='<%=session("a14")%>'></td><td>
<select name="Period5">
<option value="<%=session("a15")%>" selected><%=session("a15")%></option>
<option value="�������" >�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>2.2</td><td>�������� VIP</td><td>���-��</td><td>
<input type=textbox name=t6_1 size="20"  value='<%=session("a16")%>'></td><td>
<input type=checkbox name = ch6  value='<%=session("a17")%>'></td><td>
<select name="Period6">
<option value="<%=session("a18")%>" selected><%=session("a18")%></option>
<option value="�������">�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t6_2 size="20"  value='<%=session("a19")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.3</td><td>������� ���������</td><td>���-��</td><td>
<input type=textbox name=t7_1 size="20"  value='<%=session("a20")%>'></td><td>
<input type=checkbox name = ch7  value='<%=session("a21")%>'></td><td>
<select name="Period7">
<option value="<%=session("a22")%>" selected><%=session("a22")%></option>
<option value="�������">�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t7_2 size="20"  value='<%=session("a23")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.4</td><td>��������� ���������</td><td>���-��</td><td>
<input type=textbox name=t8_1 size="20"  value='<%=session("a24")%>'></td><td>
<input type=checkbox name = ch8  value='<%=session("a25")%>'></td><td>
<select name="Period8">
<option value="<%=session("a26")%>" selected><%=session("a26")%></option>
<option value="�������">�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t8_2 size="20"  value='<%=session("a27")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.5</td><td>��������� ���������</td><td>���-��</td><td>
<input type=textbox name=t9_1 size="20"  value='<%=session("a28")%>'></td><td>
<input type=checkbox name = ch9  value='<%=session("a29")%>'></td><td>
<select name="Period9">
<option value="<%=session("a30")%>" selected><%=session("a30")%></option>
<option value="�������">�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t9_2 size="20"  value='<%=session("a31")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.6</td><td>����������� ���������, �������</td><td>���-��</td><td>
<input type=textbox name=t10_1 size="20"  value='<%=session("a32")%>'></td><td>
<input type=checkbox name = ch10  value='<%=session("a33")%>'></td><td>
<select name="Period10">
<option value="<%=session("a34")%>" selected><%=session("a34")%></option>
<option value="�������">�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t10_2 size="20"  value='<%=session("a35")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>


<tr>
<td>2.7</td><td>��������</td><td>���-��</td><td>
<input type=textbox name=t11_1 size="20"  value='<%=session("a36")%>'></td><td>
<input type=checkbox name = ch11  value='<%=session("a37")%>'></td><td>
<select name="Period11">
<option value="<%=session("a38")%>" selected><%=session("a38")%></option>
<option value="�������">�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t11_2 size="20"  value='<%=session("a39")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.8</td><td>��������</td><td>���-��</td><td>
<input type=textbox name=t12_1 size="20"  value='<%=session("a40")%>'></td><td>
<input type=checkbox name = ch12  value='<%=session("a41")%>'></td><td>
<select name="Period12">
<option value="<%=session("a42")%>" selected><%=session("a42")%></option>
<option value="�������">�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t12_2 size="20"  value='<%=session("a43")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.9</td><td>�����</td><td>���-��</td><td>
<input type=textbox name=t13_1 size="20"  value='<%=session("a44")%>'></td><td>
<input type=checkbox name = ch13  value='<%=session("a45")%>'></td><td>
<select name="Period13">
<option value="<%=session("a46")%>" selected><%=session("a46")%></option>
<option value="�������" >�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t13_2 size="20"  value='<%=session("a47")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.10</td><td>����������</td><td>���-��</td><td>
<input type=textbox name=t14_1 size="20"  value='<%=session("a48")%>'></td><td>
<input type=checkbox name = ch14  value='<%=session("a49")%>'></td><td>
<select name="Period14">
<option value="<%=session("a50")%>" selected><%=session("a50")%></option>
<option value="�������">�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t14_2 size="20"  value='<%=session("a51")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.11</td><td>�������</td><td>���-��</td><td>
<input type=textbox name=t15_1 size="20"  value='<%=session("a52")%>'></td><td>
<input type=checkbox name = ch15  value='<%=session("a53")%>'></td><td>
<select name="Period15">
<option value="<%=session("a54")%>" selected><%=session("a54")%></option>
<option value="�������">�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t15_2 size="20"  value='<%=session("a55")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.12</td><td>�����, ���������</td><td>���-��</td><td>
<input type=textbox name=t16_1 size="20"  value='<%=session("a56")%>'></td><td>
<input type=checkbox name = ch16  value='<%=session("a57")%>'></td><td>
<select name="Period16">
<option value="<%=session("a58")%>" selected><%=session("a58")%></option>
<option value="�������">�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t16_2 size="20"  value='<%=session("a59")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.13</td><td>������ ������� (�� ����.������������)</td><td></td><td>
<input type=textbox name=t17_0 size="20"  value='<%=session("a60")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>���-��</td><td>
<input type=textbox name=t17_1 size="20"  value='<%=session("a61")%>'></td><td>
<input type=checkbox name = ch17  value='<%=session("a62")%>'></td><td>
<select name="Period17">
<option value="<%=session("a63")%>" selected><%=session("a63")%></option>
<option value="�������">�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<input type=textbox name=t17_2 size="20"  value='<%=session("a64")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>


<tr>
<td>3.</td><td><b>�-�� ����������� ����� �������� � �����������, ���. � ������� � ����</b></td><td>���.</td><td>
<input type=textbox name=t18 size="20"  value='<%=session("a65")%>'></td><td>
<input type=checkbox name = ch18  value='<%=session("a66")%>'></td><td>&nbsp;</td>
</tr>


<tr>
<td>4.</td><td><b>������������� � ����������� �/� ���������� ����������� (��������� ������ � �����)</b></td><td>���-��</td><td>&nbsp;</td><td>&nbsp;</td><td>
<select name="Period19">
<option value="<%=session("a67")%>" selected><%=session("a67")%></option>
<option value="�������">�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>��������� ������</td><td>���./���.</td><td>
<input type=textbox name=t19_1 size="20"  value='<%=session("a68")%>'></td><td>
<input type=checkbox name = ch19_1  value='<%=session("a69")%>'></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>������ ����</td><td>����/���.</td><td>
<input type=textbox name=t19_2 size="20"  value='<%=session("a70")%>'></td><td>
<input type=checkbox name = ch19_2  value='<%=session("a71")%>'></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>��� ���������</td><td>����/���.</td><td>
<input type=textbox name=t19_3 size="20"  value='<%=session("a72")%>'></td><td>
<input type=checkbox name = ch19_3  value='<%=session("a73")%>'></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>�������� ������� �/�������</td><td>��./���.</td><td>
<input type=textbox name=t19_4 size="20"  value='<%=session("a74")%>'></td><td>
<input type=checkbox name = ch19_4  value='<%=session("a75")%>'></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.</td><td><b>�����������</b></td><td>���.</td><td>&nbsp;</td><td>
<input type=checkbox name = ch20  value='<%=session("a76")%>'></td><td>
<select name="Period20">
<option value="<%=session("a77")%>" selected><%=session("a77")%></option>
<option value="�������">�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>

<tr>
<td>5.1</td><td>������ �������� (��������)</td><td>��.�</td><td>
<input type=textbox name=t20_1 size="20"  value='<%=session("a78")%>'></td><td>
<input type=checkbox name = ch20_1  value='<%=session("a79")%>'></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.2</td><td>����������� �������� (��������, ������, ���������, �������� ���, �������)</td><td>��.�</td><td>
<input type=textbox name=t20_2 size="20"  value='<%=session("a80")%>'></td><td>
<input type=checkbox name = ch20_2  value='<%=session("a81")%>'></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.3</td><td>������� �������� (������, ������, ������) ���/�����</td><td>��.�</td><td>
<input type=textbox name=t20_3 size="20"  value='<%=session("a82")%>'></td><td>
<input type=checkbox name = ch20_3  value='<%=session("a83")%>'></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.4</td><td>���������� �����������</td><td>��.�</td><td>
<input type=textbox name=t20_4 size="20"  value='<%=session("a84")%>'></td><td>
<input type=checkbox name = ch20_4  value='<%=session("a85")%>'></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.5</td><td>������������� �����������</td><td>��.�</td><td>
<input type=textbox name=t20_5 size="20"  value='<%=session("a86")%>'></td><td>
<input type=checkbox name = ch20_5  value='<%=session("a87")%>'></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.6</td><td>������� �����������</td><td>��.</td><td>
<input type=textbox name=t20_6 size="20"  value='<%=session("a88")%>'></td><td>
<input type=checkbox name = ch20_6  value='<%=session("a89")%>'></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.7</td><td>������� �����</td><td>��.</td><td>
<input type=textbox name=t20_7 size="20"  value='<%=session("a90")%>'></td><td>
<input type=checkbox name = ch20_7  value='<%=session("a91")%>'></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.8</td><td>������� ������</td><td>��.</td><td>
<input type=textbox name=t20_8 size="20"  value='<%=session("a92")%>'></td><td>
<input type=checkbox name = ch20_8  value='<%=session("a93")%>'></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.9</td><td>����������� ������</td><td>��.</td><td>
<input type=textbox name=t20_9 size="20"  value='<%=session("a94")%>'></td><td>
<input type=checkbox name = ch20_9  value='<%=session("a95")%>'></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.10</td><td>���������� ������</td><td>��.</td><td>
<input type=textbox name=t20_10 size="20"  value='<%=session("a96")%>'></td><td>
<input type=checkbox name = ch20_10  value='<%=session("a97")%>'></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.11</td><td>������ �����������</td><td>��./��.�</td><td>
<input type=textbox name=t20_11 size="20"  value='<%=session("a98")%>'></td><td>
<input type=checkbox name = ch20_11  value='<%=session("a99")%>'></td><td>&nbsp;</td>
</tr>

<tr>
<td>6.</td><td><b>������������ � ��� ������</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
<input type=checkbox name = ch21  value='<%=session("a100")%>'></td><td>
<select name="Period21">
<option value="<%=session("a101")%>" selected><%=session("a101")%></option>
<option value="�������">�������</option>
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
<input type=textbox name = t22 size="20"  value='<%=session("a102")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>��:</td><td>���</td><td>
<input type=textbox name = t23 size="20"  value='<%=session("a103")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>���� � ����:</td><td>���</td><td>
<input type=textbox name = t24 size="20"  value='<%=session("a104")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>6.2</td><td>������ ���������� �������������� ������</td><td></td><td></td><td></td><td></td>
</tr>

<tr>
<td>&nbsp;</td><td>�:</td><td>���</td><td>
<input type=textbox name = t25 size="20"  value='<%=session("a105")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>��:</td><td>���</td><td>
<input type=textbox name = t26 size="20"  value='<%=session("a106")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>���� � ����:</td><td>���</td><td>
<input type=textbox name = t27 size="20"  value='<%=session("a107")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>



<tr>
<td>6.3</td><td>���������� ���������</td><td>���.</td><td>
<input type=textbox name = t28 size="20"  value='<%=session("a108")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.</td><td><b>��������� ��� ���������� ����������������� ��������� � ������������</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
<input type=checkbox name = ch22  value='<%=session("a109")%>'></td><td>
<select name="Period22">
<option value="<%=session("a110")%>" selected><%=session("a110")%></option>
<option value="�������">�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>

<tr>
<td>7.1</td><td>������� ���������</td><td>��/���</td><td><select name="choice1">
<option value="<%=session("a111")%>" selected><%=session("a111")%></option>
<option value="��" >��</option>
<option value="���" >���</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.2</td><td>������� �������� � ���������</td><td>��/���</td><td><select name="choice2">
<option value="<%=session("a112")%>" selected><%=session("a112")%></option>
<option value="��">��</option>
<option value="���" >���</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.3</td><td>������� ���������</td><td>��.�</td><td>
<input type=textbox name = t29 size="20"  value='<%=session("a113")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>8.</td><td><b>��������� ��� ���������� ��������������� ��������� � ������������������ ������� ���� (��� �������� ����� 10000 ��.�)</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
<input type=checkbox name = ch23 value='<%=session("a114")%>'></td><td>
<select name="Period23">
<option value="<%=session("a115")%>" selected><%=session("a115")%></option>
<option value="�������">�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>

<tr>
<td>8.1</td><td>������� ���������</td><td>��/���</td><td><select name="choice3">
<option value="<%=session("a116")%>" selected><%=session("a116")%></option>
<option value="��" >��</option>
<option value="���" >���</option>
</select></td><td></td><td></td>
</tr>
<tr>
<td>8.2</td><td>������� ���������</td><td>��.�</td><td>
<input type=textbox name = t30 size="20" value='<%=session("a117")%>'></td><td></td><td></td>
</tr>
<tr>
<td>8.3</td><td>���-�� ������������������ ������� ����</td><td>��.</td><td><input type=textbox name = t31 size="20" value='<%=session("a118")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
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
<td>1.</td><td><b>����� ��������� ����������</b></td><td>��.�</td><td><input type=textbox name=b1 value='<%=session("b1")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>1.1</td><td>������� � ��������� �������� (��������, �������, ������, �������, ������)</td>
<td>&nbsp;</td><td><input type=textbox name=b2 value='<%=session("b2")%>'></td>
<td><input type=checkbox name = k1 value='<%=session("b3")%>'></td><td>
<select name="b3">
<option value="<%=session("b4")%>" selected><%=session("b4")%></option>
<option value="1" >�������</option>
<option value="2" >�����</option>
<option value="3" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>1.1.1</td><td>��������</td><td>���</td>
<td><input type=textbox name=b4  value='<%=session("b5")%>'></td><td><input type=checkbox name = k2  value='<%=session("b6")%>'></td><td>
<select name="b5">
<option value="<%=session("b7")%>" selected><%=session("b7")%></option>
<option value="1" >�������</option>
<option value="2" >�����</option>
<option value="3" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>1.1.2</td><td>�������</td><td>�.</td><td><input type=textbox name=b6  value='<%=session("b8")%>'></td>
<td><input type=checkbox name = k3  value='<%=session("b9")%>'></td><td>
<select name="b7">
<option value="<%=session("b10")%>" selected><%=session("b10")%></option>
<option value="1" >�������</option>
<option value="2" >�����</option>
<option value="3" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>1.1.3</td><td>������</td><td>�.</td><td><input type=textbox name=b8 value='<%=session("b11")%>'></td><td><input type=checkbox name = k4 value='<%=session("a1")%>'></td><td>
<select name="b9">
<option value="<%=session("b12")%>" selected><%=session("b12")%></option>
<option value="1" >�������</option>
<option value="2" >�����</option>
<option value="3" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>1.1.4</td><td>�������</td><td>�.</td><td><input type=textbox name=b10 value='<%=session("b13")%>'></td>
<td><input type=checkbox name = k5 value='<%=session("b14")%>'></td><td>
<select name="b11">
<option value="<%=session("b15")%>" selected><%=session("b15")%></option>
<option value="1" >�������</option>
<option value="2" >�����</option>
<option value="3" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>2.</td><td><b>������������ � ��� ������ ����������</b></td>
<td>��.�</td><td><input type=textbox name=b12 value='<%=session("b16")%>'></td><td><input type=checkbox name = k6 value='<%=session("b17")%>'></td><td>
<select name="b13">
<option value="<%=session("b18")%>" selected><%=session("b18")%></option>
<option value="1" >�������</option>
<option value="2" >�����</option>
<option value="3" >��� � ���</option>
</select>
</td>
</tr>

<tr>
<td>2.1</td><td>������ ������</td><td></td><td></td><td></td><td></td>
</tr>

<tr>
<td>&nbsp;</td><td>�:</td><td>���</td><td><input type=textbox name = b14 value='<%=session("b19")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>��:</td><td>���</td><td><input type=textbox name = b15 value='<%=session("b20")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>���� � ����:</td><td>���</td><td><input type=textbox name = b16 value='<%=session("b21")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>


<tr>
<td>2.2</td><td>���������� ���������</td><td>���.</td><td><input type=textbox name = b17 value='<%=session("b22")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>


<tr>
<td>3.</td><td><b>������������������ ��������� ������� (������������) ��</b></td><td>���.</td><td><input type=textbox name=b18 value='<%=session("b23")%>'></td><td><input type=checkbox name = k7 value='<%=session("b24")%>'></td><td>&nbsp;</td>
</tr>
<tr>
<td>3.1</td><td>������� ���������</td><td>���</td><td><input type=textbox name=b19 value='<%=session("b25")%>'></td><td><input type=checkbox name = k8 value='<%=session("b26")%>'></td><td>
<select name="b20">
<option value="<%=session("b27")%>" selected><%=session("b27")%></option>
<option value="1" >�������</option>
<option value="2" >�����</option>
<option value="3" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>3.2</td><td>������� ���������</td><td>���</td><td><input type=textbox name=b21 value='<%=session("b28")%>'></td><td><input type=checkbox name = k9 value='<%=session("b29")%>'></td><td>
<select name="b22">
<option value="<%=session("b30")%>" selected><%=session("b30")%></option>
<option value="1" >�������</option>
<option value="2" >�����</option>
<option value="3" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>3.3</td><td>������� ���������</td><td>���</td><td><input type=textbox name=b23 value='<%=session("b31")%>'></td><td><input type=checkbox name = k10 value='<%=session("b32")%>'></td><td>
<select name="b24">
<option value="<%=session("b33")%>" selected><%=session("b33")%></option>
<option value="1" >�������</option>
<option value="2" >�����</option>
<option value="3" >��� � ���</option>
</select>
</td>
</tr>
<tr>
<td>3.4</td><td>������� ���������</td><td>���</td><td><input type=textbox name=b25 value='<%=session("b34")%>'></td><td><input type=checkbox name = k11 value='<%=session("b35")%>'></td><td>
<select name="b26">
<option value="<%=session("b36")%>" selected><%=session("b36")%></option>
<option value="1" >�������</option>
<option value="2" >�����</option>
<option value="3" >��� � ���</option>
</select>
</td>
</tr>

<tr>
<td>4.</td><td><b>������� ����� ��� ������� ��������� �������,</b> �������������� ���������� � ������� ��� �������� ��������� ���������� � ���������, � ����� ���������� ������� �������, ������������ � ���������</td><td>��/���</td><td><select name="b27">
<option value="<%=session("b37")%>" selected><%=session("b37")%></option>
<option value="1" >��</option>
<option value="2" >���</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>5.</td><td><b>����������� � ������� ������, ������ ����� (������� � ���-�� ���)</b></td><td>��/���</td><td><select name="b28">
<option value="<%=session("b38")%>" selected><%=session("b38")%></option>
<option value="1" >��</option>
<option value="2" >���</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>5.1</td><td>�������</td><td>��.�</td><td><input type=textbox name=b29  value='<%=session("b39")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td></td><td>���������� ������� ������� � ������� �����</td><td>���-��</td><td><input type=textbox name=b30 value='<%=session("b40")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>6.</td><td><b>�������� ������ ���������� ����� </b>(A - �� �������������� ���� �����/ � - �� ��������)</td><td>�/�</td><td><select name="b31">
<option value="<%=session("b41")%>" selected><%=session("b41")%></option>
<option value="1" >�</option>
<option value="2" >�</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.</td><td><b>����� ����� � ������ ������ (����� ������ �� ����)</b></td><td>���.�.</td><td><input type=textbox name=b32 value='<%=session("b42")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>

<tr>
<td>8.</td><td><b>����� ���, ��������� �������� ����� ������</b></td><td>���.�.</td><td><input type=textbox name=b33 value='<%=session("b43")%>'></td><td>&nbsp;</td><td>&nbsp;</td>
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
<td>1.1.1</td><td>���������� �������������� ���������</td><td>���.</td><td><input type=text name = c5 value='<%=session("c5")%>'></td>
<td><input type=checkbox name = c6 value='<%=session("c6")%>'></td>
</tr>
<tr>
<td>1.2</td><td>����������� � ������ ����������, ����� � �����</td>
<td>��</td><td><input type=text name = c7 value='<%=session("c7")%>'></td>
<td><input type=checkbox name = c8 value='<%=session("c8")%>'></td>
</tr>
<tr>
<td><b>2.</b></td><td><b>�������� ���������� ������ �������� ��������</b></td>
<td>��/���</td><td><select name="c9">
<option value="<%=session("c9")%>" selected><%=session("c9")%></option>
<option value="��" >��</option>
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
<option value="��" >��</option>
<option value="���" >���</option>
</select></td><td><input type=checkbox name = c15 value='<%=session("c15")%>'></td>
</tr>
<tr>
<td>3.1</td><td>������� ���������</td> <td>��.�</td><td><input type=text name = c16 value='<%=session("c16")%>'></td>
</tr>
<tr>
<td>3.2</td><td>������������� ��������� ����</td> <td>&nbsp;</td><td><input type=text name = c16_1  value='<%=session("c17")%>'></td><td>&nbsp;</td>
</tr>
<tr>
<td><b>4.</b></td><td><b>����� ���� (������� ���� � ����� �������)</b></td> <td>��/���</td><td><select name="c17">
<option value="<%=session("c18")%>" selected><%=session("c18")%></option>
<option value="��" >��</option>
<option value="���" >���</option>
</select></td><td><input type=checkbox name = c18 value='<%=session("c19")%>'></td>
</tr>
<tr>
<td>4.1</td><td>������ ������ � ����� (� ����)</td> <td>��.�</td><td><input type=text name = c19 value='<%=session("c20")%>'></td><td><input type=checkbox name = c20 value='<%=session("c21")%>'></td>
</tr>
<tr>
<td>4.2</td><td>������������ (�� ���������)</td> <td>��.�</td><td><input type=text name = c21 value='<%=session("c22")%>'></td><td><input type=checkbox name = c22 value='<%=session("c23")%>'></td>
</tr>
<tr>
<td>4.3</td><td> � ������� ������������ ����������� </td> <td>��.�</td>
<td><input type=text name = c23 value='<%=session("c24")%>'></td><td>
<input type=checkbox name = c24 value='<%=session("c25")%>'></td>
</tr>
<tr>
<td><b>5.</b></td><td><b>������ ������, ������� �� ������ �� ��������</b></td>
<td>&nbsp;</td> <td>&nbsp;</td>
<td><input type=checkbox name = c25 value='<%=session("c26")%>'></td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
<td><input type=text name = c26 value='<%=session("c27")%>'></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
<td><input type=text name = c27 value='<%=session("c28")%>'></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
<td><input type=text name = c28 value='<%=session("c29")%>'></td><td>&nbsp;</td>
</tr>
----->
<tr>
<td></td><td></td><td>&nbsp;</td><td>&nbsp;</td><td><input type=submit value="���������"></td>
</tr>
</table>
</align>
</form>
<center>
<table border=0 >
<td>

</table>
</center>
</tr>
</table>
</center>


</td>
</tr>
		
</table>		
</td>
</tr>
</center>		

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