<%
on error resume next
set conn2 = server.CreateObject("Adodb.connection")
    
        connStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & _
         Server.MapPath("missyou.mdb")
         conn2.Open connStr
      ' Recordset object
      
    '    Set rs3 = Server.CreateObject("ADODB.Recordset")
    '    sql="select Login,Password from Logins where Login='" &trim(Request.Form("login1"))& "' and " & "Password='"&trim(Request.Form("pass1"))&"'"
       
    '    Set rs3 = conn2.Execute(sql)
        
        'rs3.MoveFirst
        
    '    if not rs3.EOF  then
'			Session("Login")=trim(Request.Form("login1"))
'			Session("Password")=trim(Request.Form("pass1"))
'		else
'			Session("Login")=""
'			Session("Password")=""
 '       end if
        Session("Client")=Request.Form("client")
         
        Session("contact")=Request.Form("contact")

dim rs4
  Set rs4 = Server.CreateObject("ADODB.Recordset")
        sql="select ClientID,ClientName from Client"
       
        Set rs4 = conn2.Execute(sql)
        
        rs4.MoveFirst

dim rs5
  Set rs5 = Server.CreateObject("ADODB.Recordset")
        sql="select ID,PersonName from Persons"
       
        Set rs5 = conn2.Execute(sql)
        
        rs5.MoveFirst



        
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
��������: ���������� �������� ������, ���������� ������. ������, ���������� � ���� ������������ ������ ���������������!
<br>
1.<b><u>���������� ������ ���������:</u></b>

������������, ����������, ���������� �� ���� �������, � ������� ���������� ����������� ������.

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
<%IF Session("Login")="" then%> 
<FONT size="3" color="#0000ff" style="FONT-WEIGHT: bold">

<table><td><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td><font color=#0000ff>������� ����� � ������, ����� �� ������� �������� ���������.<br>���� �� ����� � ������ ���, ����� �������� �� ��������" </font><a href=reg.asp>&nbsp;&nbsp;&nbsp;�����������</a></td></table>
		<tr><td  valign=top>
<table>
<tr>
<form method=post action=default1.asp id=form1 name=form1>
<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td><b>�����:&nbsp;</b> &nbsp;</td></tr>
<tr>
<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td><input type="login" id=login1 name=login1></td>
</tr>
<tr>
<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td><b>������:</b></td>
</tr>
<tr>
<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td>
<input type="password" id=pass1 name=pass1></td>
</tr>
<!-------
<tr>
<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td>�������:  </td>
</tr>
<tr>
<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td><b><%=date()%></b></td>
</tr>
<tr>
<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td><font color=red><b>���� ������:</b></font></td>
</tr>
<tr>
<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>

<td><select name=day  value="<%=Session("day")%>">
	 	 	 
	 	 	 <%if   Session("day")<>"" then%>
	 	 	 <option value="<%=Session("day")%>" selected><%=Session("day")%></option>
	 	 	 <%else%>
		  <option value="1" selected>1</option>
		  <%end if%>
		 <%for i=1 to 31%>
		 <option value="<%=i%>"><%=i%></option>
		 <%next%>
                 </select>	 	 	 
	 	 	 
	 	 <select name=month  value="<%=Session("month")%>">
	         <%if   Session("year")<>"" then%>
	 	 	 <option value="<%=Session("month")%>" selected><%=Session("monthname")%></option>
	 	 	 <%else%>
		 <option value="1">������</option>
		 <option value="2">�������</option>
		 <option value="3">����</option>
		 <option value="4">������</option>
		 <option value="5">���</option>
		 <option value="6">����</option>
		 <option value="7">����</option>
		 <option value="8">������</option>
		 <option value="9">��������</option>
		 <option value="10">�������</option>
		 <option value="11">������</option>
		 <option value="12">�������</option>
		  <%end if%>
		 
		 
                </select>
	 <select name=year> 
	 
	     <%if   Session("year")<>"" then%>
		 <option value="<%=Session("year")%>" selected><%=Session("year")%></option>
		 
		 <%else%>
		  <option value="2006" selected>2006</option>
		  <%end if%>
		 <%for i=2006 to 2050%>
		 <option value="<%=i%>"><%=i%></option>
		 <%next%>
       </select>
</td>
</tr>	 
</table>

<table>

<tr>
<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td>
������:
</td>
</tr>
<tr>
<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td>
<select name=client>
<option value="0" selected>��� �������</option>
<% while not rs4.eof %>
<option value="<%=rs4(0)%>" ><%=rs4(1)%></option>
<% rs4.movenext %>
<% wend %>
</select>
</td>
</tr>
<tr>
<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td>
���������:
</td>
</tr>
<tr>
<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td>
<select name=contact id="contact">
<option value="0" selected>��� ���������</option>
<% while not rs5.eof %>
<option value="<%=rs5(0)%>" ><%=rs5(1)%></option>
<% rs5.movenext %>
<% wend %>
</select>
</td>
</tr>
------------>
<tr>
<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td>
<input type="submit" value="�����" id=submit1 name=submit1><br>
</td>
</tr>
</table>
</form>
<%else%>



<form method=get action='create_page1.asp'>

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
<input type=checkbox name = ch1 value="ON"></td><td>
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
<input type=checkbox name = ch2 value="ON"></td><td>
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
<input type=checkbox name = ch3 value="ON"></td><td>
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
<input type=checkbox name = ch4 value="ON"></td><td>
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
<input type=checkbox name = ch5 value="ON"></td><td>
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
<input type=checkbox name = ch6 value="ON"></td><td>
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
<input type=checkbox name = ch7 value="ON"></td><td>
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
<input type=checkbox name = ch8 value="ON"></td><td>
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
<input type=checkbox name = ch9 value="ON"></td><td>
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
<input type=checkbox name = ch10 value="ON"></td><td>
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
<input type=checkbox name = ch11 value="ON"></td><td>
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
<input type=checkbox name = ch12 value="ON"></td><td>
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
<input type=checkbox name = ch13 value="ON"></td><td>
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
<input type=checkbox name = ch14 value="ON"></td><td>
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
<input type=checkbox name = ch15 value="ON"></td><td>
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
<input type=checkbox name = ch16 value="ON"></td><td>
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
<input type=checkbox name = ch17 value="ON"></td><td>
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
<input type=checkbox name = ch18 value="ON"></td><td>&nbsp;</td>
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
<input type=checkbox name = ch19_1 value="ON"></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>������ ����</td><td>����/���.</td><td>
<input type=textbox name=t19_2 size="20"></td><td>
<input type=checkbox name = ch19_2 value="ON"></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>��� ���������</td><td>����/���.</td><td>
<input type=textbox name=t19_3 size="20"></td><td>
<input type=checkbox name = ch19_3 value="ON"></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>�������� ������� �/�������</td><td>��./���.</td><td>
<input type=textbox name=t19_4 size="20"></td><td>
<input type=checkbox name = ch19_4 value="ON"></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.</td><td><b>�����������</b></td><td>���.</td><td>&nbsp;</td><td>
<input type=checkbox name = ch20 value="ON"></td><td>
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
<input type=checkbox name = ch20_1 value="ON"></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.2</td><td>����������� �������� (��������, ������, ���������, �������� ���, �������)</td><td>��.�</td><td>
<input type=textbox name=t20_2 size="20"></td><td>
<input type=checkbox name = ch20_2 value="ON"></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.3</td><td>������� �������� (������, ������, ������) ���/�����</td><td>��.�</td><td>
<input type=textbox name=t20_3 size="20"></td><td>
<input type=checkbox name = ch20_3 value="ON"></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.4</td><td>���������� �����������</td><td>��.�</td><td>
<input type=textbox name=t20_4 size="20"></td><td>
<input type=checkbox name = ch20_4 value="ON"></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.5</td><td>������������� �����������</td><td>��.�</td><td>
<input type=textbox name=t20_5 size="20"></td><td>
<input type=checkbox name = ch20_5 value="ON"></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.6</td><td>������� �����������</td><td>��.</td><td>
<input type=textbox name=t20_6 size="20"></td><td>
<input type=checkbox name = ch20_6 value="ON"></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.7</td><td>������� �����</td><td>��.</td><td>
<input type=textbox name=t20_7 size="20"></td><td>
<input type=checkbox name = ch20_7 value="ON"></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.8</td><td>������� ������</td><td>��.</td><td>
<input type=textbox name=t20_8 size="20"></td><td>
<input type=checkbox name = ch20_8 value="ON"></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.9</td><td>����������� ������</td><td>��.</td><td>
<input type=textbox name=t20_9 size="20"></td><td>
<input type=checkbox name = ch20_9 value="ON"></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.10</td><td>���������� ������</td><td>��.</td><td>
<input type=textbox name=t20_10 size="20"></td><td>
<input type=checkbox name = ch20_10 value="ON"></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.11</td><td>������ �����������</td><td>��./��.�</td><td>
<input type=textbox name=t20_11 size="20"></td><td>
<input type=checkbox name = ch20_11 value="ON"></td><td>&nbsp;</td>
</tr>

<tr>
<td>6.</td><td><b>������������ � ��� ������</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
<input type=checkbox name = ch21 value="ON"></td><td>
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
<input type=checkbox name = ch22 value="ON"></td><td>
<select name="Period22">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>

<tr>
<td>7.1</td><td>������� ���������</td><td>��/���</td><td><select name="choice1">
<option value="1" selected>��</option>
<option value="2" >���</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.2</td><td>������� �������� � ���������</td><td>��/���</td><td><select name="choice2">
<option value="1" selected>��</option>
<option value="2" >���</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.3</td><td>������� ���������</td><td>��.�</td><td>
<input type=textbox name = t29 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>8.</td><td><b>��������� ��� ���������� ��������������� ��������� � ������������������ ������� ���� (��� �������� ����� 10000 ��.�)</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
<input type=checkbox name = ch23 value="ON"></td><td>
<select name="Period23">
<option value="�������" selected>�������</option>
<option value="�����" >�����</option>
<option value="��� � ���" >��� � ���</option>
</select>
</td>
</tr>

<tr>
<td>8.1</td><td>������� ���������</td><td>��/���</td><td><select name="choice3">
<option value="1" selected>��</option>
<option value="2" >���</option>
</select></td><td></td><td></td>
</tr>
<tr>
<td>8.2</td><td>������� ���������</td><td>��.�</td><td>
<input type=textbox name = t30 size="20"></td><td></td><td></td>
</tr>
<tr>
<td>8.3</td><td>���-�� ������������������ ������� ����</td><td>��.</td><td><input type=textbox name = t31 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td></td><td></td><td></td><td></td><td></td><td><input type=submit value="���������"></td>
</tr>


</table>
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
<%end if%>

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
<%Response.write (err.description) %>
  </BODY>
</HTML>