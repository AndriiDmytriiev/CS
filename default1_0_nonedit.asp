<%
on error resume next

set conn2 = server.CreateObject("Adodb.connection")
      
        connStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & _
         Server.MapPath("missyou.mdb")
         conn2.Open connStr
      ' Recordset object
dim rs1,sql

Set rs1 = Server.CreateObject("ADODB.Recordset")      

sql="select distinct i.NickName,c.ClientName,p.PersonName,c.Address,s.App1,s.App2,s.App3,s.App4,s.App5,s.App6,i.datebeg,i.dateend,i.ID from StatusHistory s,InternClean i,Persons p, Client c  where c.clientID = p.clientID and p.Nick='" & Session("login") & "' and p.ID=i.SaleID and i.ID=" & cint(Session("ID")) & ""
set rs1=conn2.execute(sql)
Dim ID
    Dim rs7
      ID = Request("ID")
      if  (Len(ID) <> 0) then
	  Session("ID")=ID
      end if

   If Len(ID) > 0 or Session("ID")<>"" Then

        Set rs = Server.CreateObject("ADODB.Recordset")
        sql1="select [SaleID],[ClientID],[datebeg],[dateend],[status],[a1],[a2],[a3],[a4],[a5],[a6],[a7],[a8],[a9],[a10],[a11],[a12],[a13],[a14],[a15]," 
sql1= sql1 &"[a16],[a17],[a18],[a19],[a20],[a21],[a22],[a23],[a24],[a25],[a26],[a27],[a28],[a29],[a30],[a31],[a32],[a33],[a34],[a35],[a36],[a37],[a38],[a39],[a40],[a41],[a42],[a43]"
sql1= sql1 & ",[a44],[a45],[a46],[a47],[a48],[a49],[a50],[a51],[a52],[a53],[a54],[a55],[a56],[a57],[a58],[a59],[a60],[a61],[a62],[a63],[a64],[a65],[a66],[a67],[a68],[a69],[a70],[a71],[a72],[a73]"
sql1= sql1 & ",[a74],[a75],[a76],[a77],[a78],[a79],[a80],[a81],[a82],[a83],[a84],[a85],[a86],[a87],[a88],[a89],[a90],[a91],[a92],[a93],[a94],[a95],[a96],[a97],[a98],[a99],[a100]"
sql1= sql1 & ",[a101],[a102],[a103],[a104],[a105],[a106],[a107],[a108],[a109],[a110],[a111],[a112],[a113],[a114],[a115],[a116],[a117],[a118],"
sql1= sql1 & "[b1],[b2],[b3],[b4],[b5],[b6],[b7],[b8],[b9],[b10],[b11],[b12],[b13],[b14],[b15],[b16],[b17],[b18],[b19],[b20],[b21],[b22],[b23],[b24],[b25],[b26],[b27],[b28],[b29],[b30],[b31],[b32],[b33]"
sql1= sql1 & ",[b34],[b35],[b36],[b37],[b38],[b39],[b40],[b41],[b42],[b43],[b44]"
sql1= sql1 & ",[c1],[c2],[c3],[c4],[c5],[c6],[c7],[c8],[c9],[c10],[c11],[c12],[c13],[c14],[c15],[c16],[c17],[c18],[c19],[c20],[c21],[c22],[c23],[c24],[c25],[c26],[c27],[c28],[c29],[c30],[c31],[c32],[c33] from InternClean " 
sql1= sql1 & " where ID = " & cint(Session("ID")) & ""
       
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



<form method=get >

<%Session("pr")=rs1(0)%>
<table>
<tr>
<td><font size=3>�������� �������: <%=Session("pr")%></font></td>
</tr>

<tr>
<td><font size=3>���������: <%=Session("login")%></font></td>
</tr>

<tr>
 <table>
  
 <td valign="top">
		<a href="common.asp?ID=<%=Session("ID")%>"><font  size=1>����� ������(�������������)</font></a>
		<hr>
		<a href="vnutr.asp?ID=<%=Session("ID")%>"><font  size=1>���������� ������ ��������� (�������������)</font></a>
		<hr>
		<a href="terr.asp?ID=<%=Session("ID")%>"><font  size=1>������ ���������� (�������������)</font></a>
		<hr>
		<a href="spec.asp?ID=<%=Session("ID")%>"><font  size=1>�������������� ������ (�������������)</font></a>
		<hr>
		<a href="uploadform.asp"><font  size=1>��������� ����</font></a>
		<hr>
		
 </td>
<td>



<table border=0 width=100%>
<!----
<tr><td align=left>��������: ���������� �������� ������, ���������� ������. ������, ���������� � ���� ������������ ������ ���������������!
</td></tr>
-------->
<tr><td align=left>1.<b><u>���������� ������ ���������:</u></b></u></b></td></tr>
</table>
<br>
<!----

������������, ����������, ���������� �� ���� �������, � ������� ���������� ����������� ������.
-------->


<table border=1>


<tr>
<td>&nbsp;</td><td>����������</td><td>��.���.</td><td>������ � 1</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>1.</td><td><b>������� ������ � ������/�������</b></td><td></td><td>�����</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<% if session("a2")="on" then %>

<tr>
<td>1.1</td><td>���������</td><td>��.</td>
<td><%=session("a1")%></td><td>
</td>
<td>
<%=session("a3")%>
</td>
</tr>
<% end if %>

<% if session("a5") ="on" then %>

<tr>
<td>1.2</td><td>��� ���������� �������</td><td>���</td>
<td><%=session("a4")%></td><td>
</td>
<td>
<%=session("a6")%>
</td>
</tr>
<% end if %>

<% if session("a8") ="on" then %>

<tr>
<td>1.3</td><td>������ ��������</td><td>�.</td><td>
<%=session("a7")%></td><td>
</td>
<td>
<%=session("a9")%>
</td>
</tr>
<% end if %>

<% if session("a11") ="on" then %>

<tr>
<td>2.</td><td><b>����� �������</b></td><td>��.�</td><td>
<%=session("a10")%></td><td>
</td>
<td>
<%=session("a12")%>
</td>
</tr>
<% end if %>

<% if session("a14") ="on" then %>
<tr>
<td>2.1</td><td>������� ������� �����</td><td>��.�</td><td>
<%=session("a13")%></td><td>
</td>
<td>
<%=session("a15")%>
</td>
</tr>
<% end if %>
</td>
<% if session("a17") ="on" then %>
<tr>
<td>2.2</td><td>�������� VIP</td><td>���-��</td>
<td>
<%=session("a16")%>
</td>
<td>&nbsp;</td>
<td>
<%=session("a18")%>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td>
<td>
<%=session("a19")%>
</td><td>&nbsp;</td>
<td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a21") ="on" then %>

<tr>
<td>2.3</td><td>������� ���������</td><td>���-��</td><td>
<%=session("a20")%></td><td>
</td>
<td>
<%=session("a22")%>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<%=session("a23")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a25") ="on" then %>

<tr>
<td>2.4</td><td>��������� ���������</td><td>���-��</td>
<td>
<%=session("a24")%>
</td>
<td>
</td><td>
<%=session("a26")%>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td>
<td>
<%=session("a27")%>
</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a29") ="on" then %>

<tr>
<td>2.5</td><td>��������� ���������</td><td>���-��</td><td>
<%=session("a28")%></td>
<td>
</td>
<td>
<%=session("a30")%>
</td>
</tr>

<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<%=session("a31")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a33") ="on" then %>

<tr>
<td>2.6</td><td>����������� ���������, �������</td><td>���-��</td><td>
<%=session("a32")%></td><td>
</td>
<td>
<%=session("a34")%>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<%=session("a35")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a37") ="on" then %>

<tr>
<td>2.7</td><td>��������</td><td>���-��</td><td>
<%=session("a36")%></td><td>
</td>
<td>
<%=session("a38")%>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<%=session("a39")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a41") ="on" then %>

<tr>
<td>2.8</td><td>��������</td><td>���-��</td><td>
<%=session("a40")%></td><td>
</td>
<td>
<%=session("a42")%>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<%=session("a43")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a45") ="on" then %>

<tr>
<td>2.9</td><td>�����</td><td>���-��</td><td>
<%=session("a44")%></td><td>
</td>
<td>
<%=session("a46")%>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<%=session("a47")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a49") ="on" then %>

<tr>
<td>2.10</td><td>����������</td><td>���-��</td><td>
<%=session("a48")%></td><td>
</td>
<td>
<%=session("a50")%>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<%=session("a51")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a53") ="on" then %>

<tr>
<td>2.11</td><td>�������</td><td>���-��</td><td>
<%=session("a52")%></td><td>
</td>
<td>
<%=session("a54")%>
</td>
</tr>

<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<%=session("a55")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a57") ="on" then %>

<tr>
<td>2.12</td><td>�����, ���������</td><td>���-��</td><td>
<%=session("a56")%></td><td>
</td><td>
<%=session("a58")%>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<%=session("a59")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.13</td><td>������ ������� (�� ����.������������)</td><td></td><td>
<%=session("a60")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a62") ="on" then %>

<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>���-��</td><td>
<%=session("a61")%></td><td>
</td>
<td>
<%=session("a63")%>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>�������, ��.�</td><td>
<%=session("a64")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a66") ="on" then %>

<tr>
<td>3.</td><td><b>�-�� ����������� ����� �������� � �����������, ���. � ������� � ����</b></td><td>���.</td><td>
<%=session("a65")%></td><td>
</td><td>&nbsp;</td>
</tr>
<% end if %>


<% if (session("a69") ="on") or (session("a71") ="on") or (session("a73") ="on") or (session("a75") ="on") then %> 




<tr>
<td>4.</td><td><b>������������� � ����������� �/� ���������� ����������� (��������� ������ � �����)</b></td><td>���-��</td><td>&nbsp;</td><td>&nbsp;</td><td>
<%=session("a67")%>
</td>
</tr>

<% if session("a69") ="on" then %>

<tr>
<td>&nbsp;</td><td>��������� ������</td><td>���./���.</td><td>
<%=session("a68")%></td><td>
</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a71") ="on" then %>
<tr>
<td>&nbsp;</td><td>������ ����</td><td>����/���.</td><td>
<%=session("a70")%></td><td>
</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a73") ="on" then %>
<tr>
<td>&nbsp;</td><td>��� ���������</td><td>����/���.</td><td>
<%=session("a72")%></td><td>
</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a75") ="on" then %>
<tr>
<td>&nbsp;</td><td>�������� ������� �/�������</td><td>��./���.</td><td>
<%=session("a74")%></td><td>
</td><td>&nbsp;</td>
</tr>
<% end if %>

<% end if %> <!---end if of the 19 --->

<% if session("a76") ="on" then %>
<tr>
<td>5.</td><td><b>�����������</b></td><td>���.</td><td>&nbsp;</td><td>
</td><td>
<%=session("a77")%>
</td>
</tr>
<% end if %>

<% if session("a79") ="on" then %>
<tr>
<td>5.1</td><td>������ �������� (��������)</td><td>��.�</td><td>
<%=session("a78")%></td><td>
</td><td>&nbsp;</td>
</tr>
<% end if %>
<% if session("a81") ="on" then %>
<tr>
<td>5.2</td><td>����������� �������� (��������, ������, ���������, �������� ���, �������)</td><td>��.�</td><td>
<%=session("a80")%></td><td>
</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a83") ="on" then %>
<tr>
<td>5.3</td><td>������� �������� (������, ������, ������) ���/�����</td><td>��.�</td><td>
<%=session("a82")%></td><td>
</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a85") ="on" then %>
<tr>
<td>5.4</td><td>���������� �����������</td><td>��.�</td><td>
<%=session("a84")%></td><td>
</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a87") ="on" then %>
<tr>
<td>5.5</td><td>������������� �����������</td><td>��.�</td><td>
<%=session("a86")%></td><td>
</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a89") ="on" then %>
<tr>
<td>5.6</td><td>������� �����������</td><td>��.</td><td>
<%=session("a88")%></td><td>
</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a91") ="on" then %>
<tr>
<td>5.7</td><td>������� �����</td><td>��.</td><td>
<%=session("a90")%></td><td>
</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a93") ="on" then %>
<tr>
<td>5.8</td><td>������� ������</td><td>��.</td><td>
<%=session("a92")%></td><td>
</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a95") ="on" then %>
<tr>
<td>5.9</td><td>����������� ������</td><td>��.</td><td>
<%=session("a94")%></td><td>
</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a97") ="on" then %>
<tr>
<td>5.10</td><td>���������� ������</td><td>��.</td><td>
<%=session("a96")%></td><td>
</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a99") ="on" then %>
<tr>
<td>5.11</td><td>������ �����������</td><td>��./��.�</td><td>
<%=session("a98")%></td><td>
</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a100") ="on" then %>
<tr>
<td>6.</td><td><b>������������ � ��� ������</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
</td><td>
<%=session("a101")%>
</td>
</tr>


<tr>
<td>6.1</td><td>������ ���������� �������� ����������� ������</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>�:</td><td>���</td><td>
<%=session("a102")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>��:</td><td>���</td><td>
<%=session("a103")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>���� � ����:</td><td>���</td><td>
<%=session("a104")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>6.2</td><td>������ ���������� �������������� ������</td><td></td><td></td><td></td><td></td>
</tr>

<tr>
<td>&nbsp;</td><td>�:</td><td>���</td><td>
<%=session("a105")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>��:</td><td>���</td><td>
<%=session("a106")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>���� � ����:</td><td>���</td><td>
<%=session("a107")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>



<tr>
<td>6.3</td><td>���������� ���������</td><td>���.</td><td>
<%=session("a108")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a109") ="on" then %>
<tr>
<td>7.</td><td><b>��������� ��� ���������� ����������������� ��������� � ������������</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
</td><td>
<%=session("a110")%>
</td>
</tr>

<tr>
<td>7.1</td><td>������� ���������</td><td>��/���</td><td><%=session("a111")%>
</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.2</td><td>������� �������� � ���������</td><td>��/���</td>
<td>
<%=session("a112")%>
</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.3</td><td>������� ���������</td><td>��.�</td><td>
<%=session("a113")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("a114") ="on" then %>
<tr>
<td>8.</td><td><b>��������� ��� ���������� ��������������� ��������� � ������������������ ������� ���� (��� �������� ����� 10000 ��.�)</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
</td><td>
<%=session("a115")%>
</td>
</tr>

<tr>
<td>8.1</td><td>������� ���������</td><td>��/���</td><td>
<%=session("a116")%>
</td><td></td><td></td>
</tr>
<tr>
<td>8.2</td><td>������� ���������</td><td>��.�</td><td>
<%=session("a117")%></td><td></td><td></td>
</tr>
<tr>
<td>8.3</td><td>���-�� ������������������ ������� ����</td><td>��.</td><td>
<%=session("a118")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% end if %>


</table>
<br>

<table border=0 width=100%>
<tr><td align=left>��������: ���������� �������� ������, ���������� ������. ������, ���������� � ���� ������������ ������ ���������������!
</td></tr>
<tr><td align=left>2.<b><u>������ ����������:</u></b></td></tr>
</table>


<br>
<!----
������������, ����������, ���������� �� ���� �������, � ������� ���������� ����������� ������.
-------->
<br>
<table border=1>
<tr>
<td>&nbsp;</td><td>����������</td><td>��.���.</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<% if session("b3") ="on" then %>

<tr>
<td>1.</td><td><b>����� ��������� ����������</b></td><td>��.�</td><td> <%=session("b1")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>1.1</td><td>������� � ��������� �������� (��������, �������, ������, �������, ������)</td>
<td>&nbsp;</td><td><%=session("b2")%></td>
<td></td><td>
<%=session("b4")%>
</td>
</tr>
<% end if %>

<% if session("b6") ="on" then %>
<tr>
<td>1.1.1</td><td>��������</td><td>���</td>
<td><%=session("b5")%></td><td></td><td>
<%=session("b7")%>
</td>
</tr>
<% end if %>

<% if session("b9") ="on" then %>
<tr>
<td>1.1.2</td><td>�������</td><td>�.</td><td><%=session("b8")%></td>
<td></td><td>
<%=session("b10")%>
</td>
</tr>
<% end if %>

<% if session("b12") ="on" then %>
<tr>
<td>1.1.3</td><td>������</td><td>�.</td><td><%=session("b11")%></td><td></td><td>
<%=session("b13")%>
</td>
</tr>
<% end if %>

<% if session("b15") ="on" then %>
<tr>
<td>1.1.4</td><td>�������</td><td>�.</td><td><%=session("b14")%></td>
<td></td><td>
<%=session("b16")%>
</td>
</tr>
<% end if %>

<% if session("b18") ="on" then %>
<tr>
<td>2.</td><td><b>������������ � ��� ������ ����������</b></td>
<td>��.�</td><td><%=session("b17")%></td><td></td><td>
<%=session("b19")%>
</td>
</tr>


<tr>
<td>2.1</td><td>������ ������</td><td></td><td></td><td></td><td></td>
</tr>

<tr>
<td>&nbsp;</td><td>�:</td><td>���</td><td><%=session("b20")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>��:</td><td>���</td><td><%=session("b21")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>���� � ����:</td><td>���</td><td><%=session("b22")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>


<tr>
<td>2.2</td><td>���������� ���������</td><td>���.</td><td><%=session("b23")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% end if %>

<%if session("b25")="on" then%>
<tr>
<td>3.</td><td><b>������������������ ��������� ������� (������������) ��</b></td><td>���.</td><td><%=session("b24")%></td><td></td><td>&nbsp;</td>
</tr>
<%end if%>

<% if session("b27") ="on" then %>

<tr>
<td>3.1</td><td>������� ���������</td><td>���</td><td><%=session("b26")%></td><td></td><td>
<%=session("b28")%>
</td>
</tr>
<% end if %>

<% if session("b30") ="on" then %>
<tr>
<td>3.2</td><td>������� ���������</td><td>���</td><td><%=session("b29")%></td><td></td><td>
<%=session("b31")%>
</td>
</tr>
<% end if %>

<% if session("b33") ="on" then %>
<tr>
<td>3.3</td><td>������� ���������</td><td>���</td><td><%=session("b32")%></td><td></td><td>
<%=session("b34")%>
</td>
</tr>
<% end if %>

<% if session("b36") ="on" then %>
<tr>
<td>3.4</td><td>������� ���������</td><td>���</td><td><%=session("b35")%></td><td></td><td>
<%=session("b37")%>
</td>
</tr>
<% end if %>

<tr>
<td>4.</td><td><b>������� ����� ��� ������� ��������� �������,</b> �������������� ���������� � ������� ��� �������� ��������� ���������� � ���������, � ����� ���������� ������� �������, ������������ � ���������</td><td>��/���</td><td><%=session("b38")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>5.</td><td><b>����������� � ������� ������, ������ ����� (������� � ���-�� ���)</b></td><td>��/���</td><td><%=session("b39")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>5.1</td><td>�������</td><td>��.�</td><td><%=session("b40")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td></td><td>���������� ������� ������� � ������� �����</td><td>���-��</td><td><%=session("b41")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>6.</td><td><b>�������� ������ ���������� ����� </b>(A - �� �������������� ���� �����/ � - �� ��������)</td><td>�/�</td><td>
<%=session("b42")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.</td><td><b>����� ����� � ������ ������ (����� ������ �� ����)</b></td><td>���.�.</td><td> <%=session("b43")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>

<tr>
<td>8.</td><td><b>����� ���, ��������� �������� ����� ������</b></td><td>���.�.</td><td><%=session("b44")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

</table>

<align=left>
<br>
<table border=0 width=100%>
<!----
<tr><td align=left>��������: ���������� �������� ������, ���������� ������. ������, ���������� � ���� ������������ ������ ���������������!</td></tr>
---->
<tr><td align=left>3.<b><u>�������������� ������:</u></b></td></tr>
</table>
<table border=1 width=100%>
<tr>
<td>&nbsp;</td><td>����������</td><td>��.���.</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<% if session("c2") ="on" then %>
<tr>
<td><b>1.</b></td><td><b>����������� ������ ���������</b></td><td>��/���</td><td><%=session("c1")%></td><td></td>
</tr>
<% end if %>

<% if session("c4") ="on" then %>
<tr>
<td>1.1</td><td>������� ���������-���������</td><td>��/���</td><td><%=session("c3")%>
</td><td></td>
</tr>
<% end if %>

<% if session("c6") ="on" then %>

<tr>
<td>1.1.1</td><td>���������� �������������� ���������</td><td>���.</td><td><%=session("c5")%></td><td></td>
</tr>
<% end if %>

<% if session("c8") ="on" then %>

<tr>
<td>1.2</td><td>����������� � ������ ����������, ����� � �����</td><td>��</td><td><%=session("c7")%></td><td></td>
</tr>
<% end if %>

<% if session("c10") ="on" then %>
<tr>
<td><b>2.</b></td><td><b>�������� ���������� ������ �������� ��������</b></td><td>��/���</td><td><%=session("c9")%></td><td></td>
</tr>
<% end if %>

<% if session("c12") ="on" then %>
<tr>
<td>2.1</td><td>������� �������� ��������</td><td>��.�</td><td><%=session("c11")%></td><td></td>
</tr>

<tr>
<td>2.2</td><td>������������� ����������</td><td>&nbsp;</td><td><%=session("c13")%></td> <td>&nbsp;</td>
</tr>
<% end if %>

<% if session("c15") ="on" then %>
<tr>
<td><b>3.</b></td><td><b>��������� ����������� ���� �� ����������� ����</b></td><td>��/���</td><td><%=session("c14")%></td><td></td>
</tr>

<tr>
<td>3.1</td><td>������� ���������</td> <td>��.�</td><td><%=session("c16")%></td>
</tr>
<tr>
<td>3.2</td><td>������������� ��������� ����</td> <td>&nbsp;</td><td><%=session("c17")%></td><td>&nbsp;</td>
</tr>
<% end if %>

<% if session("c19") ="on" then %>
<tr>
<td><b>4.</b></td><td><b>����� ���� (������� ���� � ����� �������)</b></td> <td>��/���</td><td><%=session("c18")%></td><td></td>
</tr>
<% end if %>

<% if session("c21") ="on" then %>
<tr>
<td>4.1</td><td>������ ������ � ����� (� ����)</td> <td>��.�</td><td><%=session("c20")%></td><td></td>
</tr>
<% end if %>

<% if session("c23") ="on" then %>
<tr>
<td>4.2</td><td>������������ (�� ���������)</td> <td>��.�</td><td><%=session("c22")%></td><td></td>
</tr>
<% end if %>

<% if session("c25") ="on" then %>
<tr>
<td>4.3</td><td> � ������� ������������ ����������� </td> <td>��.�</td><td><%=session("c24")%></td><td></td>
</tr>
<% end if %>

<% if session("c26") ="on" then %>
<tr>
<td><b>5.</b></td><td><b>������ ������, ������� �� ������ �� ��������</b></td><td>&nbsp;</td> <td>&nbsp;</td><td></td>
</tr>

<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td><%=session("c27")%></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td><%=session("c28")%></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td><%=session("c29")%></td><td>&nbsp;</td>
</tr>
<% end if %>
<tr>
<td></td><td></td><td>&nbsp;</td><td>&nbsp;</td><td></td>
</tr>
</table>
</td>

</align>
</form>
<center>
<table border=0 >
<td>
<tr>		
		
		<td>
		<a href="approve.asp"><font  size=3>�������</font></a>
		<hr>
		</td>
</tr>
</table>
</center>


</td>
</tr>
		
</table>		
</td>
</tr>
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