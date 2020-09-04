<%
   ' -- show.asp --
   ' Generates a list of uploaded files
   Session("PersonID")=""
   on error resume next
   Session("first")=1
   if Request.QueryString<>"" and Session("qs")="" then
   Session("qs")=Request.QueryString
   end if
   'if len(Request.QueryString)<30 then
   '		 Session("qs")=Request.QueryString
   '  end if
   Response.Buffer = True
   
   ' Connection String
      Dim connStr
      dim conn 
      set conn = server.CreateObject("ADODB.Connection")
      
     'connStr = "DRIVER=Microsoft Access Driver (*.mdb);DBQ="
     'connStr = connStr & Server.MapPath("/andy26/database/missyou.mdb")
      
      connStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & _
      Server.MapPath("missyou.mdb")
      conn.Open connStr
      dim rs
      set rs = server.CreateObject("ADODB.Recordset")
' Add SalesID and clientID

dim a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15,a16,a17,a18,a19,a20,a21,a22,a23,a24
dim a25,a26,a27,a28,a29,a30,a31,a32,a33,a34,a35,a36,a37,a38,a39,a40,a41,a42,a43,a44,a45,a46,a47
dim a48,a49,a50,a51,a52,a53,a54,a55,a56,a57,a58,a59,a60,a61,a62,a63,a64,a65,a66,a67,a68,a69,a70
dim a71,a72,a73,a74,a75,a76,a77
dim a78,a79,a80,a81,a82,a83,a84,a85,a86,a87,a88,a89,a90,a91,a92,a93,a94,a95,a96,a97,a98,a99,a100,a101,a102,a103,a104,a105,a106,a107
dim a108,a109,a110,a111,a112,a113,a114,a115,a116,a117,a118


a1=session("a1")'request.querystring("t1")
a2=session("a2")'request.querystring("Period1")
a3=session("a3")'request.querystring("t2")
a4=session("a4")'request.querystring("Period2")
a5=session("a5")'request.querystring("t3")
a6=session("a6")'request.querystring("Period3")
a7=session("a7")'request.querystring("t4")
a8=session("a8")'request.querystring("Period4")
a9=session("a9")'request.querystring("t5")
a10=session("a10")'request.querystring("Period5")
a11=session("a11")'request.querystring("t6_1")
a12=session("a12")'request.querystring("Period6")
a13=session("a13")'request.querystring("t6_2")
a14=session("a14")'request.querystring("t7_1")
a15=session("a15")'request.querystring("t7_2")
a16=session("a16")'request.querystring("t8_1")
a17=session("a17")'request.querystring("t8_2")
a18=session("a18")'request.querystring("t9_1")
a19=session("a19")'request.querystring("Period9")
a20=session("a20")'request.querystring("t9_2")
a21=session("a21")'request.querystring("t10_1")
a22=session("a22")'request.querystring("Period10")
a23=session("a23")'request.querystring("t10_2")
a24=session("a24")'request.querystring("t11_1")
a25=session("a25")'request.querystring("Period11")
a26=session("a26")'request.querystring("t11_2")
a27=session("a27")'request.querystring("t12_1")
a28=session("a28")'request.querystring("Period12")
a29=session("a29")'request.querystring("t12_2")
a30=session("a30")'request.querystring("t13_1")
a31=session("a31")'request.querystring("t13_2")
a32=session("a32")'request.querystring("t14_1")
a33=session("a33")'request.querystring("Period14")
a34=session("a34")'request.querystring("t14_2")
a35=session("a35")'request.querystring("t15_1")
a36=session("a36")'request.querystring("Period15")
a37=session("a37")'request.querystring("t15_2")
a38=session("a38")'request.querystring("t16_1")
a39=session("a39")'request.querystring("Period16")
a40=session("a40")'request.querystring("t16_2")
a41=session("a41")'request.querystring("t17_0")
a42=session("a42")'request.querystring("t17_1")
a43=session("a43")'request.querystring("t17_2")
a44=session("a44")'request.querystring("t18")
a45=session("a45")'request.querystring("Period19")
a46=session("a46")'request.querystring("t19_1")
a47=session("a47")'request.querystring("t19_2")
a48=session("a48")'request.querystring("t19_3")
a49=session("a49")'request.querystring("t19_4")
a50=session("a50")'request.querystring("Period20")
a51=session("a51")'request.querystring("t20_1")
a52=session("a52")'request.querystring("t20_2")
a53=session("a53")'request.querystring("t20_3")
a54=session("a54")'request.querystring("t20_4")
a55=session("a55")'request.querystring("t20_5")
a56=session("a56")'request.querystring("t20_6")
a57=session("a57")'request.querystring("t20_7")
a58=session("a58")'request.querystring("t20_8")
a59=session("a59")'request.querystring("t20_9")
a60=session("a60")'request.querystring("t20_10")
a61=session("a61")'request.querystring("t20_11")
a62=session("a62")'request.querystring("Period21")
a63=session("a63")'request.querystring("t22")
a64=session("a64")'request.querystring("t23")
a65=session("a65")'request.querystring("t24")
a66=session("a66")'request.querystring("t25")
a67=session("a67")'request.querystring("t26")
a68=session("a68")'request.querystring("t27")
a69=session("a69")'request.querystring("t28")
a70=session("a70")'request.querystring("Period22")
a71=session("a71")'request.querystring("choice1")
a72=session("a72")'request.querystring("choice2")
a73=session("a73")'request.querystring("t29")
a74=session("a74")'request.querystring("Period23")
a75=session("a75")'request.querystring("choice3")
a76=session("a76")'request.querystring("t30")
a77=session("a77")'request.querystring("t31")
a78=session("a78")
a79=session("a79")
a80=session("a80")
a81=session("a81")
a82=session("a82")
a83=session("a83")
a84=session("a84")
a85=session("a85")
a86=session("a86")
a87=session("a87")
a88=session("a88")
a89=session("a89")
a90=session("a90")
a91=session("a91")
a92=session("a92")
a93=session("a93")
a94=session("a94")
a95=session("a95")
a96=session("a96")
a97=session("a97")
a98=session("a98")
a99=session("a99")
a100=session("a100")
a101=session("a101")
a102=session("a102")
a103=session("a103")
a104=session("a104")
a105=session("a105")
a106=session("a106")
a107=session("a107")
a108=session("a108")
a109=session("a109")
a110=session("a110")
a111=session("a111")
a112=session("a112")
a113=session("a113")
a114=session("a114")
a115=session("a115")
a116=session("a116")
a117=session("a117")
a118=session("a118")

dim b1,b2,b3,b4,b5,b6,b7,b8,b9,b10,b11,b12,b13,b14,b15,b16,b17,b18,b19,b20,b21,b22,b23,b24,b25,b26,b27,b28,b29,b30,b31,b32,b33
dim b34,b35,b36,b37,b38,b39,b40,b41,b42,b43,b44
b1=session("b1")
b2=session("b2")
b3=session("b3")
b4=session("b4")
b5=session("b5")
b6=session("b6")
b7=session("b7")
b8=session("b8")
b9=session("b9")
b10=session("b10")
b11=session("b11")
b12=session("b12")
b13=session("b13")
b14=session("b14")
b15=session("b15")
b16=session("b16")
b17=session("b17")
b18=session("b18")
b19=session("b19")
b20=session("b20")
b21=session("b21")
b22=session("b22")
b23=session("b23")
b24=session("b24")
b25=session("b25")
b26=session("b26")
b27=session("b27")
b28=session("b28")
b29=session("b29")
b30=session("b30")
b31=session("b31")
b32=session("b32")
b33=session("b33")
b34=session("b34")
b35=session("b35")
b36=session("b36")
b37=session("b37")
b38=session("b38")
b39=session("b39")
b40=session("b40")
b41=session("b41")
b42=session("b42")
b43=session("b43")
b44=session("b44")


dim c1,c2,c3,c4,c5,c6,c7,c8,c9,c10,c11,c12,c13,c14,c15,c16,c17,c18,c19,c20,c21,c22,c23,c24,c25,c26,c27,c28,c29,c30,c31,c32,c33
c1=request.querystring("c1")
c2=request.querystring("c2")
c3=request.querystring("c3")
c4=request.querystring("c4")
c5=request.querystring("c5")
c6=request.querystring("c6")
c7=request.querystring("c7")
c8=request.querystring("c8")
c9=request.querystring("c9")
c10=request.querystring("c10")
c11=request.querystring("c11")
c12=request.querystring("c12")
c13=request.querystring("c13")
c14=request.querystring("c14")
c15=request.querystring("c15")
c16=request.querystring("c16")
c17=request.querystring("c17")
c18=request.querystring("c18")
c19=request.querystring("c19")
c20=request.querystring("c20")
c21=request.querystring("c21")
c22=request.querystring("c22")
c23=request.querystring("c23")
c24=request.querystring("c24")
c25=request.querystring("c25")
c26=request.querystring("c26")
c27=request.querystring("c27")
c28=request.querystring("c28")
c29=request.querystring("c29")
c30=request.querystring("c30")
c31=request.querystring("c31")
c32=request.querystring("c32")
c33=request.querystring("c33")

dim sql1
dim rs0,rs5,rs6 
Set rs0 = Server.CreateObject("ADODB.Recordset")
Set rs5 = Server.CreateObject("ADODB.Recordset")
Set rs6 = Server.CreateObject("ADODB.Recordset")
set conn1 = server.CreateObject("Adodb.connection")
      
connStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & _
Server.MapPath("missyou.mdb")
conn1.Open connStr

on error resume next
Set rs0 = conn1.execute("select top 1 ID from Persons where Nick='" & trim(Session("login")) & "'")   
Set rs5 = conn1.execute("select top 1 ID from Client where ClientName='" & trim(Session("Client")) & "'")   

if rs0.eof then 
SaleID = 1
else
SaleID = rs0(0)
end if

if rs5.eof then 
'conn1.execute("insert into Client (ClientName,Address,Phone) values ("& trim(Session("Client")) & ",'','')")
'Set rs6 = conn1.execute("select top 1 ID from Client where ClientName='" & trim(Session("Client")) & "'")     
'ClientID = rs6(0)
ClientID = 1
'elseif rs6.eof then
'ClientID = rs5(0)
else 
ClientID = rs5(0)
end if




dim status
dim NickName

if Session("Status")="" then
	status=1
end if

if Session("NickName")="" then
	NickName="No NickName"
end if




sql1="insert into" 
sql1= sql1 & " InternClean(SaleID,ClientID,datebeg,dateend,status,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15,"
sql1= sql1 & "a16,a17,a18,a19,a20,a21,a22,a23,a24,a25,a26,a27,a28,a29,a30,a31,a32,a33,a34,a35,a36,a37,a38,a39,a40,a41,a42,a43"
sql1= sql1 & ",a44,a45,a46,a47,a48,a49,a50,a51,a52,a53,a54,a55,a56,a57,a58,a59,a60,a61,a62,a63,a64,a65,a66,a67,a68,a69,a70,a71,a72,a73"
sql1= sql1 & ",a74,a75,a76,a77,a78,a79,a80,a81,a82,a83,a84,a85,a86,a87,a88,a89,a90,a91,a92,a93,a94,a95,a96,a97,a98,a99,a100,a101,a102,a103"
sql1= sql1 & ",a104,a105,a106,a107,a108,a109,a110,a111,a112,a113,a114,a115,a116,a117,a118,b1,b2,b3,b4,b5,b6,b7,b8,b9,b10,b11,b12,b13,b14,b15,b16,b17,b18,b19,b20,b21,b22,b23,b24,b25,b26,b27,b28,b29,b30,b31,b32,b33,"
sql1= sql1 & "b34,b35,b36,b37,b38,b39,b40,b41,b42,b43,b44,"
sql1= sql1 & "c1,c2,c3,c4,c5,c6,c7,c8,c9,c10,c11,c12,c13,c14,c15,c16,c17,c18,c19,c20,c21,c22,c23,c24,c25,c26,c27,c28,c29,c30,c31,c32,c33, NickName ) select " 
sql1= sql1 & " " & SaleID & " ,"& ClientID & ",'01/01/2006','08/01/2006',"  & status & ",'" &  a1 & "','" & a2 & "','" & a3 & "','" & a4 & "','" & a5 & "','" & a6 & "','"& a7 & "','" & a8 & "','" & a9 & "','" & a10 & "','" & a11 & "','" & a12 & "','"& a13 & "','" & a14 & "','" & a15 & "','" & a16 & "','" & a17 & "','" & a18 & "','" & a19 & "','" & a20 & "','" & a21 & "','" & a22 & "','" & a23 & "','" & a24 & "','" & a25 & "','" & a26 & "','" & a27 & "','" & a28 & "','" & a29 & "','" & a30 & "','" & a31 & "','" & a32 & "','" & a33 & "','" & a34 & "','" & a35 & "','" & a36 & "','" & a37 & "','" & a38 & "','" & a39 & "','" & a40 & "','" & a41 & "','" & a42 & "','" & a43 & "','" & a44 & "','" & a45 & "','" & a46 & "','" & a47 & "','" & a48 & "','" & a49 & "','" & a50 & "','" & a51 & "','" & a52 & "','" & a53 & "','" & a54 & "','" & a55 & "','" & a56 & "','" & a57 & "','" & a58 & "','" & a59 & "','" & a60 & "','" & a61 & "','" & a62 & "','" & a63 & "','"  & a64 & "','" &  a65 & "','" &  a66 & "','" & a67 & "','" & a68 & "','" & a69 & "','" & a70 & "','" & a71 & "','" & a72 & "','" & a73 & "','" & a74 & "','" & a75 & "','" & a76 & "','" &  a77 & "','" &  a78 & "','"&  a79 & "','"&  a80 & "','"&  a81 & "','"&  a82 & "','"&  a83 & "','"&  a84 & "','"&  a85 & "','"&  a86 & "','"&  a87 & "','"&  a88 & "','"&  a89 & "','"&  a90 & "','"&  a91 & "','"&  a92 & "','"&  a93 & "','"&  a94 & "','"&  a95 & "','"&  a96 & "','"&  a97 & "','"&  a98 & "','"&  a99 & "','"&  a100 & "','"&  a101 & "','"&  a102 & "','"&  a103 & "','"&  a104 & "','"&  a105 & "','"&  a106 & "','"&  a107 & "','"&  a108 & "','"&  a109 & "','"&  a110 & "','"&  a111 & "','"&  a112 & "','"&  a113 & "','"&  a114 & "','"&  a115 & "','"&  a116 & "','"&  a117 & "','"&  a118 & "','" & b1 & "','" & b2 & "','" & b3 & "','" & b4 & "','" & b5 & "','" & b6 & "','"& b7 & "','" & b8 & "','" & b9 & "','" & b10 & "','" & b11 & "','" & b12 & "','"& b13 & "','" & b14 & "','" & b15 & "','" & b16 & "','" & b17 & "','" & b18 & "','" & b19 & "','" & b20 & "','" & b21 & "','" & b22 & "','" & b23 & "','" & b24 & "','" & b25 & "','" & b26 & "','" & b27 & "','" & b28 & "','" & b29 & "','" & b30 & "','" & b31 & "','" & b32 & "','" & b33 & "','" & b34 & "','" & b35 & "','"& b36 & "','"& b37 & "','"& b38 & "','"& b39 & "','"& b40 & "','"& b41 & "','"& b42 & "','"& b43 & "','" & b44 & "','" & c1 & "','" & c2 & "','" & c3 & "','" & c4 & "','" & c5 & "','" & c6 & "','"& c7 & "','" & c8 & "','" & c9 & "','" & c10 & "','" & c11 & "','" & c12 & "','"& c13 & "','" & c14 & "','" & c15 & "','" & c16 & "','" & c17 & "','" & c18 & "','" & c19 & "','" & c20 & "','" & c21 & "','" & c22 & "','" & c23 & "','" & c24 & "','" & c25 & "','" & c26 & "','" & c27 & "','" & c28 & "','" & c29 & "','" & c30 & "','" & c31 & "','" & c32 & "','" & c33 & "','" & NickName & "'"     

set rs = conn.execute(sql1)

dim MaxID
dim rs7
set rs7=Server.CreateObject("ADODB.Recordset")
set rs7 = conn.execute("select max(ID) from InternClean") 
MaxID =  rs7(0)

'response.write sql1
 
sql1="insert into StatusHistory (CleanID) select " & MaxID & ""
set rs = conn.execute(sql1) 

'response.write err.description
%>

<html>
<head>
<link rel="stylesheet" type="text/css" href="index.css">
</head>
<body>
<center>
<img src="relationshipsromance.jpeg"></img>
</center>
Инициатор:<%=Session("Login")%>|Клиент:<%=Session("client")%><table width=80% ID="Table1"><td>&nbsp;</td></table>
Дата:<%=date()%>|Контактное лицо, тел:<%=Session("contact")%><table width=80% ID="Table2"><td>&nbsp;</td></table>
Срок ответа:<%=date()%>|Дата начала работ:<%=Session("begindate")%>
<br>
1.<b><u>ВНУТРЕННЯЯ УБОРКА ПОМЕЩЕНИЙ:</u></b>
<br>
<table border=1 width=100% ID="Table3">
<tr>
<td>&nbsp;</td><td>Показатель</td><td>Ед.изм.</td><td>Здание № 1</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>1.</td><td><b>Краткие данные о здании/объекте</b></td><td></td><td>офисы</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>


<% if request.querystring("ch1") ="on" then %>
<tr>
<td>1.1</td><td>этажность</td><td>эт.</td><td>
<%=request.querystring("t1")%></td><td>&nbsp;
</td><td><%=request.querystring("Period1")%>
</td>
</tr>
<%end if %>

<% if request.querystring("ch2") ="on" then %>
<tr>
<td>1.2</td><td>год последнего ремонта</td><td>год</td><td>
<%=request.querystring("t2")%></td><td>
</td><td>
<%=request.querystring("Period2")%>
</td>
</tr>

<%end if %>

 

<% if request.querystring("ch3") ="on" then %>

<tr>
<td>1.3</td><td>высота потолков</td><td>м.</td><td>
<%=request.querystring("t3")%></td><td>
</td><td>
<%=request.querystring("Period3")%>
</td>
</tr>
<%end if %>

<% if request.querystring("ch4") ="on" then %>
 

<tr>
<td>2.</td><td><b>Общая площадь</b></td><td>кв.м</td><td>
<%=request.querystring("t4")%></td><td>
</td><td>
<%=request.querystring("Period4")%>
</td>
</tr>
<%end if %>

 

<% if request.querystring("ch5") ="on" then %>

<tr>
<td>2.1</td><td>Площадь каждого этажа</td><td>кв.м</td><td>
<%=request.querystring("t5")%></td><td>
</td><td>
<%=request.querystring("Period5")%>
</td>
</tr>

<%end if %>

<% if request.querystring("ch6") ="on" then %>
<tr>
<td>2.2</td><td>Кабинеты VIP</td><td>кол-во</td><td>
<%=request.querystring("t6_1")%></td><td>
</td><td>
<%=request.querystring("Period6")%>
</td>
</tr>

<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<%=request.querystring("t6_2")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<%end if %>


<% if request.querystring("ch7") ="on" then %>

<tr>
<td>2.3</td><td>Офисные помещения</td><td>кол-во</td><td>
<%=request.querystring("t7_1")%></td><td>
</td><td>
<%=request.querystring("Period7")%>
</td>
</tr>


<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<%=request.querystring("t7_2")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<%end if %>

<% if request.querystring("ch8") ="on" then %>
<tr>
<td>2.4</td><td>Складские помещения</td><td>кол-во</td><td>
<%=request.querystring("t8_1")%></td><td>
</td><td>
<%=request.querystring("Period8")%>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<%=request.querystring("t8_2")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<%end if %>


<% if request.querystring("ch9") ="on" then %>

<tr>
<td>2.5</td><td>Складские помещения</td><td>кол-во</td><td>
<%=request.querystring("t9_1")%></td><td>
</td><td>
<%=request.querystring("Period9")%></td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<%=request.querystring("t9_2")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<%end if %>

 

<% if request.querystring("ch10") ="on" then %>
<tr>
<td>2.6</td><td>Технические помещения, подвалы</td><td>кол-во</td><td>
<%=request.querystring("t10_1")%></td><td>
</td><td>
<%=request.querystring("Period10")%>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<%=request.querystring("t10_2")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<%end if %>


<% if request.querystring("ch11") ="on" then %>
<tr>
<td>2.7</td><td>Коридоры</td><td>кол-во</td><td>
<%=request.querystring("t11_1")%></td><td>
</td><td>
<%=request.querystring("Period11")%>
</td>
</tr>


<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<%=request.querystring("t11_2")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<%end if %>

<% if request.querystring("ch12") ="on" then %>
<tr>
<td>2.8</td><td>Лестницы</td><td>кол-во</td><td>
<%=request.querystring("t12_1")%></td><td>
</td><td>
<%=request.querystring("Period12")%>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<%=request.querystring("t12_2")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<%end if %>

<% if request.querystring("ch13") ="on" then %>
<tr>
<td>2.9</td><td>Лифты</td><td>кол-во</td><td>
<%=request.querystring("t13_1")%></td><td>
</td><td>
<%=request.querystring("Period13")%>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<%=request.querystring("t13_2")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<%end if %>

<% if request.querystring("ch14") ="on" then %>
<tr>
<td>2.10</td><td>Эскалаторы</td><td>кол-во</td><td>
<%=request.querystring("t14_1")%></td><td>
</td><td>
<%=request.querystring("Period14")%>
</td>
</tr>

<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<%=request.querystring("t14_2")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<%end if %>

<% if request.querystring("ch15") ="on" then %>
<tr>
<td>2.11</td><td>Санузлы</td><td>кол-во</td><td>
<%=request.querystring("t15_1")%></td><td>
</td><td>
<%=request.querystring("Period15")%>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<%=request.querystring("t15_2")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<%end if %>

<% if request.querystring("ch16") ="on" then %>
<tr>
<td>2.12</td><td>холлы, вестибюли</td><td>кол-во</td><td>
<%=request.querystring("t16_1")%></td><td>
</td><td>
<%=request.querystring("Period16")%>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<%=request.querystring("t16_2")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<%end if %>

<% if request.querystring("ch17") ="on" then %>
<tr>
<td>2.13</td><td>Другие площади (по возм.расшифровать)</td><td></td><td>
<%=request.querystring("t17_0")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>кол-во</td><td>
<%=request.querystring("t17_1")%></td><td>
</td><td>
<%=request.querystring("Period17")%>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<%=request.querystring("t17_2")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<%end if %>


<% if request.querystring("ch18") ="on" then %>
<tr>
<td>3.</td><td><b>К-во сотрудников Вашей компании и посетителей, чел. в среднем в день</b></td><td>чел.</td><td>
<%=request.querystring("t18")%></td><td>
</td><td>&nbsp;</td>
</tr>
<%end if %>


<% if request.querystring("ch19_1") ="on" or request.querystring("ch19_2") ="on" or request.querystring("ch19_3") ="on" or request.querystring("ch19_4")="on" then %>
<tr>
<td>4.</td><td><b>Необходимость в обеспечении с/у расходными материалами (примерный расход в месяц)</b></td><td>кол-во</td><td>&nbsp;</td><td>&nbsp;</td><td>
<%=request.querystring("Period19")%>
</td>
</tr>

<% if request.querystring("ch19_1") ="on" then %>
<tr>
<td>&nbsp;</td><td>туалетная бумага</td><td>рул./мес.</td><td>
<%=request.querystring("t19_1")%></td><td>
</td><td>&nbsp;</td>
</tr>

<%end if%>

<% if request.querystring("ch19_2") ="on" then %>
<tr>
<td>&nbsp;</td><td>жидкое мыло</td><td>литр/мес.</td><td>
<%=request.querystring("t19_2")%></td><td>
</td><td>&nbsp;</td>
</tr>

<%end if%>

<% if request.querystring("ch19_3") ="on" then %>

<tr>
<td>&nbsp;</td><td>бум полотенца</td><td>лист/мес.</td><td>
<%=request.querystring("t19_3")%></td><td>
</td><td>&nbsp;</td>
</tr>

<%end if%>

<% if request.querystring("ch19_4") ="on" then %>

<tr>
<td>&nbsp;</td><td>бумажные сидения д/унитаза</td><td>шт./мес.</td><td>
<%=request.querystring("t19_4")%></td><td>
</td><td>&nbsp;</td>
</tr>

<%end if %>

<%end if %>

<% if request.querystring("ch20") ="on" then %>
<tr>
<td>5.</td><td><b>Поверхности</b></td><td>чел.</td><td>&nbsp;</td><td>
</td><td>
<%=request.querystring("Period20")%>
</td>
</tr>
 

<% if request.querystring("ch20_1") ="on" then %>

<tr>
<td>5.1</td><td>Мягкие покрытия (ковролин)</td><td>кв.м</td><td>
<%=request.querystring("t20_1")%></td><td>
</td><td>&nbsp;</td>
</tr>
 

<%end if %>

<% if request.querystring("ch20_2") ="on" then %>

<tr>
<td>5.2</td><td>Полутвердые покрытия (линолеум, паркет, мармолеум, наливной пол, ламинат)</td><td>кв.м</td><td>
<%=request.querystring("t20_2")%></td><td>
</td><td>&nbsp;</td>
</tr>
 

<%end if %>

<% if request.querystring("ch20_3") ="on" then %>

<tr>
<td>5.3</td><td>Твердые покрытия (плитка, мрамор, гранит) пол/стены</td><td>кв.м</td><td>
<%=request.querystring("t20_3")%></td><td>
</td><td>&nbsp;</td>
</tr>
 

<%end if %>

<% if request.querystring("ch20_4") ="on" then %>

<tr>
<td>5.4</td><td>Стеклянные поверхности</td><td>кв.м</td><td>
<%=request.querystring("t20_4")%></td><td>
</td><td>&nbsp;</td>
</tr>
 

<%end if %>

<% if request.querystring("ch20_5") ="on" then %>

<tr>
<td>5.5</td><td>Металлические поверхности</td><td>кв.м</td><td>
<%=request.querystring("t20_5")%></td><td>
</td><td>&nbsp;</td>
</tr>
 

<%end if %>

<% if request.querystring("ch20_6") ="on" then %>

<tr>
<td>5.6</td><td>Офисные перегородки</td><td>шт.</td><td>
<%=request.querystring("t20_6")%></td><td>
</td><td>&nbsp;</td>
</tr>
 

<%end if %>

<% if request.querystring("ch20_7") ="on" then %>

<tr>
<td>5.7</td><td>Офисные места</td><td>шт.</td><td>
<%=request.querystring("t20_7")%></td><td>
</td><td>&nbsp;</td>
</tr>
 

<%end if %>

<% if request.querystring("ch20_8") ="on" then %>

<tr>
<td>5.8</td><td>Кожаная мебель</td><td>шт.</td><td>
<%=request.querystring("t20_8")%></td><td>
</td><td>&nbsp;</td>
</tr>
 

<%end if %>

<% if request.querystring("ch20_9") ="on" then %>

<tr>
<td>5.9</td><td>Пластиковая мебель</td><td>шт.</td><td>
<%=request.querystring("t20_9")%></td><td>
</td><td>&nbsp;</td>
</tr>
 

<%end if %>

<% if request.querystring("ch20_10") ="on" then %>

<tr>
<td>5.10</td><td>Деревянная мебель</td><td>шт.</td><td>
<%=request.querystring("t20_10")%></td><td>
</td><td>&nbsp;</td>
</tr>
 

<%end if %>

<% if request.querystring("ch20_11") ="on" then %>

<tr>
<td>5.11</td><td>Другие поверхности</td><td>шт./кв.м</td><td>
<%=request.querystring("t20_11")%></td><td>
</td><td>&nbsp;</td>
</tr>
 

<%end if %>

<%end if %>

<% if request.querystring("ch21") ="on" then %>
<tr>
<td>6.</td><td><b>Существующая у Вас уборка</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
</td><td>
<%=request.querystring("Period21")%>
</td>
</tr>

<tr>
<td>6.1</td><td>График проведения основной комплексной уборки</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>с:</td><td>час</td><td>
<%=request.querystring("t22")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>до:</td><td>час</td><td>
<%=request.querystring("t23")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>дней в году:</td><td>дни</td><td>
<%=request.querystring("t24")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>6.2</td><td>График проведения поддерживающей уборки</td><td></td><td></td><td></td><td></td>
</tr>

<tr>
<td>&nbsp;</td><td>с:</td><td>час</td><td>
<%=request.querystring("t25")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>до:</td><td>час</td><td>
<%=request.querystring("t26")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>дней в году:</td><td>дни</td><td>
<%=request.querystring("t27")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>



<tr>
<td>6.3</td><td>Количество уборщиков</td><td>чел.</td><td>
<%=request.querystring("t28")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<%end if %>
<% if request.querystring("ch22") ="on" then %>
<tr>
<td>7.</td><td><b>Помещения для размещения производственного персонала и оборудования</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
</td><td>
<%=request.querystring("Period22")%>
</td>
</tr>

<tr>
<td>7.1</td><td>Наличие помещения</td><td>да/нет</td><td><%=request.querystring("choice1")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.2</td><td>Наличие телефона в помещении</td><td>да/нет</td><td>
<%=request.querystring("choice2")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.3</td><td>площадь помещения</td><td>кв.м</td><td>
<%=request.querystring("t29")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<%end if %>
<% if request.querystring("ch23") ="on" then %>
<tr>
<td>8.</td><td><b>Помещения для размещения управленческого персонала и автоматизированных рабочих мест (для объектов более 10000 

кв.м)</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
</td><td>
<%=request.querystring("Period23")%>
</td>
</tr>

<tr>
<td>8.1</td><td>Наличие помещения</td><td>да/нет</td><td>
<%=request.querystring("choice3")%></td><td></td><td></td>
</tr>
<tr>
<td>8.2</td><td>Площадь помещения</td><td>кв.м</td><td>
<%=request.querystring("t30")%></td><td></td><td></td>
</tr>
<tr>
<td>8.3</td><td>кол-во автоматизированных рабочих мест</td><td>шт.</td><td><%=request.querystring("t31")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<%end if %>
</table>
<br>
2.<b><u>УБОРКА ТЕРРИТОРИИ:</u></b>
Предоставьте, пожалуйста, информацию по всем зданиям, в которых необходимо осуществить уборку.
<br>
<table border=1 width=100% ID="Table4">
<tr>
<td>&nbsp;</td><td>Показатель</td><td>Ед.изм.</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>1.</td><td><b>Общая убираемая территория</b></td><td>кв.м</td><td><%=request.querystring("b1")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% if request.querystring("k1") then%>
<tr>
<td>1.1</td><td>Площади с указанием покрытия (тротуары, стоянки, газоны, проезды, 

другое)</td><td>&nbsp;</td><td><%=request.querystring("b2")%></td><td></td><td>
<%=request.querystring("b3")%>
</td>
</tr>
<% if request.querystring("k2") then%>
<tr>
<td>1.1.1</td><td>Тротуары</td><td>год</td><td><%=request.querystring("b4")%></td><td></td><td>
<%=request.querystring("b5")%>
</td>
</tr>
<%end if %>
<% if request.querystring("k3") then%>
<tr>
<td>1.1.2</td><td>Стоянки</td><td>м.</td><td><%=request.querystring("b6")%></td><td></td><td>
<%=request.querystring("b7")%>
</td>
</tr>
<%end if %>
<% if request.querystring("k4") then%>
<tr>
<td>1.1.3</td><td>Газоны</td><td>м.</td><td><%=request.querystring("b8")%></td><td></td><td>
<%=request.querystring("b9")%>
</td>
</tr>
<%end if %>
<% if request.querystring("k5") then%>
<tr>
<td>1.1.4</td><td>Проезды</td><td>м.</td><td><%=request.querystring("b10")%></td><td></td><td>
<%=request.querystring("b11")%>
</td>
</tr>
<%end if %>
<%end if %>

<% if request.querystring("k6")="on" then%>
<tr>
<td>2.</td><td><b>Существующая у Вас уборка территории</b></td><td>кв.м</td><td><%=request.querystring("b12")%></td><td></td><td>
<%=request.querystring("b13")%>
</td>
</tr>

<tr>
<td>2.1</td><td>График уборки</td><td></td><td></td><td></td><td></td>
</tr>

<tr>
<td>&nbsp;</td><td>с:</td><td>час</td><td><%=request.querystring("b14")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>до:</td><td>час</td><td><%=request.querystring("b15")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>дней в году:</td><td>дни</td><td><%=request.querystring("b16")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>


<tr>
<td>2.2</td><td>Количество уборщиков</td><td>чел.</td><td><%=request.querystring("b17")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<%end if %>

<% if request.querystring("k7")="on" then%>
<tr>
<td>3.</td><td><b>Инвентаризационная ведомость техники (оборудования) 

на</b></td><td>чел.</td><td><%=request.querystring("b18")%></td><td></td><td>&nbsp;</td>
</tr>
<tr>
<td>3.1</td><td>Офисные помещения</td><td>мес</td><td><%=request.querystring("b19")%></td><td></td><td>
<%=request.querystring("b20")%>
</td>
</tr>
<tr>
<td>3.2</td><td>Офисные помещения</td><td>мес</td><td><%=request.querystring("b21")%></td><td></td><td>
<%=request.querystring("b22")%>
</td>
</tr>
<tr>
<td>3.3</td><td>Офисные помещения</td><td>мес</td><td><%=request.querystring("b23")%></td><td></td><td>
<%=request.querystring("b24")%>
</td>
</tr>
<tr>
<td>3.4</td><td>Офисные помещения</td><td>мес</td><td><%=request.querystring("b25")%></td><td></td><td>
<%=request.querystring("b26")%>
</td>
</tr>
<%end if %>

<% if request.querystring("k8")="on" then%>
<tr>
<td>4.</td><td><b>Наличие места для стоянки уборочной техники,</b> оборудованного отоплением и складом для хранения расходных материалов и запчастей, а также 

проведения ремонта техники, оборудования и инвентаря</td><td>да/нет</td><td><%=request.querystring("b27")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<%end if %>
<% if request.querystring("k9")="on" then%>
<tr>
<td>5.</td><td><b>Потребность в стрижке газона, покосе травы (площадь и кол-во 

раз)</b></td><td>да/нет</td><td><%=request.querystring("b28")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>5.1</td><td>площадь</td><td>кв.м</td><td><%=request.querystring("b29")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td></td><td>количество стрижек газонов и покосов травы</td><td>кол-во</td><td><%=request.querystring("b30")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<%end if %>
<% if request.querystring("k10")="on" then%>
<tr>
<td>6.</td><td><b>Качество уборки территории зимой </b>(A - до утрамбованного слоя снега/ Б - до 

покрытия)</td><td>А/Б</td><td><%=request.querystring("b31")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<%end if %>

<% if request.querystring("k11")="on" then%>
<tr>
<td>7.</td><td><b>Вывоз снега в зимний период (объем вывоза за 

зиму)</b></td><td>куб.м.</td><td><%=request.querystring("b32")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<%end if %>

<% if request.querystring("k12")="on" then%>
<tr>
<td>8.</td><td><b>Вывоз ТБО, примерный месячный объем мусора</b></td><td>куб.м.</td><td><%=request.querystring("b33")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<%end if %>



</table>
<br>
<br>

<table border=0 width=100% ID="Table5">
<tr><td align=left>3.<b><u>ДОПОЛНИТЕЛЬНЫЕ УСЛУГИ:</u></b></td></tr>
</table>
<table border=1 width=100% ID="Table6">
<tr>
<td>&nbsp;</td><td>Показатель</td><td>Ед.изм.</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% if request.querystring("c2")="on" then%>
<tr>
<td><b>1.</b></td><td><b>Организация работы прачечных</b></td><td>да/нет</td><td><select name="c1" ID="Select1">
<option value="да" selected>да</option>
<option value="нет" >нет</option>
</select></td><td><input type=checkbox name = c2 ID="Checkbox1"></td>
</tr>
<% if request.querystring("c4")="on" then%>
<tr>
<td>1.1</td><td>Наличие химчистки-прачечной</td><td>да/нет</td><td><select name="c3" ID="Select2">
<option value="да" selected>да</option>
<option value="нет" >нет</option>
</select></td><td><input type=checkbox name = c4 ID="Checkbox2"></td>
</tr>
<%end if%>
<% if request.querystring("c6")="on" then%>
<tr>
<td>1.1.1</td><td>количество обслуживающего персонала</td><td>чел.</td><td><input type=text name = c5 ID="Text1"></td><td><input type=checkbox name = c6 ID="Checkbox3"></td>
</tr>
<%end if%>
<% if request.querystring("c8")="on" then%>
<tr>
<td>1.2</td><td>Потребность в стирке спецодежды, объем в месяц</td><td>кг</td><td><input type=text name = c7 ID="Text2"></td><td><input type=checkbox name = c8 ID="Checkbox4"></td>
</tr>
<%end if%>
<%end if%>
<% if request.querystring("c10")="on" then%>
<tr>
<td><b>2.</b></td><td><b>Глубокая химическая чистка ковровых покрытий</b></td><td>да/нет</td><td><select name="c9" ID="Select3">
<option value="да" selected>да</option>
<option value="нет" >нет</option>
</select></td><td><input type=checkbox name = c10 ID="Checkbox5"></td>
</tr>
<tr>
<td>2.1</td><td>Площадь ковровых покрытий</td><td>кв.м</td><td><input type=text name = c11 ID="Text3"></td><td><input type=checkbox name = c12 ID="Checkbox6"></td>
</tr>
<tr>
<td>2.2</td><td>Периодичность проведения</td><td>&nbsp;</td><td><input type=text name = c13 ID="Text4"></td> <td>&nbsp;</td>
</tr>
<%end if%>
<% if request.querystring("c15")="on" then%>
<tr>
<td><b>3.</b></td><td><b>Нанесение полимерного лака на линолеумные полы</b></td><td>да/нет</td><td><select name="c14" ID="Select4">
<option value="да" selected>да</option>
<option value="нет" >нет</option>
</select></td><td><input type=checkbox name = c15 ID="Checkbox7"></td>
</tr>
<tr>
<td>3.1</td><td>Площадь линолеума</td> <td>кв.м</td><td><input type=text name = c16 ID="Text5"></td>
</tr>
<tr>
<td>3.2</td><td>Периодичность нанесения лака</td> <td>&nbsp;</td><td><input type=text name = c16 ID="Text6"></td><td>&nbsp;</td>
</tr>
<%end if%>
<% if request.querystring("c18")="on" then%>
<tr>
<td><b>4.</b></td><td><b>Мытье окон (площадь окон с одной стороны)</b></td> <td>да/нет</td><td><select name="c17" ID="Select5">
<option value="да" selected>да</option>
<option value="нет" >нет</option>
</select></td><td><input type=checkbox name = c18 ID="Checkbox8"></td>
</tr>
<% if request.querystring("c20")="on" then%>
<tr>
<td>4.1</td><td>Легкий доступ к окнам (с пола)</td> <td>кв.м</td><td><input type=text name = c19 ID="Text7"></td><td><input type=checkbox name = c20 ID="Checkbox9"></td>
</tr>
<%end if%>
<% if request.querystring("c22")="on" then%>
<tr>
<td>4.2</td><td>Затрудненный (со стремянки)</td> <td>кв.м</td><td><input type=text name = c21 ID="Text8"></td><td><input type=checkbox name = c22 ID="Checkbox10"></td>
</tr>
<%end if%>
<% if request.querystring("c24")="on" then%>
<tr>
<td>4.3</td><td> С помощью промышленных альпинистов </td> <td>кв.м</td><td><input type=text name = c23 ID="Text9"></td><td><input type=checkbox name = c24 ID="Checkbox11"></td>
</tr>
<%end if%>
<%end if%>
<% if request.querystring("c25")="on" then%>
<tr>
<td><b>5.</b></td><td><b>Другие услуги, которые вы хотели бы получать</b></td> <td>&nbsp;</td><td><input type=checkbox name = c25 ID="Checkbox12"></td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td><input type=text name = c26 ID="Text10"></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td><input type=text name = c27 ID="Text11"></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td><input type=text name = c28 ID="Text12"></td><td>&nbsp;</td>
</tr>
<tr>
<td></td><td></td><td></td>
</tr>
<%end if%>
</table>

<br><br><br>
<br><br><br>
Согласовано:<br><br><br>
Генеральный директор _______________________________________/Московиц Д.С.
<br><br><br><br><br><br><br><br><br><br><br><br>
<table width=100% ID="Table7">
<tr><td>Tip-Top 

Cleaning</td><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;</td><td>Тип-Топ Клининг</td></tr>		
<tr><td>Ul.Ordzhonikidze,11</td>
<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</td><td>ул.Орджоникидзе,11</td></tr>		
<tr><td>115419 Moscow,Russia</td>
<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;</td><td>115419 Москва, Россия</td></tr>
<tr><td>www.tiptop.com.ru</td>
<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;</td><td>www.tiptop.com.ru</td></tr>
<tr><td>+7(095)234 45 20</td><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td>+7(095)234 45 20</td></tr>
</table>




</body>

</html>