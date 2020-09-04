<%
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

'--------------Сохраняем данные
dim a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15,a16,a17,a18,a19,a20,a21,a22,a23,a24
dim a25,a26,a27,a28,a29,a30,a31,a32,a33,a34,a35,a36,a37,a38,a39,a40,a41,a42,a43,a44,a45,a46,a47
dim a48,a49,a50,a51,a52,a53,a54,a55,a56,a57,a58,a59,a60,a61,a62,a63,a64,a65,a66,a67,a68,a69,a70
dim a71,a72,a73,a74,a75,a76,a77
dim a78,a79,a80,a81,a82,a83,a84,a85,a86,a87,a88,a89,a90,a91,a92,a93,a94,a95,a96,a97,a98,a99,a100,a101,a102,a103,a104,a105,a106,a107
dim a108,a109,a110,a111,a112,a113,a114,a115,a116,a117,a118


a1=request.querystring("t1")
if a1<>"" then a2 ="on"
'a2=request.querystring("ch1")
a3=request.querystring("Period1")
a4=request.querystring("t2")
if a4<>"" then a5 ="on"
'a5=request.querystring("ch2")
a6=request.querystring("Period2")
a7=request.querystring("t3")
if a7<>"" then a8 ="on"
'a8=request.querystring("ch3")
a9=request.querystring("Period3")
a10=request.querystring("t4")

if a10<>"" then a11 ="on"
'a11=request.querystring("ch4")
a12=request.querystring("Period4")
a13=request.querystring("t5")
if a13<>"" then a14 ="on"
a14=request.querystring("ch5")
a15=request.querystring("Period5")
a16=request.querystring("t6_1")
if a16<>"" then a17 ="on"
'a17=request.querystring("ch6")
a18=request.querystring("Period6")
a19=request.querystring("t6_2")
a20=request.querystring("t7_1")
a21=request.querystring("ch7")

a22=request.querystring("Period7")

a23=request.querystring("t7_2")
a24=request.querystring("t8_1")
a25=request.querystring("Period8")
a26=request.querystring("ch8")

a27=request.querystring("t8_2")

a28=request.querystring("t9_1")
a29=request.querystring("ch9")
a30=request.querystring("Period9")
a31=request.querystring("t9_2")
a32=request.querystring("t10_1")
a33=request.querystring("ch10")
a34=request.querystring("Period10")
a35=request.querystring("t10_2")
a36=request.querystring("t11_1")
a37=request.querystring("ch11")
a38=request.querystring("Period11")
a39=request.querystring("t11_2")
a40=request.querystring("t12_1")
a41=request.querystring("ch12")
a42=request.querystring("Period12")
a43=request.querystring("t12_2")
a44=request.querystring("t13_1")
a45=request.querystring("ch13")

a46=request.querystring("Period13")

a47=request.querystring("t13_2")
a48=request.querystring("t14_1")
a49=request.querystring("ch14")
a50=request.querystring("Period14")
a51=request.querystring("t14_2")
a52=request.querystring("t15_1")
a53=request.querystring("ch15")
a54=request.querystring("Period15")
a55=request.querystring("t15_2")
a56=request.querystring("t16_1")
a57=request.querystring("ch16")
a58=request.querystring("Period16")
a59=request.querystring("t16_2")
a60=request.querystring("t17_0")
a61=request.querystring("t17_1")
a62=request.querystring("ch17")

a63=request.querystring("Period7")

a64=request.querystring("t17_2")
a65=request.querystring("t18")
a66=request.querystring("ch18")
a67=request.querystring("Period19")
a68=request.querystring("t19_1")
a69=request.querystring("ch19_1")
a70=request.querystring("t19_2")
a71=request.querystring("ch19_2")
a72=request.querystring("t19_3")
a73=request.querystring("ch19_3")
a74=request.querystring("t19_4")
a75=request.querystring("ch19_4")
a76=request.querystring("ch20")
a77=request.querystring("Period20")
a78=request.querystring("t20_1")
a79=request.querystring("ch20_1")
a80=request.querystring("t20_2")
a81=request.querystring("ch20_2")
a82=request.querystring("t20_3")
a83=request.querystring("ch20_3")
a84=request.querystring("t20_4")
a85=request.querystring("ch20_4")
a86=request.querystring("t20_5")
a87=request.querystring("ch20_5")
a88=request.querystring("t20_6")
a89=request.querystring("ch20_6")
a90=request.querystring("t20_7")
a91=request.querystring("ch20_7")
a92=request.querystring("t20_8")
a93=request.querystring("ch20_8")
a94=request.querystring("t20_9")
a95=request.querystring("ch20_9")
a96=request.querystring("t20_10")
a97=request.querystring("ch20_10")
a98=request.querystring("t20_11")
a99=request.querystring("ch20_11")
a100=request.querystring("ch21")
a101=request.querystring("Period21")
a102=request.querystring("t22")
a103=request.querystring("t23")
a104=request.querystring("t24")
a105=request.querystring("t25")
a106=request.querystring("t26")
a107=request.querystring("t27")
a108=request.querystring("t28")
a109=request.querystring("ch22")
a110=request.querystring("Period22")
a111=request.querystring("choice1")
a112=request.querystring("choice2")
a113=request.querystring("t29")
a114=request.querystring("ch23")
a115=request.querystring("Period23")
a116=request.querystring("choice3")
a117=request.querystring("t30")
a118=request.querystring("t31")


dim b1,b2,b3,b4,b5,b6,b7,b8,b9,b10,b11,b12,b13,b14,b15,b16,b17,b18,b19,b20,b21,b22,b23,b24,b25,b26,b27,b28,b29,b30,b31,b32,b33
dim b34,b35,b36,b37,b38,b39,b40,b41,b42,b43,b44
b1=request.querystring("b1")
b2=request.querystring("b2")
b3=request.querystring("k1")
b4=request.querystring("b3")
b5=request.querystring("b4")
b6=request.querystring("k2")
b7=request.querystring("b5")
b8=request.querystring("b6")
b9=request.querystring("k3")
b10=request.querystring("b7")
b11=request.querystring("b8")
b12=request.querystring("k4")
b13=request.querystring("b9")
b14=request.querystring("b10")
b15=request.querystring("k5")
b16=request.querystring("b11")
b17=request.querystring("b12")
b18=request.querystring("k6")
b19=request.querystring("b13")
b20=request.querystring("b14")
b21=request.querystring("b15")
b22=request.querystring("b16")
b23=request.querystring("b17")
b24=request.querystring("b18")
b25=request.querystring("k7")
b26=request.querystring("b19")
b27=request.querystring("k8")
b28=request.querystring("b20")
b29=request.querystring("b21")
b30=request.querystring("k9")
b31=request.querystring("b22")
b32=request.querystring("b23")
b33=request.querystring("k10")
b34=request.querystring("b24")
b35=request.querystring("b25")
b36=request.querystring("k11")
b37=request.querystring("b26")
b38=request.querystring("b27")
b39=request.querystring("b28")
b40=request.querystring("b29")
b41=request.querystring("b30")
b42=request.querystring("b31")
b43=request.querystring("b32")
b44=request.querystring("b33")

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

sql1="Update InternClean " 

sql1= sql1 & " select Status from InternClean where ID=" & Session ("ID")& ""   

err.clear
set rs = conn.execute(sql1) 


end if

if Session("NickName")="" then
	NickName="No NickName"
end if



status = rs(0)

response.write "00000000000"

response.write rs(0)

response.write "11111111111111"

if len(rs(0))=0 then 
Status=1
else
Status=Status+1
end if       

sql1="Update InternClean " 

sql1= sql1 & " set Status=" &  status & " where ID=" & Session ("ID")& ""   

err.clear
set rs = conn.execute(sql1) 

sql1="Update StatusHistory " 

if status=2 then
sql1= sql1 & " set Creator='"& Session("login") &"',App1=true, Statusdate='" & date() & "',Statusnum=" & status & "  where CleanID=" & Session ("ID")& ""   
end if

if status=3 then
sql1= sql1 & " set Creator='"& Session("login") &"',App2=true, Statusdate='" & date() & "',Statusnum=" & status & "  where CleanID=" & Session ("ID")& ""   
end if

if status=4 then
sql1= sql1 & " set Creator='"& Session("login") &"',App3=true, Statusdate='" & date() & "',Statusnum=" & status & "  where CleanID=" & Session ("ID")& ""   
end if

if status=5 then
sql1= sql1 & " set Creator='"& Session("login") &"',App4=true, Statusdate='" & date() & "',Statusnum=" & status & "  where CleanID=" & Session ("ID")& ""   
end if

if status=6 then
sql1= sql1 & " set Creator='"& Session("login") &"',App5=true, Statusdate='" & date() & "',Statusnum=" & status & "  where CleanID=" & Session ("ID")& ""   
end if

if status=7 then
sql1= sql1 & " set Creator='"& Session("login") &"',App6=true, Statusdate='" & date() & "',Statusnum=" & status & "  where CleanID=" & Session ("ID")& ""   
end if

if status=8 then
sql1= sql1 & " set Creator='"& Session("login") &"',App7=true, Statusdate='" & date() & "',Statusnum=" & status & "  where CleanID=" & Session ("ID")& ""   
end if



'response.write sql1

set rs = conn.execute(sql1) 

response.write err.description
'err.clear

sql1="Update InternClean " 

sql1= sql1 & " set b1='" & b1 & "',b2='" & b2 & "',b3='" & b3 & "',b4='" & b4 & "',b5='" & b5 & "',b6='" & b6 & "',b7='"& b7 & "',b8='" & b8 & "',b9='" & b9 & "',b10='" & b10 & "',b11='" & b11 & "',b12='" & b12 & "',b13='"& b13 & "',b14='" & b14 & "',b15='" & b15 & "',b16='" & b16 & "',b17='" & b17 & "',b18='" & b18 & "',b19='" & b19 & "',b20='" & b20 & "',b21='" & b21 & "',b22='" & b22 & "',b23='" & b23 & "',b24='" & b24 & "',b25='" & b25 & "',b26='" & b26 & "',b27='" & b27 & "',b28='" & b28 & "',b29='" & b29 & "',b30='" & b30 & "',b31='" & b31 & "',b32='" & b32 & "',b33='" & b33 & "',b34='" & b34 & "',b35='" & b35 & "',b36='"& b36 & "',b37='"& b37 & "',b38='"& b38 & "',b39='"& b39 & "',b40='"& b40 & "',b41='"& b41 & "',b42='"& b42 & "',b43='"& b43 & "',b44='" & b44 & "',c1='" & c1 & "',c2='" & c2 & "',c3='" & c3 & "',c4='" & c4 & "',c5='" & c5 & "',c6='" & c6 & "',c7='"& c7 & "',c8='" & c8 & "',c9='" & c9 & "',c10='" & c10 & "',c11='" & c11 & "',c12='" & c12 & "',c13='"& c13 & "',c14='" & c14 & "',c15='" & c15 & "',c16='" & c16 & "',c17='" & c17 & "',c18='" & c18 & "',c19='" & c19 & "',c20='" & c20 & "',c21='" & c21 & "',c22='" & c22 & "',c23='" & c23 & "',c24='" & c24 & "',c25='" & c25 & "',c26='" & c26 & "',c27='" & c27 & "',c28='" & c28 & "',c29='" & c29 & "',c30='" & c30 & "',c31='" & c31 & "',c32='" & c32 & "',c33='" & c33 & "', NickName='" & NickName & "' where ID=" & Session ("ID")& ""   






'set rs = conn.execute(sql1) 


'--------------
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
<table>
<tr>
 <table>
  
 <td valign="top">
		<a href="vnutr.asp?ID='<%=Session("ID")%>'"><font  size=1>ВНУТРЕННЯЯ УБОРКА ПОМЕЩЕНИЙ (редактировать)</font></a>
		<hr>



		<a href="terr.asp?ID='<%=Session("ID")%>'"><font  size=1>УБОРКА ТЕРРИТОРИИ (редактировать)</font></a>
		<hr>

		<a href="spec.asp?ID='<%=Session("ID")%>'"><font  size=1>ДОПОЛНИТЕЛЬНЫЕ УСЛУГИ (редактировать)</font></a>
		<hr>
                <a href="default1_0_nonedit.asp?ID='<%=Session("ID")%>'"><font  size=1>Заявка (1 PAGE)</font></a>
		<hr> 
               <a href="uploadform.asp"><font  size=1>Загрузить файл</font></a>
				<hr>
 </td>
<td valign="top">
		Данные обновлены...
</td>
</tr>

</table>
