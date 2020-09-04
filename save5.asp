<%
'on error resume next
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

dim id
id =Session("ID")
if left(id,1)="'" then 
	id=right(left(id,len(id)-1),len(id)-2 )
end if

'--------------Сохраняем данные
dim a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15,a16,a17,a18,a19,a20,a21,a22,a23,a24
dim a25,a26,a27,a28,a29,a30,a31,a32,a33,a34,a35,a36,a37,a38,a39,a40,a41,a42,a43,a44,a45,a46,a47
dim a48,a49,a50,a51,a52,a53,a54,a55,a56,a57,a58,a59,a60,a61,a62,a63,a64,a65,a66,a67,a68,a69,a70
dim a71,a72,a73,a74,a75,a76,a77
dim a78,a79,a80,a81,a82,a83,a84,a85,a86,a87,a88,a89,a90,a91,a92,a93,a94,a95,a96,a97,a98,a99,a100,a101,a102,a103,a104,a105,a106,a107
dim a108,a109,a110,a111,a112,a113,a114,a115,a116,a117,a118
dim d1,d2,d3,d4,d5,d6,d7,d8,d9,d10,d11,d12


d1=request.querystring("d1")
d2=request.querystring("d2")
d3=request.querystring("d3")
d4=request.querystring("d4")
d5=request.querystring("d5")
d6=request.querystring("d6")
d7=request.querystring("d7")
d8=request.querystring("d8")
d9=request.querystring("d9")
d10=request.querystring("d10")
d11=request.querystring("d11")
d12=request.querystring("d12")

'response.write request.querystring


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





sql1="Update InternClean " 

sql1= sql1 & " set ClientID='" &  d1 & "',dateend='" & d2 & "',NickName='" & d3 & "',datebeg='" & d5 &  "',Comments='" & d7 & "',Address1='" & d8 & "',Address2='" & d9 & "' where ID = " & id & ""

set rs = conn.execute(sql1) 



err.clear


response.write err.description
response.write sql1
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
<a href="common.asp?ID=<%=Session("ID")%>"><font  size=1>Общие данные(редактировать)</font></a>
		<hr>

		<a href="vnutr.asp?ID='<%=Session("ID")%>'"><font  size=1>Заявка (редактировать)</font></a>
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