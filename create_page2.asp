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
dim sum0,note0,sum1,note1,sum2,note2,sum3,note3,sum4,note4,sum5,note5,sum6,note6,sum7,note7,sum8,note8,sum9,note9,sum10,note10,sum11,note11,sum12,note12,sum13,note13,sum14,note14,sum15,note15,sum16,note16,sum17,note17

sum0=request.querystring("sum0")
sum1=request.querystring("sum1")
sum2=request.querystring("sum2")
sum3=request.querystring("sum3")
sum4=request.querystring("sum4")
sum5=request.querystring("sum5")
sum6=request.querystring("sum6")
sum7=request.querystring("sum7")
sum8=request.querystring("sum8")
sum9=request.querystring("sum9")
sum10=request.querystring("sum10")
sum11=request.querystring("sum11")
sum12=request.querystring("sum12")
sum13=request.querystring("sum13")
sum14=request.querystring("sum14")
sum15=request.querystring("sum15")
sum16=request.querystring("sum16")
sum17=request.querystring("sum17")
note0=request.querystring("note0")
note1=request.querystring("note1")
note2=request.querystring("note2")
note3=request.querystring("note3")
note4=request.querystring("note4")
note5=request.querystring("note5")
note6=request.querystring("note6")
note7=request.querystring("note7")
note8=request.querystring("note8")
note9=request.querystring("note9")
note10=request.querystring("note10")
note11=request.querystring("note11")
note12=request.querystring("note12")
note13=request.querystring("note13")
note14=request.querystring("note14")
note15=request.querystring("note15")
note16=request.querystring("note16")
note17=request.querystring("note17")
dim sql1
sql1="insert into Finance (SalesID,ClientID,sum0,note0,sum1,note1,sum2,note2,sum3,note3,sum4,note4,sum5,note5,sum6,note6,sum7,note7,sum8,note8,sum9,note9,sum10,note10,sum11,note11,sum12,note12,sum13,note13,sum14,note14,sum15,note15,sum16,note16,sum17,note17) select 1,1," & sum0 & ",'" & note0 & "',"& sum1 & ",'" & note1 & "',"& sum2 & ",'" & note2 & "',"& sum3 & ",'" & note3 & "',"& sum4 & ",'" & note4 & "',"& sum5 & ",'" & note5 & "',"& sum6 & ",'" & note6 & "',"& sum7 & ",'" & note7 & "',"& sum8 & ",'" & note8 & "',"& sum9 & ",'" & note9 & "',"& sum10 & ",'" & note10 & "',"& sum11 & ",'" & note11 & "',"& sum12 & ",'" & note12 & "',"& sum13 & ",'" & note13 & "',"& sum14 & ",'" & note14 & "',"& sum15 & ",'" & note15 & "',"& sum16 & ",'" & note16 & "',"& sum17 & ",'" & note17 & "'" 
'response.write sql1
set rs = conn.execute(sql1) 
response.write err.description
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
Инициатор:<%=Session("Login")%>|Клиент:<%=Session("client")%><table width=80%><td>&nbsp;</td></table>
Дата:<%=date()%>|Контактное лицо, тел:<%=Session("contact")%><table width=80%><td>&nbsp;</td></table>
Срок ответа:<%=date()%>|Дата начала работ:<%=Session("begindate")%>
<br>
<br>
<br>
<br>

<form name=form2 method=get action=create_page2.asp>
<table border=1>
<tr>
<td>№</td><td>Статья затрат</td><td>USD, без НДС в месяц</td><td>Примечание (при наличии)</td>
</tr>
<tr>
<td>1.</td><td>ФОТ управленческого персонала (администаторы, мастеры, завхозы и т.д.)</td><td><%=request.querystring("sum0")%></td><td><%=request.querystring("note0")%></td>
</tr>
<tr>
<td>2.</td><td>ФОТ производственного персонала (уборщики, дворники и т.д.) (включая отпускные, оплату больничных листов, премии, надбавки, доплаты)</td><td><%=request.querystring("sum1")%></td><td><%=request.querystring("note1")%></td>
</tr>
<tr>
<td>3.</td><td>Отчисления с ФОТ управленческого и производственного персонала (ЕСН + НС)</td><td><%=request.querystring("sum2")%></td><td><%=request.querystring("note2")%></td>
</tr>
<tr>
<td>4.</td><td>Льготы персоналу, итого:</td><td><%=request.querystring("sum3")%></td><td><%=request.querystring("note3")%></td>
</tr>
<tr>
<td>4.1</td><td>медицинское обслуживание (добровольное страхование)</td><td><%=request.querystring("sum4")%></td><td><%=request.querystring("note4")%></td>
</tr>
<tr>
<td>4.2</td><td>материальная помощь персоналу</td><td><%=request.querystring("sum5")%></td><td><%=request.querystring("note5")%></td>
</tr>
<tr>
<td>4.3</td><td>оплата путевок персоналу в санатории, курорты и т.п.</td><td><%=request.querystring("sum6")%></td><td><%=request.querystring("note6")%></td>
</tr>
<tr>
<td>5.</td><td>Химические средства и расходные материалы для уборки</td><td><%=request.querystring("sum7")%></td><td><%=request.querystring("note7")%></td>
</tr>
<tr>
<td>6.</td><td>Расходные материалы гигиенические (бум. полотенца, туалетная бумага, ж/мыло и т.п.)</td><td><%=request.querystring("sum8")%></td><td><%=request.querystring("note8")%></td>
</tr>
<tr>
<td>7.</td><td>Затраты на спецодежду, СИЗ, и т.п.</td><td><%=request.querystring("sum9")%></td><td><%=request.querystring("note9")%></td>
</tr>
<tr>
<td>8.</td><td>ГСМ, запчасти к технике, ремонт</td><td><%=request.querystring("sum10")%></td><td><%=request.querystring("note10")%></td>
</tr>
<tr>
<td>9.</td><td>Затраты на уборочный инвентарь и т.п.</td><td><%=request.querystring("sum11")%></td><td><%=request.querystring("note11")%></td>
</tr>
<tr>
<td>10.</td><td>Амортизации основных средств, связанных с уборкой (техника, оборудование, недвижимость и т.п.)</td><td><%=request.querystring("sum12")%></td><td><%=request.querystring("note12")%></td>
</tr>
<tr>
<td>11.</td><td>Транспортные расходы (доставка персонала, ТМЦ и т.п.)</td><td><%=request.querystring("sum13")%></td><td><%=request.querystring("note13")%></td>
</tr>
<tr>
<td>12.</td><td>Вывоз мусора</td><td><%=request.querystring("sum14")%></td><td><%=request.querystring("note14")%></td>
</tr>
<tr>
<td>13.</td><td>Вывоз снега в зимний период</td><td><%=request.querystring("sum15")%></td><td><%=request.querystring("note15")%></td>
</tr>
<tr>
<td>14.</td><td>Другие затраты</td><td><%=request.querystring("sum16")%></td><td><%=request.querystring("note16")%></td>
</tr>
<tr>
<td>&nbsp;</td><td>ИТОГО текущие затраты:</td><td><%=request.querystring("sum17")%></td><td><%=request.querystring("note17")%></td>
</tr>


<tr>
<td><input type=submit value="Сохранить"></td>
</tr>
</table>
</form>
<br><br><br>
<br><br><br>
Согласовано:<br><br><br>
Генеральный директор _______________________________________/Московиц Д.С.
<br><br><br><br><br><br><br><br><br><br><br><br>
<table>
<tr><td>Tip-Top Cleaning</td><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td>Тип-Топ Клининг</td></tr>		
<tr><td>Ul.Ordzhonikidze, 11</td><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td>ул.Орджоникидзе,11</td></tr>		
<tr><td>115419 Moscow, Russia</td><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td>115419 Москва, Россия</td></tr>
<tr><td>www.tiptop.com.ru</td><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td>www.tiptop.com.ru</td></tr>
<tr><td>+7(095)234 45 20</td><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td>+7(095)234 45 20</td></tr>
</table>

</body>
</html>