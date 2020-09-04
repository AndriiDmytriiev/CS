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
dim b1,b2,b3,b4,b5,b6,b7,b8,b9,b10,b11,b12,b13,b14,b15,b16,b17,b18,b19,b20,b21,b22,b23,b24,b25,b26,b27,b28,b29,b30,b31,b32,b33
b1=request.querystring("b1")
b2=request.querystring("b2")
b3=request.querystring("b3")
b4=request.querystring("b4")
b5=request.querystring("b5")
b6=request.querystring("b6")
b7=request.querystring("b7")
b8=request.querystring("b8")
b9=request.querystring("b9")
b10=request.querystring("b10")
b11=request.querystring("b11")
b12=request.querystring("b12")
b13=request.querystring("b13")
b14=request.querystring("b14")
b15=request.querystring("b15")
b16=request.querystring("b16")
b17=request.querystring("b17")
b18=request.querystring("b18")
b19=request.querystring("b19")
b20=request.querystring("b20")
b21=request.querystring("b21")
b22=request.querystring("b22")
b23=request.querystring("b23")
b24=request.querystring("b24")
b25=request.querystring("b25")
b26=request.querystring("b26")
b27=request.querystring("b27")
b28=request.querystring("b28")
b29=request.querystring("b29")
b30=request.querystring("b30")
b31=request.querystring("b31")
b32=request.querystring("b32")
b33=request.querystring("b33")


dim sql1
sql1="insert into InternClean (SaleID,ClientID,b1,b2,b3,b4,b5,b6,b7,b8,b9,b10,b11,b12,b13,b14,b15,b16,b17,b18,b19,b20,b21,b22,b23,b24,b25,b26,b27,b28,b29,b30,b31,b32,b33) select 1,1,'" & b1 & "','" & b2 & "','" & b3 & "','" & b4 & "','" & b5 & "','" & b6 & "','"& b7 & "','" & b8 & "','" & b9 & "','" & b10 & "','" & b11 & "','" & b12 & "','"& b13 & "','" & b14 & "','" & b15 & "','" & b16 & "','" & b17 & "','" & b18 & "','" & b19 & "','" & b20 & "','" & b21 & "','" & b22 & "','" & b23 & "','" & b24 & "','" & b25 & "','" & b26 & "','" & b27 & "','" & b28 & "','" & b29 & "','" & b30 & "','" & b31 & "','" & b32 & "','" & b33 & "'" 
'response.write sql1
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
Инициатор:<%=Session("Login")%>|Клиент:<%=Session("client")%><table width=80%><td>&nbsp;</td></table>
Дата:<%=date()%>|Контактное лицо, тел:<%=Session("contact")%><table width=80%><td>&nbsp;</td></table>
Срок ответа:<%=date()%>|Дата начала работ:<%=Session("begindate")%>
<br>

<form method=get name=form2 action=create_page3.asp>
ВНИМАНИЕ: Заполнению подлежат ячейки, выделенные цветом. Данные, полученные в ходе исследования строго конфиденциальны!

2.<b><u>УБОРКА ТЕРРИТОРИИ:</u></b>
Предоставьте, пожалуйста, информацию по всем зданиям, в которых необходимо осуществить уборку.
<br>
<table border=1>
<tr>
<td>&nbsp;</td><td>Показатель</td><td>Ед.изм.</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>1.</td><td><b>Общая убираемая территория</b></td><td>кв.м</td><td><%=request.querystring("b1")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% if request.querystring("ch1") then%>
<tr>
<td>1.1</td><td>Площади с указанием покрытия (тротуары, стоянки, газоны, проезды, другое)</td><td>&nbsp;</td><td><%=request.querystring("b2")%></td><td><input type=checkbox name = ch1></td><td>
<%=request.querystring("b3")%>
</td>
</tr>
<% if request.querystring("ch2") then%>
<tr>
<td>1.1.1</td><td>Тротуары</td><td>год</td><td><%=request.querystring("b4")%></td><td><input type=checkbox name = ch2></td><td>
<%=request.querystring("b5")%>
</td>
</tr>
<% end if %>
<% if request.querystring("ch3") then%>
<tr>
<td>1.1.2</td><td>Стоянки</td><td>м.</td><td><%=request.querystring("b6")%></td><td><input type=checkbox name = ch3></td><td>
<%=request.querystring("b7")%>
</td>
</tr>
<% end if %>
<% if request.querystring("ch4") then%>
<tr>
<td>1.1.3</td><td>Газоны</td><td>м.</td><td><%=request.querystring("b8")%></td><td><input type=checkbox name = ch4></td><td>
<%=request.querystring("b9")%>
</td>
</tr>
<% end if %>
<% if request.querystring("ch5") then%>
<tr>
<td>1.1.4</td><td>Проезды</td><td>м.</td><td><%=request.querystring("b10")%></td><td><input type=checkbox name = ch5></td><td>
<%=request.querystring("b11")%>
</td>
</tr>
<% end if %>
<% end if %>

<% if request.querystring("ch6")="ON" then%>
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

<% end if %>

<% if request.querystring("ch7")="ON" then%>
<tr>
<td>3.</td><td><b>Инвентаризационная ведомость техники (оборудования) на</b></td><td>чел.</td><td><%=request.querystring("b18")%></td><td><input type=checkbox name = ch7></td><td>&nbsp;</td>
</tr>
<tr>
<td>3.1</td><td>Офисные помещения</td><td>мес</td><td><%=request.querystring("b19")%></td><td><input type=checkbox name = ch7></td><td>
<%=request.querystring("b20")%>
</td>
</tr>
<tr>
<td>3.2</td><td>Офисные помещения</td><td>мес</td><td><%=request.querystring("b21")%></td><td><input type=checkbox name = ch7></td><td>
<%=request.querystring("b22")%>
</td>
</tr>
<tr>
<td>3.3</td><td>Офисные помещения</td><td>мес</td><td><%=request.querystring("b23")%></td><td><input type=checkbox name = ch7></td><td>
<%=request.querystring("b24")%>
</td>
</tr>
<tr>
<td>3.4</td><td>Офисные помещения</td><td>мес</td><td><%=request.querystring("b25")%></td><td><input type=checkbox name = ch7></td><td>
<%=request.querystring("b26")%>
</td>
</tr>
<% end if %>

<% if request.querystring("ch8")="ON" then%>
<tr>
<td>4.</td><td><b>Наличие места для стоянки уборочной техники,</b> оборудованного отоплением и складом для хранения расходных материалов и запчастей, а также проведения ремонта техники, оборудования и инвентаря</td><td>да/нет</td><td><%=request.querystring("b27")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% end if %>
<% if request.querystring("ch9")="ON" then%>
<tr>
<td>5.</td><td><b>Потребность в стрижке газона, покосе травы (площадь и кол-во раз)</b></td><td>да/нет</td><td><%=request.querystring("b28")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>5.1</td><td>площадь</td><td>кв.м</td><td><%=request.querystring("b29")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td></td><td>количество стрижек газонов и покосов травы</td><td>кол-во</td><td><%=request.querystring("b30")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% end if %>
<% if request.querystring("ch10")="ON" then%>
<tr>
<td>6.</td><td><b>Качество уборки территории зимой </b>(A - до утрамбованного слоя снега/ Б - до покрытия)</td><td>А/Б</td><td><%=request.querystring("b31")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if request.querystring("ch11")="ON" then%>
<tr>
<td>7.</td><td><b>Вывоз снега в зимний период (объем вывоза за зиму)</b></td><td>куб.м.</td><td><%=request.querystring("b32")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% end if %>

<% if request.querystring("ch12")="ON" then%>
<tr>
<td>8.</td><td><b>Вывоз ТБО, примерный месячный объем мусора</b></td><td>куб.м.</td><td><%=request.querystring("b33")%></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<% end if %>



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