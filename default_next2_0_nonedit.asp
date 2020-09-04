<%
on error resume next
set conn2 = server.CreateObject("Adodb.connection")
      
        connStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & _
         Server.MapPath("missyou.mdb")
         conn2.Open connStr
      ' Recordset object
      
        
        
      
'-Сохраняем данные из второй формы---------
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
<td>Инициатор:<%=Session("Login")%></td><td>|</td><td>Клиент:<%=Session("client")%></td>
</tr>
<tr>
<td>Дата:<%=date()%></td><td>|</td><td>Контактное лицо, тел:<%=Session("contact")%></td>
</tr>
<tr>
<td>Срок ответа:<%=Session("day") & "." & Session("month") & "." & Session("year") %></td><td>|</td><td>Дата начала работ:<%=date()%></td>
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
<tr><td align=left>ВНИМАНИЕ: Заполнению подлежат ячейки, выделенные цветом. Данные, полученные в ходе исследования строго конфиденциальны!
</td></tr>
<tr><td align=left>1.<b><u>ВНУТРЕННЯЯ УБОРКА ПОМЕЩЕНИЙ:</u></b></u></b></td></tr>
</table>
<br>


Предоставьте, пожалуйста, информацию по всем зданиям, в которых необходимо осуществить уборку.



<table border=1>
<tr>
<td>&nbsp;</td><td>Показатель</td><td>Ед.изм.</td><td>Здание № 1</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>1.</td><td><b>Краткие данные о здании/объекте</b></td><td></td><td>офисы</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>1.1</td><td>этажность</td><td>эт.</td><td>
<input type=textbox name=t1 size="20"></td><td>
<input type=checkbox name = ch1 ></td><td>
<select name="Period1">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>1.2</td><td>год последнего ремонта</td><td>год</td><td>
<input type=textbox name=t2 size="20"></td><td>
<input type=checkbox name = ch2 ></td><td>
<select name="Period2">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>1.3</td><td>высота потолков</td><td>м.</td><td>
<input type=textbox name=t3 size="20"></td><td>
<input type=checkbox name = ch3 ></td><td>
<select name="Period3">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>2.</td><td><b>Общая площадь</b></td><td>кв.м</td><td>
<input type=textbox name=t4 size="20"></td><td>
<input type=checkbox name = ch4 ></td><td>
<select name="Period4">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>2.1</td><td>Площадь каждого этажа</td><td>кв.м</td><td>
<input type=textbox name=t5 size="20"></td><td>
<input type=checkbox name = ch5 ></td><td>
<select name="Period5">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>2.2</td><td>Кабинеты VIP</td><td>кол-во</td><td>
<input type=textbox name=t6_1 size="20"></td><td>
<input type=checkbox name = ch6 ></td><td>
<select name="Period6">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t6_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.3</td><td>Офисные помещения</td><td>кол-во</td><td>
<input type=textbox name=t7_1 size="20"></td><td>
<input type=checkbox name = ch7 ></td><td>
<select name="Period7">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t7_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.4</td><td>Складские помещения</td><td>кол-во</td><td>
<input type=textbox name=t8_1 size="20"></td><td>
<input type=checkbox name = ch8 ></td><td>
<select name="Period8">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t8_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.5</td><td>Складские помещения</td><td>кол-во</td><td>
<input type=textbox name=t9_1 size="20"></td><td>
<input type=checkbox name = ch9 ></td><td>
<select name="Period9">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t9_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.6</td><td>Технические помещения, подвалы</td><td>кол-во</td><td>
<input type=textbox name=t10_1 size="20"></td><td>
<input type=checkbox name = ch10 ></td><td>
<select name="Period10">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t10_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>


<tr>
<td>2.7</td><td>Коридоры</td><td>кол-во</td><td>
<input type=textbox name=t11_1 size="20"></td><td>
<input type=checkbox name = ch11 ></td><td>
<select name="Period11">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t11_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.8</td><td>Лестницы</td><td>кол-во</td><td>
<input type=textbox name=t12_1 size="20"></td><td>
<input type=checkbox name = ch12 ></td><td>
<select name="Period12">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t12_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.9</td><td>Лифты</td><td>кол-во</td><td>
<input type=textbox name=t13_1 size="20"></td><td>
<input type=checkbox name = ch13 ></td><td>
<select name="Period13">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t13_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.10</td><td>Эскалаторы</td><td>кол-во</td><td>
<input type=textbox name=t14_1 size="20"></td><td>
<input type=checkbox name = ch14 ></td><td>
<select name="Period14">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t14_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.11</td><td>Санузлы</td><td>кол-во</td><td>
<input type=textbox name=t15_1 size="20"></td><td>
<input type=checkbox name = ch15 ></td><td>
<select name="Period15">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t15_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.12</td><td>холлы, вестибюли</td><td>кол-во</td><td>
<input type=textbox name=t16_1 size="20"></td><td>
<input type=checkbox name = ch16 ></td><td>
<select name="Period16">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t16_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.13</td><td>Другие площади (по возм.расшифровать)</td><td></td><td>
<input type=textbox name=t17_0 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>кол-во</td><td>
<input type=textbox name=t17_1 size="20"></td><td>
<input type=checkbox name = ch17 ></td><td>
<select name="Period17">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t17_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>


<tr>
<td>3.</td><td><b>К-во сотрудников Вашей компании и посетителей, чел. в среднем в день</b></td><td>чел.</td><td>
<input type=textbox name=t18 size="20"></td><td>
<input type=checkbox name = ch18 ></td><td>&nbsp;</td>
</tr>


<tr>
<td>4.</td><td><b>Необходимость в обеспечении с/у расходными материалами (примерный расход в месяц)</b></td><td>кол-во</td><td>&nbsp;</td><td>&nbsp;</td><td>
<select name="Period19">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>туалетная бумага</td><td>рул./мес.</td><td>
<input type=textbox name=t19_1 size="20"></td><td>
<input type=checkbox name = ch19_1 ></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>жидкое мыло</td><td>литр/мес.</td><td>
<input type=textbox name=t19_2 size="20"></td><td>
<input type=checkbox name = ch19_2 ></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>бум полотенца</td><td>лист/мес.</td><td>
<input type=textbox name=t19_3 size="20"></td><td>
<input type=checkbox name = ch19_3 ></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>бумажные сидения д/унитаза</td><td>шт./мес.</td><td>
<input type=textbox name=t19_4 size="20"></td><td>
<input type=checkbox name = ch19_4 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.</td><td><b>Поверхности</b></td><td>чел.</td><td>&nbsp;</td><td>
<input type=checkbox name = ch20 ></td><td>
<select name="Period20">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>

<tr>
<td>5.1</td><td>Мягкие покрытия (ковролин)</td><td>кв.м</td><td>
<input type=textbox name=t20_1 size="20"></td><td>
<input type=checkbox name = ch20_1 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.2</td><td>Полутвердые покрытия (линолеум, паркет, мармолеум, наливной пол, ламинат)</td><td>кв.м</td><td>
<input type=textbox name=t20_2 size="20"></td><td>
<input type=checkbox name = ch20_2 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.3</td><td>Твердые покрытия (плитка, мрамор, гранит) пол/стены</td><td>кв.м</td><td>
<input type=textbox name=t20_3 size="20"></td><td>
<input type=checkbox name = ch20_3 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.4</td><td>Стеклянные поверхности</td><td>кв.м</td><td>
<input type=textbox name=t20_4 size="20"></td><td>
<input type=checkbox name = ch20_4 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.5</td><td>Металлические поверхности</td><td>кв.м</td><td>
<input type=textbox name=t20_5 size="20"></td><td>
<input type=checkbox name = ch20_5 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.6</td><td>Офисные перегородки</td><td>шт.</td><td>
<input type=textbox name=t20_6 size="20"></td><td>
<input type=checkbox name = ch20_6 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.7</td><td>Офисные места</td><td>шт.</td><td>
<input type=textbox name=t20_7 size="20"></td><td>
<input type=checkbox name = ch20_7 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.8</td><td>Кожаная мебель</td><td>шт.</td><td>
<input type=textbox name=t20_8 size="20"></td><td>
<input type=checkbox name = ch20_8 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.9</td><td>Пластиковая мебель</td><td>шт.</td><td>
<input type=textbox name=t20_9 size="20"></td><td>
<input type=checkbox name = ch20_9 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.10</td><td>Деревянная мебель</td><td>шт.</td><td>
<input type=textbox name=t20_10 size="20"></td><td>
<input type=checkbox name = ch20_10 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.11</td><td>Другие поверхности</td><td>шт./кв.м</td><td>
<input type=textbox name=t20_11 size="20"></td><td>
<input type=checkbox name = ch20_11 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>6.</td><td><b>Существующая у Вас уборка</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
<input type=checkbox name = ch21 ></td><td>
<select name="Period21">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>


<tr>
<td>6.1</td><td>График проведения основной комплексной уборки</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>с:</td><td>час</td><td>
<input type=textbox name = t22 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>до:</td><td>час</td><td>
<input type=textbox name = t23 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>дней в году:</td><td>дни</td><td>
<input type=textbox name = t24 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>6.2</td><td>График проведения поддерживающей уборки</td><td></td><td></td><td></td><td></td>
</tr>

<tr>
<td>&nbsp;</td><td>с:</td><td>час</td><td>
<input type=textbox name = t25 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>до:</td><td>час</td><td>
<input type=textbox name = t26 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>дней в году:</td><td>дни</td><td>
<input type=textbox name = t27 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>



<tr>
<td>6.3</td><td>Количество уборщиков</td><td>чел.</td><td>
<input type=textbox name = t28 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.</td><td><b>Помещения для размещения производственного персонала и оборудования</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
<input type=checkbox name = ch22 ></td><td>
<select name="Period22">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>

<tr>
<td>7.1</td><td>Наличие помещения</td><td>да/нет</td><td><select name="choice1">
<option value="да" >да</option>
<option value="нет" >нет</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.2</td><td>Наличие телефона в помещении</td><td>да/нет</td><td><select name="choice2">
<option value="да" >да</option>
<option value="нет" >нет</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.3</td><td>площадь помещения</td><td>кв.м</td><td>
<input type=textbox name = t29 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>8.</td><td><b>Помещения для размещения управленческого персонала и автоматизированных рабочих мест (для объектов более 10000 кв.м)</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
<input type=checkbox name = ch23 ></td><td>
<select name="Period23">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>

<tr>
<td>8.1</td><td>Наличие помещения</td><td>да/нет</td><td><select name="choice3">
<option value="да" >да</option>
<option value="нет" >нет</option>
</select></td><td></td><td></td>
</tr>
<tr>
<td>8.2</td><td>Площадь помещения</td><td>кв.м</td><td>
<input type=textbox name = t30 size="20"></td><td></td><td></td>
</tr>
<tr>
<td>8.3</td><td>кол-во автоматизированных рабочих мест</td><td>шт.</td><td><input type=textbox name = t31 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>



</table>

<br>

<table border=0 width=100%>
<tr><td align=left>ВНИМАНИЕ: Заполнению подлежат ячейки, выделенные цветом. Данные, полученные в ходе исследования строго конфиденциальны!
</td></tr>
<tr><td align=left>2.<b><u>УБОРКА ТЕРРИТОРИИ:</u></b></td></tr>
</table>


<br>

Предоставьте, пожалуйста, информацию по всем зданиям, в которых необходимо осуществить уборку.
<br>
<table border=1>
<tr>
<td>&nbsp;</td><td>Показатель</td><td>Ед.изм.</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>1.</td><td><b>Общая убираемая территория</b></td><td>кв.м</td><td><input type=textbox name=b1></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>1.1</td><td>Площади с указанием покрытия (тротуары, стоянки, газоны, проезды, другое)</td><td>&nbsp;</td><td><input type=textbox name=b2></td><td><input type=checkbox name = k1></td><td>
<select name="b3">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год " >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>1.1.1</td><td>Тротуары</td><td>год</td><td><input type=textbox name=b4></td><td><input type=checkbox name = k2></td><td>
<select name="b5">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год " >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>1.1.2</td><td>Стоянки</td><td>м.</td><td><input type=textbox name=b6></td><td><input type=checkbox name = k3></td><td>
<select name="b7">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год " >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>1.1.3</td><td>Газоны</td><td>м.</td><td><input type=textbox name=b8></td><td><input type=checkbox name = k4></td><td>
<select name="b9">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год " >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>1.1.4</td><td>Проезды</td><td>м.</td><td><input type=textbox name=b10></td><td><input type=checkbox name = k5></td><td>
<select name="b11">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год " >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>2.</td><td><b>Существующая у Вас уборка территории</b></td><td>кв.м</td><td><input type=textbox name=b12></td><td><input type=checkbox name = k6></td><td>
<select name="b13">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год " >раз в год</option>
</select>
</td>
</tr>

<tr>
<td>2.1</td><td>График уборки</td><td></td><td></td><td></td><td></td>
</tr>

<tr>
<td>&nbsp;</td><td>с:</td><td>час</td><td><input type=textbox name = b14></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>до:</td><td>час</td><td><input type=textbox name = b15></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>дней в году:</td><td>дни</td><td><input type=textbox name = b16></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>


<tr>
<td>2.2</td><td>Количество уборщиков</td><td>чел.</td><td><input type=textbox name = b17></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>


<tr>
<td>3.</td><td><b>Инвентаризационная ведомость техники (оборудования) на</b></td><td>чел.</td><td><input type=textbox name=b18></td><td><input type=checkbox name = k7></td><td>&nbsp;</td>
</tr>
<tr>
<td>3.1</td><td>Офисные помещения</td><td>мес</td><td><input type=textbox name=b19></td><td><input type=checkbox name = k8></td><td>
<select name="b20">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год " >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>3.2</td><td>Офисные помещения</td><td>мес</td><td><input type=textbox name=b21></td><td><input type=checkbox name = k9></td><td>
<select name="b22">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год " >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>3.3</td><td>Офисные помещения</td><td>мес</td><td><input type=textbox name=b23></td><td><input type=checkbox name = k10></td><td>
<select name="b24">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год " >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>3.4</td><td>Офисные помещения</td><td>мес</td><td><input type=textbox name=b25></td><td><input type=checkbox name = k11></td><td>
<select name="b26">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год " >раз в год</option>
</select>
</td>
</tr>

<tr>
<td>4.</td><td><b>Наличие места для стоянки уборочной техники,</b> оборудованного отоплением и складом для хранения расходных материалов и запчастей, а также проведения ремонта техники, оборудования и инвентаря</td><td>да/нет</td><td><select name="b27">
<option value="да" >да</option>
<option value="нет" >нет</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>5.</td><td><b>Потребность в стрижке газона, покосе травы (площадь и кол-во раз)</b></td><td>да/нет</td><td><select name="b28">
<option value="да" >да</option>
<option value="нет" >нет</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>5.1</td><td>площадь</td><td>кв.м</td><td><input type=textbox name=b29></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td></td><td>количество стрижек газонов и покосов травы</td><td>кол-во</td><td><input type=textbox name=b30></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>6.</td><td><b>Качество уборки территории зимой </b>(A - до утрамбованного слоя снега/ Б - до покрытия)</td><td>А/Б</td><td><select name="b31">
<option value="1" selected>А</option>
<option value="2" >Б</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.</td><td><b>Вывоз снега в зимний период (объем вывоза за зиму)</b></td><td>куб.м.</td><td><input type=textbox name=b32></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>

<tr>
<td>8.</td><td><b>Вывоз ТБО, примерный месячный объем мусора</b></td><td>куб.м.</td><td><input type=textbox name=b33></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

</table>
--------->

<align=left>
<br>
<table border=0 width=100%>
<tr><td align=left>ВНИМАНИЕ: Заполнению подлежат ячейки, выделенные цветом. Данные, полученные в ходе исследования строго конфиденциальны!</td></tr>
<tr><td align=left>3.<b><u>ДОПОЛНИТЕЛЬНЫЕ УСЛУГИ:</u></b></td></tr>
</table>
<table border=1 width=100%>
<tr>
<td>&nbsp;</td><td>Показатель</td><td>Ед.изм.</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td><b>1.</b></td><td><b>Организация работы прачечных</b></td><td>да/нет</td><td><select name="c1">
<option value="<%=session("c1")%>" selected><%=session("c1")%></option>
<option value="да" >да</option>
<option value="нет" >нет</option>
</select></td><td><input type=checkbox name = c2 value='<%=session("c2")%>'></td>
</tr>
<tr>
<td>1.1</td><td>Наличие химчистки-прачечной</td><td>да/нет</td><td><select name="c3">
<option value="<%=session("c3")%>" selected><%=session("c3")%></option>
<option value="да" >да</option>
<option value="нет" >нет</option>
</select></td><td><input type=checkbox name = c4 value='<%=session("c4")%>'></td>
</tr>
<tr>
<td>1.1.1</td><td>количество обслуживающего персонала</td><td>чел.</td><td><input type=text name = c5 value='<%=session("c5")%>'></td><td><input type=checkbox name = c6 value='<%=session("c6")%>'></td>
</tr>
<tr>
<td>1.2</td><td>Потребность в стирке спецодежды, объем в месяц</td><td>кг</td><td><input type=text name = c7 value='<%=session("c7")%>'></td><td><input type=checkbox name = c8 value='<%=session("c8")%>'></td>
</tr>
<tr>
<td><b>2.</b></td><td><b>Глубокая химическая чистка ковровых покрытий</b></td><td>да/нет</td><td><select name="c9">
<option value="<%=session("c9")%>" selected><%=session("c9")%></option>
<option value="да" selected>да</option>
<option value="нет" >нет</option>
</select></td><td><input type=checkbox name = c10 value='<%=session("c10")%>'></td>
</tr>
<tr>
<td>2.1</td><td>Площадь ковровых покрытий</td><td>кв.м</td><td><input type=text name = c11 value='<%=session("c11")%>'></td><td><input type=checkbox name = c12 value='<%=session("c12")%>'></td>
</tr>
<tr>
<td>2.2</td><td>Периодичность проведения</td><td>&nbsp;</td><td><input type=text name = c13 value='<%=session("c13")%>'></td> <td>&nbsp;</td>
</tr>
<tr>
<td><b>3.</b></td><td><b>Нанесение полимерного лака на линолеумные полы</b></td><td>да/нет</td><td><select name="c14">
<option value="<%=session("c14")%>" selected><%=session("c14")%></option>
<option value="да" selected>да</option>
<option value="нет" >нет</option>
</select></td><td><input type=checkbox name = c15 value='<%=session("c15")%>'></td>
</tr>
<tr>
<td>3.1</td><td>Площадь линолеума</td> <td>кв.м</td><td><input type=text name = c16 value='<%=session("c16")%>'></td>
</tr>
<tr>
<td>3.2</td><td>Периодичность нанесения лака</td> <td>&nbsp;</td><td><input type=text name = c17 value='<%=session("c17")%>'></td><td>&nbsp;</td>
</tr>
<tr>
<td><b>4.</b></td><td><b>Мытье окон (площадь окон с одной стороны)</b></td> <td>да/нет</td><td><select name="c18">
<option value="<%=session("c18")%>" selected><%=session("c18")%></option>
<option value="да" selected>да</option>
<option value="нет" >нет</option>
</select></td><td><input type=checkbox name = c19 value='<%=session("c19")%>'></td>
</tr>
<tr>
<td>4.1</td><td>Легкий доступ к окнам (с пола)</td> <td>кв.м</td><td><input type=text name = c20 value='<%=session("c20")%>'></td><td><input type=checkbox name = c21 value='<%=session("c21")%>'></td>
</tr>
<tr>
<td>4.2</td><td>Затрудненный (со стремянки)</td> <td>кв.м</td><td><input type=text name = c22 value='<%=session("c22")%>'></td><td><input type=checkbox name = c23 value='<%=session("c23")%>'></td>
</tr>
<tr>
<td>4.3</td><td> С помощью промышленных альпинистов </td> <td>кв.м</td><td><input type=text name = c24 value='<%=session("c24")%>'></td><td><input type=checkbox name = c25 value='<%=session("c25")%>'></td>
</tr>
<tr>
<td><b>5.</b></td><td><b>Другие услуги, которые вы хотели бы получать</b></td><td>&nbsp;</td> <td>&nbsp;</td><td><input type=checkbox name = c26 value='<%=session("c26")%>'></td>
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
<td></td><td></td><td>&nbsp;</td><td>&nbsp;</td><td><input type=submit value="Сохранить >>"></td>
</tr>
</table>
</align>
</form>
<center>
<table border=0 >
<td>
<tr>		
		<td>
		<a href="search.asp"><font  size=3>Найти заявки от клиентов</font></a>
		<hr>
		</td>
		<td>
		<a href="finansi.asp"><font  size=3>Финансы</font></a>
		<hr>
		</td>
		<td>
		<a href="history.asp"><font  size=3>Журнал</font></a>
		<hr>
		</td>
		<td>
		<a href="reg.asp"><font  size=3>Зарегистрироваться</font></a>
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